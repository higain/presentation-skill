"""
Native presentation editing through LibreOffice's UNO API (python3-uno).

LibreOffice opens the real document and saves it back in its own format, so an
.odp stays ODF and a .pptx stays OOXML -- no conversion, nothing dropped.
Masters, layouts, speaker notes, animations, transitions, charts and embedded
media survive an edit here that a rewrite-from-scratch library would lose.

    import sys; sys.path.insert(0, "<skill>/scripts")
    from office.impress import impress_document, slides, set_notes

    with impress_document("deck.odp") as doc:
        page = slides(doc)[0]
        page.getByIndex(0).setString("Q3 Ergebnisse")
        set_notes(page, "Mit der Umsatzzahl aufmachen.")
    # saved in place, still .odp

All geometry is in **1/100 mm** (UNO's drawing unit), not points or EMU.
``mm(15)`` and ``pt(18)`` convert.

**Read the slide size, never assume it.**  ``slide_size(doc)`` is the only
answer that is right for the deck in front of you: a LibreOffice default page
is 28000 x 15750, but a brand template is usually its own size -- one is
32001 x 18000.  Hardcoding the default silently shifts every shape you place
on a template that is 14% wider.

One headless LibreOffice is started on first use and shut down when the process
ends, so many documents in one script cost one startup.
"""

import atexit
import contextlib
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

from office.soffice import SOFFICE_MISSING, free_port, get_soffice_env, profile_arg, soffice_path

# Store filters by output extension.  storeToURL() needs the exact name.
FILTERS = {
    ".odp": "impress8",
    ".otp": "impress8_template",
    ".pptx": "Impress MS PowerPoint 2007 XML",
    ".potx": "Impress MS PowerPoint 2007 XML Template",
    ".ppt": "MS PowerPoint 97",
    ".pdf": "impress_pdf_Export",
}

# Extensions this skill will open.  .ppt is read-only in practice: convert it.
PRESENTATION_EXTS = {".odp", ".otp", ".fodp", ".pptx", ".potx", ".ppt", ".pps", ".ppsx"}

# OOXML packages -- the ones repair_pptx() knows how to mend after a save.
_OOXML_EXTS = {".pptx", ".potx"}

PRESENTATION_SERVICE = "com.sun.star.presentation.PresentationDocument"


class SofficeError(RuntimeError):
    pass


#: No slide is a metre across -- a 16:9 brand template is about 320 mm.  A
#: value this large reached ``mm()`` already in 1/100 mm and is being converted
#: twice: ``mm(26000)`` is 26 metres.  Measured on a real build, that is where
#: "extends 2471929 past the right edge" comes from, three steps downstream of
#: the mistake and reported against a shape rather than against the number.
MAX_MM = 1000


def mm(value: float) -> int:
    """Millimetres to UNO's 1/100 mm."""
    if abs(value) > MAX_MM:
        raise ValueError(
            f"mm({value:g}) is {value / 1000:g} metres.  mm() takes "
            f"MILLIMETRES -- a slide is about 320 mm wide, so nothing on one "
            f"is over {MAX_MM} mm.\n"
            f"  Geometry you read back -- from probe.py, slide_size(doc), "
            f"canvas() or a Box -- is ALREADY in 1/100 mm.  Pass it straight "
            f"through:\n"
            f"      D.rect(doc, page, area.top(mm(20)), ...)   # not mm(area.y)\n"
            f"      D.Box(area.x, area.y + mm(10), area.w, mm(20))\n"
            f"  If {value:g} really is a 1/100 mm figure, drop the mm(): it is "
            f"{value / 100:g} mm."
        )
    return int(round(value * 100))


def pt(value: float) -> int:
    """Points to UNO's 1/100 mm (1 pt = 1/72 inch)."""
    return int(round(value * 2540 / 72))


def parse_slide_spec(spec: str) -> list[int]:
    """``'3,7-9'`` -> ``[3, 7, 8, 9]``.  The ``--slides`` argument, once.

    Lives here because every entry point that takes ``--slides`` -- render,
    qa, probe, make_deck -- had grown its own copy, and they disagreed about
    whether a reversed range was an error.
    """
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, _, hi = part.partition("-")
            try:
                lo, hi = int(lo), int(hi)
            except ValueError:
                raise ValueError(f"not a slide range: {part!r}") from None
            if lo > hi:
                raise ValueError(f"reversed slide range: {part!r}")
            out.update(range(lo, hi + 1))
        else:
            try:
                out.add(int(part))
            except ValueError:
                raise ValueError(f"not a slide number: {part!r}") from None
    if not out or min(out) < 1:
        raise ValueError("slide numbers start at 1")
    return sorted(out)


class _Server:
    """A headless soffice listening on a private loopback port."""

    def __init__(self):
        self.proc = None
        self.ctx = None
        self.profile = None

    def desktop(self, timeout: float = 60.0):
        if self.ctx is not None and not self._alive():
            # LibreOffice died under us -- a crash on a malformed shape, an OOM,
            # an outside kill.  Without this every later call in the process
            # fails with "Binary URP bridge disposed" or "cannot get value
            # ServiceManager", which reads as a bug in the caller's script
            # rather than a dead server.  Drop the corpse and start again.
            self.stop()
        if self.ctx is None:
            self._start(timeout)
        try:
            return self.ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
        except Exception:
            self.stop()                  # bridge disposed between check and use
            self._start(timeout)
            return self.ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )

    def _alive(self) -> bool:
        return self.proc is not None and self.proc.poll() is None

    def _start(self, timeout: float):
        port = free_port()
        self.profile = tempfile.mkdtemp(prefix="lo-uno-profile-")
        try:
            exe = soffice_path()
        except FileNotFoundError:
            raise SofficeError(SOFFICE_MISSING) from None

        self.proc = subprocess.Popen(
            [
                exe,
                "--headless",
                "--invisible",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                "--nologo",
                profile_arg(self.profile),
                f"--accept=socket,host=127.0.0.1,port={port};urp;",
            ],
            env=get_soffice_env(),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        atexit.register(self.stop)

        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        deadline = time.monotonic() + timeout
        while True:
            if self.proc.poll() is not None:
                self.stop()
                raise SofficeError(
                    f"LibreOffice exited with code {self.proc.returncode} before accepting "
                    "a connection"
                )
            try:
                self.ctx = resolver.resolve(url)
                return
            except Exception:  # NoConnectException until the listener is up
                if time.monotonic() > deadline:
                    self.stop()
                    raise SofficeError(
                        f"LibreOffice did not accept a UNO connection within {timeout:g}s"
                    ) from None
                time.sleep(0.25)

    def stop(self):
        if self.ctx is not None:
            with contextlib.suppress(Exception):
                self.ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.ctx
                ).terminate()
            self.ctx = None
        if self.proc is not None:
            try:
                self.proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                self.proc.kill()
                self.proc.wait(timeout=10)
            self.proc = None
        if self.profile:
            shutil.rmtree(self.profile, ignore_errors=True)
            self.profile = None


_SERVER = _Server()


def shutdown():
    """Stop the shared LibreOffice.  Called automatically at process exit."""
    _SERVER.stop()


def kill():
    """Hard-kill the shared LibreOffice, for when it stops responding."""
    proc, _SERVER.proc, _SERVER.ctx = _SERVER.proc, None, None
    if proc is not None and proc.poll() is None:
        proc.kill()


def _props(**kwargs):
    out = []
    for name, value in kwargs.items():
        p = PropertyValue()
        p.Name = name
        p.Value = value
        out.append(p)
    return tuple(out)


def filter_for(path) -> str:
    """The storeToURL filter name for an output path's extension."""
    ext = Path(path).suffix.lower()
    try:
        return FILTERS[ext]
    except KeyError:
        raise SofficeError(
            f"no LibreOffice filter for '{ext}'; supported: {', '.join(sorted(FILTERS))}"
        ) from None


def load_document(path, read_only: bool = False, timeout: float = 60.0):
    """Open a presentation and return the UNO document.  Caller must close it."""
    path = Path(path).absolute()
    if not path.exists():
        raise FileNotFoundError(f"{path} does not exist")
    if not os.access(path, os.R_OK):
        raise SofficeError(f"{path} is not readable")
    if not read_only and not os.access(path, os.W_OK):
        raise SofficeError(
            f"{path} is not writable, so it cannot be opened for editing. "
            "Copy it first, or open it with read_only=True."
        )

    options = {
        "Hidden": True,
        "MacroExecutionMode": 1,  # NEVER_EXECUTE
        # False opens a template (.otp/.potx) as itself, so store() has somewhere
        # to write.  Omit it and LibreOffice hands back a new untitled document
        # based on the template instead, whose getLocation() is empty.
        "AsTemplate": False,
        # 0 = NO_UPDATE: keep cached values of links to other files instead of
        # trying to resolve files that are probably absent.
        "UpdateDocMode": 0,
    }
    # Pass ReadOnly only to turn it ON.  Explicitly asking for ReadOnly=False on
    # a file the process cannot write makes type detection fail outright, and
    # LibreOffice reports it as the useless "Unsupported URL ... type detection
    # failed" rather than a permission error.
    if read_only:
        options["ReadOnly"] = True

    desktop = _SERVER.desktop(timeout=timeout)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(path)), "_blank", 0, _props(**options)
    )
    if doc is None:
        raise SofficeError(f"LibreOffice could not open {path} as a document")
    if not doc.supportsService(PRESENTATION_SERVICE):
        doc.close(False)
        raise SofficeError(f"{path} is not a presentation")
    return doc


def new_document(timeout: float = 60.0):
    """A blank Impress document, not yet on disk.  Save it with save_as().

    **It already contains one slide.**  Use ``slides(doc)[0]`` for your title
    slide; calling add_slide() for it leaves a stray blank slide at the front.
    """
    desktop = _SERVER.desktop(timeout=timeout)
    doc = desktop.loadComponentFromURL(
        "private:factory/simpress", "_blank", 0, _props(Hidden=True)
    )
    if doc is None:
        raise SofficeError("LibreOffice could not create a new presentation")
    return doc


# Content types for the media LibreOffice writes into a .pptx but forgets to
# declare.  Straight from the OOXML defaults; the validator names the same ones.
_CONTENT_TYPES = {
    "bin": "application/vnd.openxmlformats-officedocument.oleObject",
    "bmp": "image/bmp",
    "emf": "image/x-emf",
    "gif": "image/gif",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "m4a": "audio/mp4",
    "mp3": "audio/mpeg",
    "mp4": "video/mp4",
    "png": "image/png",
    "svg": "image/svg+xml",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "wav": "audio/wav",
    "wmf": "image/x-wmf",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}

_DEFAULT_RE = re.compile(r'<Default\s+Extension="([^"]+)"', re.I)
_TYPES_OPEN_RE = re.compile(r"(<Types\b[^>]*>)", re.I)


def repair_pptx(path) -> list:
    """Fix the two package defects LibreOffice writes into every ``.pptx``.

    Both are introduced by the save, not by the template -- a pristine
    ``company-template.pptx`` validates clean and its own saved copy does not:

    * ``_rels/.rels`` gets ``Target="../customXml/item1.xml"``.  The package
      root has no parent, so the reference dangles.  ``validate.py`` reports it
      **CRITICAL, "will cause the document to appear corrupt"**.
    * ``[Content_Types].xml`` loses the ``<Default>`` for extensions still
      present in the package -- typically ``wmf``, from think-cell artwork.

    Left alone these fire 22 findings on every deck this skill produces, which
    trains you to ignore the validator, which is how a real defect ships.
    Returns a list of what it repaired, empty when there was nothing to do.
    """
    path = Path(path)
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        parts = {name: archive.read(name) for name in names}

    fixed = []

    # 1. Root-relative relationship targets that climb out of the package.
    root_rels = "_rels/.rels"
    if root_rels in parts:
        before = parts[root_rels].decode("utf8")
        after = re.sub(r'(\bTarget=")\.\./', r"\1", before)
        if after != before:
            parts[root_rels] = after.encode("utf8")
            fixed.append(f"{root_rels}: dropped '../' from "
                         f"{before.count('Target=\"../')} relationship target(s)")

    # 2. Media in the package with no content type declared for its extension.
    types_part = "[Content_Types].xml"
    if types_part in parts:
        content_types = parts[types_part].decode("utf8")
        declared = {m.lower() for m in _DEFAULT_RE.findall(content_types)}
        present = {
            Path(name).suffix.lstrip(".").lower()
            for name in names
            if not name.endswith("/")
        }
        missing = sorted(
            ext for ext in present - declared - {"xml", "rels"}
            if ext in _CONTENT_TYPES
        )
        if missing:
            added = "".join(
                f'<Default Extension="{ext}" ContentType="{_CONTENT_TYPES[ext]}"/>'
                for ext in missing
            )
            content_types = _TYPES_OPEN_RE.sub(r"\1" + added, content_types, count=1)
            parts[types_part] = content_types.encode("utf8")
            fixed.append(f"{types_part}: declared {', '.join(missing)}")

    if not fixed:
        return []

    # Rewrite whole: editing a zip entry in place leaves the old one behind.
    staged = path.with_suffix(path.suffix + ".repair-tmp")
    with zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, parts[name])
    staged.replace(path)
    return fixed


def save_as(doc, path, repair: bool = True):
    """Write the document to ``path``, picking the filter from its extension.

    Leaves the document's own location alone, so a load-edit-save_as chain never
    overwrites the input.

    For a ``.pptx``, ``repair_pptx`` then fixes the package defects the export
    filter introduces -- pass ``repair=False`` to see the raw output.
    """
    path = Path(path).absolute()
    doc.storeToURL(
        uno.systemPathToFileUrl(str(path)), _props(FilterName=filter_for(path), Overwrite=True)
    )
    if repair and path.suffix.lower() in _OOXML_EXTS:
        repair_pptx(path)
    return path


@contextlib.contextmanager
def impress_document(path, save: bool = True):
    """Open a presentation; save it in place, in its own format, on a clean exit.

    ``save=False`` opens read-only -- use it for inspection, and whenever you
    must not touch the file.  An exception inside the block leaves the file on
    disk untouched.
    """
    doc = load_document(path, read_only=not save)
    saved = False
    try:
        yield doc
        if save:
            doc.setModified(True)  # store() is a no-op on an unmodified document
            doc.store()
            saved = True
    finally:
        with contextlib.suppress(Exception):
            doc.close(False)
    if saved and Path(path).suffix.lower() in _OOXML_EXTS:
        repair_pptx(path)          # same export filter, same defects to mend


# --- slides ---------------------------------------------------------------


def slides(doc):
    """The document's slides, in order."""
    pages = doc.getDrawPages()
    return [pages.getByIndex(i) for i in range(pages.getCount())]


def masters(doc):
    """The document's master pages, in order."""
    pages = doc.getMasterPages()
    return [pages.getByIndex(i) for i in range(pages.getCount())]


def add_slide(doc, index: int = None, layout: int = 20):
    """Insert a slide and return it.  ``index`` defaults to the end.

    ``layout`` is an autolayout number: 0 title+subtitle, 1 title+content,
    19 title only, 20 blank.  Anything you place yourself wants 20.
    """
    pages = doc.getDrawPages()
    if index is None:
        index = pages.getCount()
    pages.insertNewByIndex(index)
    page = pages.getByIndex(index)
    page.Layout = layout
    return page


def duplicate_slide(doc, page):
    """Copy a slide, with its shapes, notes and layout, right after the original.

    This is the reliable way to reuse a template slide: it clones every part the
    slide references, which hand-copying the underlying XML does not.
    """
    return doc.duplicate(page)


def delete_slide(doc, page):
    """Remove a slide.  A document must keep at least one.

    Removing the last one wedges LibreOffice -- the call never returns and the
    bridge has to be killed -- so this refuses instead.
    """
    pages = doc.getDrawPages()
    if pages.getCount() <= 1:
        raise SofficeError(
            "cannot delete the only slide: a presentation must keep at least one. "
            "Add the replacement slide first, then delete this one."
        )
    pages.remove(page)


class _Shapes(list):
    """A shape list that says what is on the slide when an index misses.

    ``shapes(page)[12]`` on a nine-shape design is a routine mistake -- the
    indices come from a *different* design, or from before ``draw.canvas()``
    removed half of them.  A bare "list index out of range" sends you looking
    at the wrong thing; the slide's actual contents end it immediately.
    """

    def __getitem__(self, key):
        try:
            return list.__getitem__(self, key)
        except IndexError:
            lines = []
            for i, shape in enumerate(self):
                try:
                    text = " ".join(shape.getString().split())[:52]
                except Exception:
                    text = ""
                lines.append("    [%d] %s%s" % (
                    i, getattr(shape, "Name", "") or shape.getShapeType(),
                    "  %r" % text if text else ""))
            raise IndexError(
                "shapes(page)[%r]: this slide has %d shape(s), indices 0-%d.\n"
                "%s\n  Indices are per design and change after draw.canvas() "
                "removes shapes.\n  Prefer retext()/find_shapes() by wording, "
                "which does not depend on position." % (
                    key, len(self), len(self) - 1, "\n".join(lines) or "    (none)")
            ) from None


def shapes(page):
    """Every shape on a slide (or on a notes page), in z-order."""
    return _Shapes(page.getByIndex(i) for i in range(page.getCount()))


def shape_text(shape) -> str:
    return getattr(shape, "String", "") or ""


def notes_page(page):
    """The notes page belonging to a slide."""
    return page.NotesPage


def notes_text(page) -> str:
    """The speaker notes of a slide, as plain text.

    The notes page also carries a thumbnail of the slide itself; only the
    notes placeholder holds text, so the rest contributes nothing.
    """
    parts = []
    for shape in shapes(page.NotesPage):
        if shape.getShapeType() == "com.sun.star.presentation.NotesShape":
            text = shape_text(shape).strip()
            if text:
                parts.append(text)
    return "\n".join(parts)


def set_notes(page, text: str):
    """Replace a slide's speaker notes."""
    for shape in shapes(page.NotesPage):
        if shape.getShapeType() == "com.sun.star.presentation.NotesShape":
            shape.setString(text)
            return shape
    raise SofficeError(f"slide {page.Name!r} has no notes placeholder")


def placeholder(page, kind: str = "Title"):
    """A slide's presentation placeholder by kind, or None.

    ``kind`` matches the shape type: Title, Subtitle, Outline, Notes.
    Filling these rather than adding text boxes is what keeps a slide on its
    layout -- text typed into a placeholder inherits the master's fonts and
    positions, a free text box inherits nothing.
    """
    want = f"com.sun.star.presentation.{kind}TextShape"
    legacy = f"com.sun.star.presentation.{kind}Shape"
    for shape in shapes(page):
        if shape.getShapeType() in (want, legacy):
            return shape
    return None


# Presentation placeholder shape types -> a short, readable slot label.  Only
# the ones with a stable meaning are named; anything else falls back to the
# shape's own Name (which for imported PowerPoint carries the placeholder name,
# e.g. "Action Title").
_PLACEHOLDER_LABELS = {
    "TitleTextShape": "title",
    "SubtitleShape": "subtitle",
    "OutlineTextShape": "content",
    "OutlinerShape": "content",
    "NotesTextShape": "notes",
    "NotesShape": "notes",
    "HeaderTextShape": "header",
    "FooterTextShape": "footer",
    "SlideNumberTextShape": "slide number",
    "DateTimeTextShape": "date/time",
}


def is_empty_placeholder(shape) -> bool:
    """True if this shape is an empty presentation placeholder.

    Keys on ``IsEmptyPresentationObject`` -- the one signal that is set on a
    placeholder inheriting its geometry from the master and currently holding no
    text, and is ``False`` on a filled or baked-in shape.  This is what makes an
    empty *action title* (or any custom placeholder) visible instead of the
    slide reading as blank: the shape type alone does not tell you, because a
    custom placeholder imports under an unpredictable type.
    """
    return getattr(shape, "IsEmptyPresentationObject", False) is True


def placeholder_kind(shape) -> str:
    """A readable label for a placeholder slot: 'title', 'content', or, for a
    custom placeholder with no standard type, its own name ('Action Title')."""
    short = shape.getShapeType().rsplit(".", 1)[-1]
    label = _PLACEHOLDER_LABELS.get(short)
    if label:
        return label
    name = (getattr(shape, "Name", "") or "").strip()
    return name or short


# Chrome placeholders that sit on every master and say nothing about what a
# layout is *for* -- excluded from the slot report.
_META_PLACEHOLDER_TYPES = {
    "HeaderTextShape", "HeaderShape",
    "FooterTextShape", "FooterShape",
    "SlideNumberTextShape", "SlideNumberShape",
    "DateTimeTextShape", "DateTimeShape",
    "NotesTextShape", "NotesShape",
    "PageShape",
}


def _presentation_placeholders(page):
    """Presentation placeholder shapes on a page, by z-order.  Decoration on a
    master (logos, background rectangles) is drawing.* and excluded; only
    presentation.* placeholder shapes come back."""
    return [s for s in shapes(page)
            if s.getShapeType().startswith("com.sun.star.presentation.")]


# Which placeholder kinds each Impress autolayout actually offers.  The master
# always carries the full default set (title, outline, footer, number, date),
# so the master alone would over-report; the autolayout is what says which slots
# are live for a given slide.  Only the layouts this stack is sure of are listed
# (measured -- see SKILL.md); an unknown layout falls back to the master, which
# is where a custom PowerPoint placeholder (e.g. an action title) shows up.
_AUTOLAYOUT_KINDS = {
    0: {"title", "subtitle"},   # title slide
    1: {"title", "content"},    # title + content
    19: {"title"},              # title only
    20: set(),                  # blank -- no placeholders
}


def layout_slots(page):
    """Content placeholder slots the slide's *layout* offers but the slide has
    no shape for -- the "Click to add Title" case, where the placeholder lives
    only in the layout/master and never lands on the slide's own page.

    Without this a slide that inherits an empty title from its layout reads as
    blank: ``shapes(page)`` returns nothing for it.  The slots are gated by the
    autolayout (``page.Layout``) so the master's always-present default
    placeholders are not over-reported, then read off ``MasterPage`` for their
    geometry, minus the chrome (footer, page number, date) and minus any slot
    the slide already carries a shape for (so it is not double-counted with what
    ``shapes(page)`` shows).

    Returns dicts: ``kind`` (label), ``name``, and the master's geometry.
    """
    layout = getattr(page, "Layout", None)
    allowed = _AUTOLAYOUT_KINDS.get(layout)   # None == layout we don't model
    if allowed is not None and not allowed:
        return []                             # known blank layout: no slots
    master = getattr(page, "MasterPage", None)
    if master is None:
        return []
    on_page = {placeholder_kind(s) for s in _presentation_placeholders(page)}
    out = []
    for s in _presentation_placeholders(master):
        short = s.getShapeType().rsplit(".", 1)[-1]
        if short in _META_PLACEHOLDER_TYPES:
            continue
        kind = placeholder_kind(s)
        if kind in on_page:
            continue
        if allowed is not None and kind not in allowed:
            continue                          # known layout: only its own kinds
        pos, size = s.getPosition(), s.getSize()
        out.append({
            "kind": kind,
            "name": getattr(s, "Name", ""),
            "x": pos.X, "y": pos.Y,
            "width": size.Width, "height": size.Height,
        })
    return out


def add_textbox(doc, page, x: int, y: int, width: int, height: int,
                text: str = "", **style):
    """Add a text box.  All geometry in 1/100 mm -- use mm() or pt().

    With no ``style`` the box inherits Impress's drawing defaults (Arial 18,
    black, left) -- which is almost never what a slide wants.  Pass any of
    ``size`` (pt), ``color``, ``weight`` (``draw.BOLD``), ``align``, ``valign``,
    ``font``, ``italic``, ``underline``, ``pad``, ``fill``, ``line`` and the box
    is styled for you; see ``office.draw.text``, which this then delegates to.

    **Style after the text, never before.**  ``setString()`` on an empty shape
    re-inherits the drawing default and silently discards character formatting
    set before it -- the reason font sizes "don't stick".  Both paths here get
    the order right; if you call ``setString()`` yourself, follow it with
    ``office.draw.restyle(shape, ...)``.
    """
    if style:
        from . import draw          # deferred: draw imports this module
        return draw.text(doc, page, x, y, width, height, text, **style)

    shape = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(shape)
    shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", x, y))
    shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", width, height))
    if text:
        shape.setString(text)
    return shape


def add_image(doc, page, image_path, x: int, y: int, width: int, height: int):
    """Place an image file on a slide.  Geometry in 1/100 mm."""
    image_path = Path(image_path).absolute()
    if not image_path.exists():
        raise FileNotFoundError(f"{image_path} does not exist")
    shape = doc.createInstance("com.sun.star.drawing.GraphicObjectShape")
    page.add(shape)
    shape.GraphicURL = uno.systemPathToFileUrl(str(image_path))
    shape.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", x, y))
    shape.setSize(uno.createUnoStruct("com.sun.star.awt.Size", width, height))
    return shape


def slide_size(doc):
    """(width, height) of a slide in 1/100 mm.

    Ask this before placing anything.  Templates are not all the same size and
    the difference is large enough to push a right-aligned column off the page.
    """
    page = doc.getDrawPages().getByIndex(0)
    return page.Width, page.Height


def set_slide_size(doc, width: int, height: int):
    """Resize every slide and master.  A 16:9 LibreOffice page is 28000 x 15750.

    Only for decks you are creating.  Never resize a brand template: its masters
    are drawn to its own dimensions and everything on them would shift.
    """
    for page in slides(doc) + masters(doc):
        page.Width = width
        page.Height = height
