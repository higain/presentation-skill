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

All geometry is in **1/100 mm** (UNO's drawing unit), not points or EMU:
a 16:9 slide is 28000 x 15750.  ``mm(15)`` and ``pt(18)`` convert.

One headless LibreOffice is started on first use and shut down when the process
ends, so many documents in one script cost one startup.
"""

import atexit
import contextlib
import os
import shutil
import subprocess
import tempfile
import time
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

PRESENTATION_SERVICE = "com.sun.star.presentation.PresentationDocument"


class SofficeError(RuntimeError):
    pass


def mm(value: float) -> int:
    """Millimetres to UNO's 1/100 mm."""
    return int(round(value * 100))


def pt(value: float) -> int:
    """Points to UNO's 1/100 mm (1 pt = 1/72 inch)."""
    return int(round(value * 2540 / 72))


class _Server:
    """A headless soffice listening on a private loopback port."""

    def __init__(self):
        self.proc = None
        self.ctx = None
        self.profile = None

    def desktop(self, timeout: float = 60.0):
        if self.ctx is None:
            self._start(timeout)
        return self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

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


def save_as(doc, path):
    """Write the document to ``path``, picking the filter from its extension.

    Leaves the document's own location alone, so a load-edit-save_as chain never
    overwrites the input.
    """
    path = Path(path).absolute()
    doc.storeToURL(
        uno.systemPathToFileUrl(str(path)), _props(FilterName=filter_for(path), Overwrite=True)
    )
    return path


@contextlib.contextmanager
def impress_document(path, save: bool = True):
    """Open a presentation; save it in place, in its own format, on a clean exit.

    ``save=False`` opens read-only -- use it for inspection, and whenever you
    must not touch the file.  An exception inside the block leaves the file on
    disk untouched.
    """
    doc = load_document(path, read_only=not save)
    try:
        yield doc
        if save:
            doc.setModified(True)  # store() is a no-op on an unmodified document
            doc.store()
    finally:
        with contextlib.suppress(Exception):
            doc.close(False)


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


def shapes(page):
    """Every shape on a slide (or on a notes page), in z-order."""
    return [page.getByIndex(i) for i in range(page.getCount())]


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


def add_textbox(doc, page, x: int, y: int, width: int, height: int, text: str = ""):
    """Add a text box.  All geometry in 1/100 mm -- use mm() or pt()."""
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
    """(width, height) of a slide in 1/100 mm."""
    page = doc.getDrawPages().getByIndex(0)
    return page.Width, page.Height


def set_slide_size(doc, width: int, height: int):
    """Resize every slide and master.  16:9 is 28000 x 15750."""
    for page in slides(doc) + masters(doc):
        page.Width = width
        page.Height = height
