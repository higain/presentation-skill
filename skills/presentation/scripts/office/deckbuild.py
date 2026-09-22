"""Build a deck from a template at full fidelity, in any slide order.

``clone_slide`` rebuilds every shape property by property, and it only does that
because UNO cannot reorder pages: ``DrawPage.Number`` is read-only, and the
slide-sorter dispatches (``.uno:MovePageUp`` and friends) are silent no-ops in a
headless frame.  Rebuilding is lossy -- bullet markers, per-run colour inside a
paragraph, and think-cell's OLE data all go.

But slide order in a ``.pptx`` is nothing more than the order of ``<p:sldId>``
entries in ``ppt/presentation.xml``; the slide parts themselves are
order-independent.  So this module never rebuilds a shape.  It duplicates the
designs it wants -- ``duplicate_slide`` is a real Impress duplicate, lossless --
drops the template originals, saves, and then permutes that one list in the
saved file.

Measured against a 21-design corporate template, picking 7, 8, 13, 20, 1 and
diffing each result against its source render: mean RMS 7.8 here versus 22.8 for
``clone_slide``, better on every slide, and 6 think-cell references preserved
where ``clone_slide`` kept none.  Output size and validator findings are the
same either way.

``.pptx`` only -- the reorder step is OOXML.  For ``.odp``, or for copying a
slide between two open documents, use ``clone_slide``.
"""

from __future__ import annotations

import inspect
import os
import re
import shutil
import zipfile
from pathlib import Path

from .pack import (
    SKILL, TEMPLATES, active, default_pack, describe_install, pack_named,
)

from .impress import (
    load_document, save_as, slides, duplicate_slide, delete_slide,
)
from .template import set_footer

#: Set by ``make_deck.py`` around the build script it runs.  QA on a 15-slide
#: deck is a full LibreOffice load and about a minute of measuring, and
#: ``make_deck.py`` runs exactly the same check on the same file moments later
#: -- the documented recipe was paying for it twice.  A build script run on its
#: own still QAs itself.
SKIP_QA_ENV = "PRESENTATION_SKIP_BUILD_QA"

#: ``footer=KEEP`` -- the default -- leaves the template's footer alone, which
#: is not the same instruction as ``footer=None`` (hide it) or ``footer=""``
#: (blank it).  A plain ``None`` default could not tell the first from the
#: second.
KEEP = object()

#: Document property the design list is written to, so what a deck was built
#: from survives the build script and can be read back off the file.  ``qa.py``
#: reports the variety from it; nothing depends on it being there.
DESIGNS_PROPERTY = "Designs"

_SLDID_RE = re.compile(r'<p:sldId[^>]*r:id="rId\d+"[^>]*/>')
_SLDIDLST_RE = re.compile(r"<p:sldIdLst>.*?</p:sldIdLst>", re.S)

_MASTER_PART_RE = re.compile(r"^ppt/slideMasters/slideMaster(\d+)\.xml$")
_THEME_TARGET_RE = re.compile(r"theme/(theme\d+\.xml)")
_LAYOUT_TARGET_RE = re.compile(r"slideLayouts/(slideLayout\d+\.xml)")
_CSLD_NAME_RE = re.compile(r'<p:cSld[^>]*\bname="([^"]*)"')
_THEME_ELEMENTS_RE = re.compile(r"<a:themeElements>.*?</a:themeElements>", re.S)
_CLRSCHEME_RE = re.compile(r"<a:clrScheme\b.*?</a:clrScheme>", re.S)

#: The twelve slots of a theme colour scheme, in schema order.
_SCHEME_SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink")


def reorder_pptx(path: Path, slot_at_position) -> Path:
    """Permute a saved deck's slide order.

    ``slot_at_position[j]`` is the slot number of the slide currently sitting at
    document position ``j``; on return, slot ``k`` sits at position ``k``.
    Rewrites only ``<p:sldIdLst>`` -- every slide part, relationship id and
    media file is left exactly as it was.
    """
    path = Path(path)
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        parts = {name: archive.read(name) for name in names}

    presentation = parts["ppt/presentation.xml"].decode("utf8")
    entries = _SLDID_RE.findall(presentation)
    if len(entries) != len(slot_at_position):
        raise ValueError(
            f"{path.name} has {len(entries)} slides but {len(slot_at_position)} "
            f"order slots were given"
        )

    ordered: list[str | None] = [None] * len(entries)
    for position, slot in enumerate(slot_at_position):
        if ordered[slot] is not None:
            raise ValueError(f"slot {slot} given twice in the slide order")
        ordered[slot] = entries[position]

    parts["ppt/presentation.xml"] = _SLDIDLST_RE.sub(
        "<p:sldIdLst>" + "".join(ordered) + "</p:sldIdLst>",
        presentation, count=1,
    ).encode("utf8")

    # Rewrite whole: an in-place zip edit would leave the old entry behind.
    staged = path.with_suffix(".reorder-tmp")
    with zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
        for name in names:
            out.writestr(name, parts[name])
    staged.replace(path)
    return path


def _scheme_colours(theme_xml: str):
    """The twelve theme colours as lowercase hex, or ``None`` if unreadable.

    Compares by *resolved colour* rather than by markup, because the two files
    write the same colour differently: a hand-built template says
    ``<a:sysClr val="window" lastClr="FFFFFF"/>`` where LibreOffice's export
    says ``<a:srgbClr val="ffffff"/>``.  Comparing the XML would call those a
    difference and rewrite every theme in every deck.
    """
    elements = _THEME_ELEMENTS_RE.search(theme_xml)
    if not elements:
        return None
    scheme = _CLRSCHEME_RE.search(elements.group(0))
    if not scheme:
        return None
    body = scheme.group(0)
    found = []
    for slot in _SCHEME_SLOTS:
        m = re.search(
            rf"<a:{slot}>\s*<a:(?:srgbClr\s+val|sysClr[^>]*?\blastClr)="
            rf'"([0-9A-Fa-f]{{6}})"', body)
        if not m:
            return None
        found.append(m.group(1).lower())
    return tuple(found)


def _master_map(parts: dict[str, bytes]):
    """``{master part: (theme part, {layout names})}`` for one .pptx."""
    out = {}
    for name in parts:
        if not _MASTER_PART_RE.match(name):
            continue
        rels_name = name.replace("slideMasters/", "slideMasters/_rels/") + ".rels"
        rels = parts.get(rels_name, b"").decode("utf8", "replace")
        theme = _THEME_TARGET_RE.search(rels)
        if not theme:
            continue
        names = set()
        for layout in _LAYOUT_TARGET_RE.findall(rels):
            xml = parts.get("ppt/slideLayouts/" + layout, b"").decode("utf8", "replace")
            found = _CSLD_NAME_RE.search(xml)
            if found:
                names.add(found.group(1))
        out[name] = ("ppt/theme/" + theme.group(1), names)
    return out


def repair_theme_colours(template: Path, out: Path) -> list[str]:
    """Restore the template's theme colours on a deck LibreOffice re-exported.

    **The reason this exists.**  ``build_deck`` copies the template, opens it in
    Impress and saves it back, and that round trip does not preserve the theme
    colour scheme.  Measured on a 3-master template whose every theme sets
    ``dk2`` to the brand navy ``#022A3A``: the saved deck came back with *six*
    masters, of which two were invented by LibreOffice and carried its own
    default scheme (``dk2=cccccc``), while the template's real masters were
    renumbered out to 44-47.  A section-divider layout paints its background as
    ``<a:schemeClr val="dk2"/>``, so the divider rendered light grey with white
    text on it -- about 1.6:1, and nothing in QA can see it because QA reads
    shapes, not the theme behind them.

    The fix maps each exported master back to the template master it came from
    by **layout-name overlap** -- layout names survive the round trip even when
    master numbering does not -- and rewrites that master's ``<a:clrScheme>``
    to the one the template gave it.  Only the colour scheme is touched; the
    font and format schemes LibreOffice writes are left alone.

    Conservative by construction.  A theme is rewritten only when it is
    reachable from a master, maps to exactly one template scheme, and actually
    differs from it.  A theme shared by masters that disagree is left alone --
    rewriting it would fix one master by breaking another.  Returns a line per
    theme changed, for the caller to report.
    """
    with zipfile.ZipFile(template) as archive:
        src_parts = {name: archive.read(name) for name in archive.namelist()}
    with zipfile.ZipFile(out) as archive:
        names = archive.namelist()
        parts = {name: archive.read(name) for name in names}

    src_masters = _master_map(src_parts)
    if not src_masters:
        return []

    # {template theme part: (scheme colours, layout names it covers)}
    src_schemes = {}
    for theme, layout_names in src_masters.values():
        colours = _scheme_colours(src_parts.get(theme, b"").decode("utf8", "replace"))
        if colours is None:
            continue
        known, covered = src_schemes.get(theme, (colours, set()))
        src_schemes[theme] = (known, covered | layout_names)
    if not src_schemes:
        return []

    fallback = None
    if len({colours for colours, _ in src_schemes.values()}) == 1:
        fallback = next(iter(src_schemes))

    # Choose a template scheme per output theme, refusing ties.
    wanted: dict[str, set[str]] = {}
    for theme, layout_names in _master_map(parts).values():
        scored = [(len(layout_names & covered), part)
                  for part, (_, covered) in src_schemes.items()]
        best = max(score for score, _ in scored)
        if best == 0:
            chosen = fallback
        else:
            tied = [part for score, part in scored if score == best]
            # Masters that tie but specify the same colours are not in
            # conflict: any of them names the scheme the deck should carry.
            chosen = tied[0] if len(
                {src_schemes[part][0] for part in tied}) == 1 else None
        if chosen:
            wanted.setdefault(theme, set()).add(chosen)

    changed = []
    for theme, sources in wanted.items():
        if len({src_schemes[part][0] for part in sources}) != 1:
            continue                          # masters disagree; leave it
        source = next(iter(sources))
        target = src_schemes[source][0]
        xml = parts.get(theme, b"").decode("utf8", "replace")
        if _scheme_colours(xml) == target:
            continue
        elements = _THEME_ELEMENTS_RE.search(xml)
        replacement = _CLRSCHEME_RE.search(
            _THEME_ELEMENTS_RE.search(
                src_parts[source].decode("utf8", "replace")).group(0))
        if not elements or not replacement:
            continue
        patched = _CLRSCHEME_RE.sub(
            lambda _: replacement.group(0), elements.group(0), count=1)
        parts[theme] = (xml[:elements.start()] + patched
                        + xml[elements.end():]).encode("utf8")
        changed.append(f"{theme}: colour scheme restored from {source}")

    if not changed:
        return []

    staged = Path(out).with_suffix(".theme-tmp")
    with zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.writestr(name, parts[name])
    staged.replace(out)
    return changed


def _call_fill(fill, doc, page, slot, design):
    """Call ``fill`` as ``(doc, page, slot, design)``, or the old ``(page, slot,
    design)`` if that is what it takes.

    ``doc`` is what every shape constructor needs -- ``doc.createInstance(...)``
    -- so a filler that draws cannot work without it.  Older fillers that only
    retext do not want it, and still run unchanged.
    """
    try:
        params = inspect.signature(fill).parameters
        wants_doc = len(params) >= 4 or any(
            p.kind is inspect.Parameter.VAR_POSITIONAL for p in params.values()
        )
    except (TypeError, ValueError):
        wants_doc = True                      # builtin or C callable: assume new
    return fill(doc, page, slot, design) if wants_doc else fill(page, slot, design)


def find_template(template=None) -> Path:
    """Resolve a template: a path, a pack name, or the default pack.

    A build script is written next to the deck and run from there, but the
    template it names is quoted from this skill's documentation -- so a bare
    pack-relative path is relative to the *skill*, not to the working
    directory, and every first build died on ``FileNotFoundError`` before the
    script had a chance to be wrong about anything else.  Resolving it here
    costs nothing and an absolute path still wins.

    ``find_template()`` with nothing at all is the common case now: with one
    pack installed it is that pack's template, which is what "build a deck"
    with no template named should mean.  ``find_template("acme")`` names a pack
    directly, so a build script never has to spell out a path inside the skill.

    Anything that takes a template path from the caller comes through here.
    ``make_deck.py`` used to test ``--template`` with a bare ``Path.exists()``,
    so the *same string* the documentation prints twice -- once for
    ``build_deck``, once for ``--template`` -- resolved in one place and died
    in the other.
    """
    if template is None:
        pack = default_pack()
        if pack is None:
            raise FileNotFoundError(
                "build_deck() was given no template. " + describe_install())
        return pack.template

    path = Path(template)
    if path.exists():
        return path

    named = pack_named(str(template))
    if named is not None and named.template is not None:
        return named.template

    for candidate in (SKILL / path, TEMPLATES / path.name):
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        f"no template at {template!r}, and none at {SKILL / path} either.\n"
        f"  {describe_install()}"
    )


def build_deck(template, out, designs, fill=None, qa=True,
               footer=KEEP) -> Path:
    """Write ``out``: one slide per entry of ``designs``, in that order.

    ``designs`` holds 1-based slide numbers in ``template`` -- the numbers a
    layout catalog cites -- and may repeat one.  **Repeating one for the whole
    deck is the failure mode to watch**: four separate runs against the house
    template built ``[8] * N``, because design 8 stripped to a canvas is the
    documented way to draw anything, and a deck of nothing but drawn white
    slides uses none of the twenty designs it was handed.  The list is recorded
    on the file and ``qa.py`` reports what share of the deck sits on one design.

    ``footer`` sets the footer line on every slide -- the template's own is a
    page property that ``retext`` cannot reach, so a deck built from a template
    in another language keeps that language along the bottom of every slide
    unless this is passed.  ``footer=None`` hides it; the default leaves the
    template's alone.

    ``fill(doc, page, slot, design)``, if given, is called on each new page
    before the save.  This is where **all** of the content goes: retexting with
    ``office.template.retext`` and drawing with ``office.draw``.  Do not reopen
    the finished file to add shapes -- that is a second .pptx round trip, it
    grows the masters, and it is how overlays end up on top of template text
    nobody replaced.

    ``qa=True`` runs ``scripts/qa.py``'s checks on the result and prints what it
    finds -- overflow and leftover template wording, the two defects that are
    invisible without rendering.  It never raises; it reports.  Under
    ``make_deck.py`` it is skipped (see :data:`SKIP_QA_ENV`), because that
    command runs the same check itself a moment later.

    The two traps this handles, both of which silently produce the wrong deck:
    ``duplicate_slide`` inserts each copy next to its original, so every later
    index shifts and the sources must all be resolved up front; and the copies
    end up interleaved with the originals, so the originals have to be deleted
    by identity rather than by index.
    """
    template, out = find_template(template), Path(out)
    if out.suffix.lower() != ".pptx":
        raise ValueError(
            f"build_deck writes .pptx only (got {out.suffix or 'no suffix'}); "
            f"reordering is an OOXML step -- use clone_slide for .odp"
        )
    if not designs:
        raise ValueError("no designs given")

    shutil.copyfile(template, out)
    try:
        # Everything drawn inside `fill` resolves its colours and type sizes
        # against *this* template, not against whatever pack is installed.
        with active(template):
            return _build(template, out, designs, fill, qa, footer)
    except BaseException:
        # The output at this point is a copy of the template -- every design in
        # it, none of your content.  Left on disk it looks like a built deck,
        # and the next command to touch it (a bare make_deck.py, a QA run, a
        # human opening it) is looking at the template and does not know.
        out.unlink(missing_ok=True)
        raise


def _build(template: Path, out: Path, designs, fill, qa, footer=KEEP) -> Path:
    """The body of :func:`build_deck`, on an ``out`` that already exists."""
    doc = load_document(out, read_only=False)
    try:
        pages = slides(doc)
        for design in designs:
            if not 1 <= design <= len(pages):
                raise ValueError(
                    f"design {design} is out of range: {template.name} has "
                    f"{len(pages)} slides"
                )
        # Resolve every source page before duplicating anything.
        sources = [pages[design - 1] for design in designs]
        originals = list(pages)
        made = [duplicate_slide(doc, page) for page in sources]

        for page in originals:
            delete_slide(doc, page)

        # Read the resulting order back rather than predicting it: where a
        # duplicate lands is the API's business, not ours.
        order = [_index_of(page, made) for page in slides(doc)]
        if sorted(order) != list(range(len(made))):
            raise RuntimeError(f"could not place every duplicate: {order}")

        # Before ``fill``, so a build script that wants a per-slide footer can
        # still set its own over the top.
        if footer is not KEEP:
            for page in made:
                set_footer(page, footer)

        if fill is not None:
            for slot, page in enumerate(made):
                _call_fill(fill, doc, page, slot, designs[slot])

        _record_designs(doc, designs)
        save_as(doc, out)
    finally:
        doc.close(False)

    reorder_pptx(out, order)
    # After the save, because it repairs what the save damaged.
    for line in repair_theme_colours(template, out):
        print(f"[build_deck] {line}")
    if qa and not os.environ.get(SKIP_QA_ENV):
        _report_qa(out, template)
    return out


def _record_designs(doc, designs):
    """Write the design list onto the document as a custom property.

    Which designs a deck used is knowable only while the build script is
    running -- the built file has twenty duplicated slides and no memory of
    which template page each came from -- and it is exactly what the run after
    this one wants to know.  A user-defined document property survives the
    ``.pptx`` save (it lands in ``docProps/custom.xml``) and costs nothing.

    Best effort: a format or a build that will not take one loses the variety
    line in the report and nothing else.
    """
    value = ",".join(str(d) for d in designs)
    try:
        props = doc.DocumentProperties.UserDefinedProperties
        try:
            props.addProperty(DESIGNS_PROPERTY, 0, value)
        except Exception:                     # already there, from the template
            props.setPropertyValue(DESIGNS_PROPERTY, value)
    except Exception:
        pass


def recorded_designs(doc):
    """The design list :func:`build_deck` recorded, or ``[]``."""
    try:
        raw = doc.DocumentProperties.UserDefinedProperties.getPropertyValue(
            DESIGNS_PROPERTY)
    except Exception:
        return []
    out = []
    for part in str(raw).split(","):
        part = part.strip()
        if part.isdigit():
            out.append(int(part))
    return out


def _report_qa(out: Path, template: Path):
    """Print the QA findings for a freshly built deck.

    Built in rather than left to the caller because the runs that most need it
    are the ones that never get round to running it.  ``build_deck(..., qa=False)``
    turns it off; ``scripts/qa.py`` is the same check with an exit code.
    """
    try:
        from qa import check                  # scripts/ is on sys.path already
    except ImportError:                       # imported from outside scripts/
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        try:
            from qa import check
        except ImportError:
            return
    try:
        issues, checked, sparse, notes, prose = check(out, template=template)
    except Exception as e:                    # QA must never lose you the deck
        print(f"[build_deck] QA could not run: {e}")
        return
    print(f"\n[build_deck] {out.name}: {checked} slide(s) built.")
    for note in notes:
        print(f"[build_deck] {note}")
    if issues:
        print(f"[build_deck] {len(issues)} QA issue(s) -- fix these before shipping:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("[build_deck] QA: no overflow, no collisions, no leftover "
              "template text.")
    if sparse:
        print(f"[build_deck] {len(sparse)} slide(s) worth looking at -- "
              f"{', '.join(str(n) for n, _ in sparse)}: may be half empty.")
    # The ghost deck goes to `scripts/qa.py`, not here: a build prints as it
    # goes and a list of every title would bury the findings above it.  What a
    # build owes you is the count and where to see them.
    prose_issues, _titles = prose
    if prose_issues:
        print(f"[build_deck] {len(prose_issues)} line(s) to look at -- titles "
              f"and wording. `python3 scripts/qa.py {out.name}` prints them "
              f"with the ghost deck.")
    print("[build_deck] QA cannot see the slides. Render them "
          "(python3 scripts/render.py) and look.\n")


def _index_of(page, pages):
    """Position of ``page`` in ``pages``, by UNO object identity."""
    for i, candidate in enumerate(pages):
        if candidate == page:
            return i
    raise RuntimeError("a page went missing between duplicate and save")
