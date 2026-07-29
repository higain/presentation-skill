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

Measured against the delta template, picking designs 7, 8, 13, 20 and 1 and
diffing each result against its source render: mean RMS 7.8 here versus 22.8 for
``clone_slide``, better on every slide, and 6 think-cell references preserved
where ``clone_slide`` kept none.  Output size and validator findings are the
same either way.

``.pptx`` only -- the reorder step is OOXML.  For ``.odp``, or for copying a
slide between two open documents, use ``clone_slide``.
"""

from __future__ import annotations

import inspect
import re
import shutil
import zipfile
from pathlib import Path

from .impress import (
    load_document, save_as, slides, duplicate_slide, delete_slide,
)

_SLDID_RE = re.compile(r'<p:sldId[^>]*r:id="rId\d+"[^>]*/>')
_SLDIDLST_RE = re.compile(r"<p:sldIdLst>.*?</p:sldIdLst>", re.S)


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


def build_deck(template, out, designs, fill=None, qa=True) -> Path:
    """Write ``out``: one slide per entry of ``designs``, in that order.

    ``designs`` holds 1-based slide numbers in ``template`` -- the numbers a
    layout catalog cites -- and may repeat one.

    ``fill(doc, page, slot, design)``, if given, is called on each new page
    before the save.  This is where **all** of the content goes: retexting with
    ``office.template.retext`` and drawing with ``office.draw``.  Do not reopen
    the finished file to add shapes -- that is a second .pptx round trip, it
    grows the masters, and it is how overlays end up on top of template text
    nobody replaced.

    ``qa=True`` runs ``scripts/qa.py``'s checks on the result and prints what it
    finds -- overflow and leftover template wording, the two defects that are
    invisible without rendering.  It never raises; it reports.

    The two traps this handles, both of which silently produce the wrong deck:
    ``duplicate_slide`` inserts each copy next to its original, so every later
    index shifts and the sources must all be resolved up front; and the copies
    end up interleaved with the originals, so the originals have to be deleted
    by identity rather than by index.
    """
    template, out = Path(template), Path(out)
    if out.suffix.lower() != ".pptx":
        raise ValueError(
            f"build_deck writes .pptx only (got {out.suffix or 'no suffix'}); "
            f"reordering is an OOXML step -- use clone_slide for .odp"
        )
    if not designs:
        raise ValueError("no designs given")

    shutil.copyfile(template, out)
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

        if fill is not None:
            for slot, page in enumerate(made):
                _call_fill(fill, doc, page, slot, designs[slot])

        save_as(doc, out)
    finally:
        doc.close(False)

    reorder_pptx(out, order)
    if qa:
        _report_qa(out, template)
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
        issues, checked = check(out, template=template)
    except Exception as e:                    # QA must never lose you the deck
        print(f"[build_deck] QA could not run: {e}")
        return
    print(f"\n[build_deck] {out.name}: {checked} slide(s) built.")
    if issues:
        print(f"[build_deck] {len(issues)} QA issue(s) -- fix these before shipping:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("[build_deck] QA: no overflow and no leftover template text.")
    print("[build_deck] QA cannot see the slides. Render them "
          "(python3 scripts/render.py) and look.\n")


def _index_of(page, pages):
    """Position of ``page`` in ``pages``, by UNO object identity."""
    for i, candidate in enumerate(pages):
        if candidate == page:
            return i
    raise RuntimeError("a page went missing between duplicate and save")
