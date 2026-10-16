"""Where the ink is: sparse and top-heavy slides, design variety, stray template pictures."""

from __future__ import annotations

import re
from pathlib import Path

from ..impress import PPTX_EXTS, slides
from ..template import ink_box, walk_shapes

# A template graphic smaller than this share of the slide is chrome -- a rule,
# a bullet marker, a corner mark.  Above it, it is a picture the design chose
# for its own content, and it is still on the slide because nobody removed it.
MIN_STRAY_IMAGE = 0.004

# When a slide is worth a second look because there may be nothing on it.
#
# Measured on two agent decks whose slides had already been judged by eye: the
# ones that read as unfinished painted 9-24% of the slide, the ones that read
# as designed painted 32-89%, and nothing landed between 24% and 32%.  A
# quarter is the middle of that gap.  The band rule catches the other shape of
# the same defect -- content packed against the title with the bottom half of
# the slide bare -- which scored 30-51% where a finished slide scored 8%, the
# margin under the footer.
#
# Neither is a defect a machine can be sure of: a one-sentence statement slide
# is sparse on purpose and scores the same. That is why these are reported as
# slides to look at rather than as findings, and why they do not fail a build.
SPARSE_AREA = 0.25
SPARSE_BAND = 0.30


def _painted_boxes(doc, page):
    """``(x, y, w, h)`` of what each shape on the page actually paints.

    Text goes through ``ink_box``, which lays the words out and measures the
    block they make.  Its own box is nearly always bigger -- and a build that
    drops one full-slide text frame holding two lines would otherwise measure
    as a slide covered corner to corner.  Everything else is its geometry: an
    empty panel is still something to look at.
    """
    out = []
    for _, shape in walk_shapes(page):
        try:
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if size.Width <= 0 or size.Height <= 0:
            continue
        ink = ink_box(doc, page, shape)
        out.append(ink[:4] if ink else (pos.X, pos.Y, size.Width, size.Height))
    return out


def _widest_empty_band(bands, top, bottom):
    """The tallest run of slide between ``top`` and ``bottom`` nothing paints.

    ``bands`` are ``(top, bottom)`` pairs, in any order and overlapping freely;
    this is the classic sweep over them, keeping the largest hole.
    """
    spans = sorted((max(t, top), min(b, bottom)) for t, b in bands
                   if b > top and t < bottom)
    widest, at, cursor = 0, top, top
    for start, end in spans:
        if start - cursor > widest:
            widest, at = start - cursor, cursor
        cursor = max(cursor, end)
    if bottom - cursor > widest:
        widest, at = bottom - cursor, cursor
    return widest, at


def _paint_metrics(doc, page):
    """See :func:`_measure_paint`; once per page within a QA pass."""
    from .. import template
    if template._MEASURED is None:
        return _measure_paint(doc, page)
    key = ("paint", getattr(page, "Name", ""), id(doc))
    if key not in template._MEASURED:
        template._MEASURED[key] = _measure_paint(doc, page)
    return template._MEASURED[key]


def _measure_paint(doc, page):
    """``(area, band, at, height)`` for one slide, or ``None``.

    ``area`` is the painted share (0..1, double-counting overlaps),
    ``band``/``at`` the widest empty vertical band and where it starts, both in
    the page's height units.  Shared by ``sparse_slides`` and
    ``balance_findings`` so the two never disagree about a slide's numbers.
    """
    try:
        width, height = page.Width, page.Height
    except Exception:
        return None
    if width <= 0 or height <= 0:
        return None
    boxes = _painted_boxes(doc, page)
    if not boxes:
        return (0.0, height, 0, height)
    area = sum(w * h for _, _, w, h in boxes) / (width * height)
    bands = [(y, y + h) for _, y, _, h in boxes]
    band, at = _widest_empty_band(bands, min(t for t, _ in bands), height)
    return area, band, at, height


def sparse_slides(doc, wanted=None):
    """``[(number, reason)]`` for slides that may have nothing on them.

    Not findings.  The two checks a deck build actually needs eyes for are "is
    this slide half empty" and "is this diagram off-centre", and the second one
    genuinely cannot be measured.  The first one can: a slide that paints a
    tenth of itself, or leaves a third of its height in one bare band under the
    title, is either a deliberate statement slide or a slide somebody forgot to
    finish, and no measurement tells those apart.

    So this does not fail a build -- it names the slides, and naming four of
    fifteen is a different instruction from "open the PNGs and look at them",
    which is what the last two runs were given.  One of them looked and shipped
    the empty slides anyway; the other never opened an image at all.
    """
    out = []
    for number, page in enumerate(slides(doc), start=1):
        if wanted and number not in wanted:
            continue
        metrics = _paint_metrics(doc, page)
        if metrics is None:
            continue
        area, band, at, height = metrics
        if area <= 0.0:
            out.append((number, "nothing on it at all"))
            continue
        if area < SPARSE_AREA:
            out.append((number, f"paints {area:.0%} of the slide"))
        elif band >= SPARSE_BAND * height:
            out.append((number, f"{band / height:.0%} of its height is one "
                                f"empty band, from {at / height:.0%} down"))
    return out


#: A slide is *top-heavy*, and fails the build, when it paints enough that it
#: is clearly a content slide (not a bare one-sentence statement) yet leaves at
#: least a third of its height as one empty band starting in the lower half.
#: That combination -- real content crammed against the title with the bottom
#: of the slide bare -- is the signature of hardcoded offsets, and unlike a
#: genuinely sparse statement slide it is never a deliberate choice.  The
#: measurements this came from: the phase-canvas slides this check was written
#: for painted 24-34 % with a 33-38 % empty band from 54-58 % down, and every
#: one had been marked "sparse by design" and shipped.
BALANCE_AREA = 0.18
BALANCE_BAND = 0.30
BALANCE_START = 0.35


def _is_top_heavy(area, band, at, height):
    """``True`` if a slide with these paint metrics is a top-heavy finding.

    Pure so the thresholds are testable without a document: real content packed
    above a bare bottom.  ``area`` is the painted share, ``band`` the widest
    empty vertical band and ``at`` where it starts, ``height`` the page height.
    """
    return (area >= BALANCE_AREA
            and band >= BALANCE_BAND * height
            and at >= BALANCE_START * height)


def balance_findings(doc, wanted=None):
    """``[(slide, reason)]`` -- content slides that are top-heavy.

    A hard finding, unlike ``sparse_slides``: a slide with real content packed
    into its top that leaves its bottom half bare is not a deliberate choice,
    and it is invisible to every other check in this file -- nothing overflows,
    nothing collides, nothing is off the slide, the slide is just badly placed.
    It fails the build unless acknowledged in the deck's allowlist.

    ``sparse_slides`` still exists alongside it for the genuine statements: a
    slide that paints under a quarter may be a one-liner on purpose, and that
    stays a "look at it" rather than a finding.
    """
    out = []
    for number, page in enumerate(slides(doc), start=1):
        if wanted and number not in wanted:
            continue
        metrics = _paint_metrics(doc, page)
        if metrics is None:
            continue
        area, band, at, height = metrics
        if _is_top_heavy(area, band, at, height):
            out.append(
                (number, f"top-heavy: paints {area:.0%}, then {band / height:.0%} "
                         f"of the slide's height is one empty band from "
                         f"{at / height:.0%} down. Content is packed against "
                         f"the title with the bottom bare -- lay it out from "
                         f"the `Box` `D.canvas` returns so it fills the slide "
                         f"(`references/drawing.md` §Composition)")
            )
    return out


#: A slide with more hand-placed shapes than this, and nothing from a
#: composition helper, is worth a look first.  Not a defect: a custom diagram
#: no helper draws is laid out from ``Box.grid`` with primitives, exactly as
#: ``drawing.md`` says, and looks the same in the file as a slide placed by
#: page arithmetic.  What actually goes wrong with the second kind -- edges
#: that almost line up, gaps that almost match -- is measured below and
#: fails; the count only points the vision-agent at the slide.
HAND_PLACED = 12
#: Shape names ``office.draw`` gives what a build script placed one by one.
PRIMITIVES = {"rectangle", "text", "ellipse", "line", "polygon", "connector",
              "polyline"}

#: Two edges closer than this are aligned (rounding, a line's width).
ALIGNED = 30
#: Two edges this far apart, and no closer, were meant to line up and do not:
#: 0.6 mm is visible on a projected slide, past 3 mm the offset is a choice.
NEAR_MISS = (60, 300)
#: Peers: same kind, and sizes within this share of each other -- the cells
#: of one row, the cards of one column.
PEER_SIZE = 0.10


def _drawn(page):
    """``[(address, kind, x, y, w, h)]`` -- the shapes a build placed by hand."""
    out = []
    for address, shape in walk_shapes(page):
        name = getattr(shape, "Name", "") or ""
        if not name.startswith("D:") or name[2:] not in PRIMITIVES:
            continue
        try:
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        out.append((address, name[2:], pos.X, pos.Y, size.Width, size.Height))
    return out


def hand_placed_note(page):
    """A reason to look at the slide first, or ``None``.  Never a finding."""
    hand, laid_out = 0, 0
    for _, shape in walk_shapes(page):
        name = getattr(shape, "Name", "") or ""
        if not name.startswith("D:"):
            continue
        if name[2:] in PRIMITIVES:
            hand += 1
        else:
            laid_out += 1
    if hand <= HAND_PLACED or laid_out:
        return None
    return (f"{hand} shapes placed one by one, no layout helper: check that "
            f"rows and columns line up")


#: Text that is a source or a note line, in the languages decks arrive in.
SOURCE_LINE = re.compile(
    r"\b(quellen?|source[sn]?|notes?|anmerkungs?|fußnoten?|basis|stand)\b"
    r"|beispieldaten|illustrati|sample data|example data", re.I)
#: A figure worth a source: a percentage, a currency amount, a decimal, or a
#: number of three digits or more that is not a year.  Step numbers and page
#: numbers are not figures.
FIGURE = re.compile(
    r"[$€£]\s?\d|\d\s?(%|pp|€|\$|£)|\d[.,]\d|\b(?!(?:19|20)\d\d\b)\d{3,}\b")
#: This many figures on a slide, and no source line: worth a look.
SOURCED_FIGURES = 3


def look_notes(page, texts):
    """Reasons to look at a slide first -- never findings.

    ``texts`` are the slide's ``(address, text)`` pairs.  A slide drawn shape
    by shape, a chart placeholder the author still has to fill, figures
    without a source line: each goes into the verdict table, where the
    vision-agent reads it, and none fails a build.
    """
    notes = []
    hand = hand_placed_note(page)
    if hand:
        notes.append(hand)
    if any((getattr(shape, "Name", "") or "") == "D:chart_placeholder"
           for _, shape in walk_shapes(page)):
        notes.append("chart placeholder: the author pastes the Excel chart "
                     "here -- not a defect")
    words = " ".join(text for _, text in texts)
    figures = len(FIGURE.findall(words))
    if figures >= SOURCED_FIGURES and not SOURCE_LINE.search(words):
        notes.append(f"{figures} figures and no source line: say where the "
                     f"numbers come from (`D.source`)")
    return "; ".join(notes) or None


def _peers(a, b):
    if a[1] != b[1]:
        return False
    for i in (4, 5):
        big = max(a[i], b[i])
        if big and abs(a[i] - b[i]) > PEER_SIZE * big:
            return False
    return True


def _inside(a, b):
    """``a`` sits within ``b`` (an inset label), or the other way round."""
    def within(p, q):
        return (p[2] >= q[2] - ALIGNED and p[3] >= q[3] - ALIGNED
                and p[2] + p[4] <= q[2] + q[4] + ALIGNED
                and p[3] + p[5] <= q[3] + q[5] + ALIGNED)
    return within(a, b) or within(b, a)


def _edges(item, axis):
    x, y, w, h = item[2:]
    return (x, x + w / 2, x + w) if axis == "x" else (y, y + h / 2, y + h)


def misalignments(items):
    """``[text]`` -- peers that almost line up, and rows that almost space evenly.

    ``items`` are ``(address, kind, x, y, w, h)`` in 1/100 mm.  Pure, so it is
    tested without LibreOffice.
    """
    misses = []
    for i, a in enumerate(items):
        for b in items[i + 1:]:
            if not _peers(a, b) or _inside(a, b):
                continue
            for axis, other in (("x", "y"), ("y", "x")):
                # Compare left edges of shapes stacked in a column, top edges
                # of shapes side by side in a row.
                ea, eb = _edges(a, axis), _edges(b, axis)
                diffs = [abs(p - q) for p, q in zip(ea, eb)]
                if min(diffs) <= ALIGNED:
                    continue
                miss = min(diffs)
                if NEAR_MISS[0] <= miss <= NEAR_MISS[1]:
                    misses.append((a[0], b[0], axis, miss))
    out = []
    if misses:
        # One drift moves many shapes; say how many, and show one pair.
        shapes = {m[0] for m in misses} | {m[1] for m in misses}
        lo = min(m[3] for m in misses) / 100
        hi = max(m[3] for m in misses) / 100
        first = misses[0]
        edge = "left/centre/right" if first[2] == "x" else "top/middle/bottom"
        span = f"{lo:.1f} mm" if lo == hi else f"{lo:.1f}-{hi:.1f} mm"
        out.append(f"{len(shapes)} shapes miss a peer's edge by {span}, e.g. "
                   f"[{first[0]}] and [{first[1]}] ({edge}) by "
                   f"{first[3] / 100:.1f} mm")
    out += _uneven_gaps(items)
    return out


def _uneven_gaps(items):
    """Runs of three or more peers in a line whose gaps differ slightly."""
    out = []
    for axis in ("x", "y"):
        lane = 3 if axis == "x" else 2           # same row: equal y; column: equal x
        pos = 2 if axis == "x" else 3
        size = 4 if axis == "x" else 5
        groups = {}
        for item in items:
            key = (item[1], round(item[lane] / ALIGNED), round(item[size] / 100))
            groups.setdefault(key, []).append(item)
        for run in groups.values():
            if len(run) < 3:
                continue
            run = sorted(run, key=lambda it: it[pos])
            gaps = [b[pos] - (a[pos] + a[size]) for a, b in zip(run, run[1:])]
            if min(gaps) < 0:
                continue
            spread = max(gaps) - min(gaps)
            if NEAR_MISS[0] <= spread <= NEAR_MISS[1]:
                out.append(
                    f"the {'row' if axis == 'x' else 'column'} "
                    f"[{run[0][0]}] … [{run[-1][0]}] is spaced "
                    + ", ".join(f"{g / 100:.1f}" for g in gaps)
                    + " mm apart -- almost even")
    return out


def alignment_issues(page, number):
    """``[(text, key)]`` -- hand-placed peers that almost line up.

    The defect a slide placed by page arithmetic actually has: rows a
    millimetre off, gaps that differ by two.  A custom diagram laid out from
    a ``Box`` lines up exactly and passes however many shapes it has.
    """
    misses = misalignments(_drawn(page))
    if not misses:
        return []
    return [(f"slide {number}: drawn shapes almost line up but do not -- "
             + "; ".join(misses)
             + ". Lay them out from one `Box` (`grid`, `cut_*`) so peers share "
               "their edges and gaps exactly (`references/drawing.md`)",
             ("alignment", number, "near-miss"))]


#: Above this share of the deck on a single design, the variety note says so
#: outright.  Four separate agent runs against the house template built
#: ``[8] * N`` -- 100%, 100%, 100% and 94% -- so the number is not a warning
#: about a tendency, it is a warning about the observed default.
DESIGN_DOMINANCE = 0.7


def _catalog_size(template):
    """How many designs a ``.pptx`` template offers, or ``None``.

    Counted out of the zip rather than by opening the document: this runs on
    every check and the answer is one number nobody should pay a LibreOffice
    load for.
    """
    if not template or Path(template).suffix.lower() not in PPTX_EXTS:
        return None
    import zipfile
    try:
        with zipfile.ZipFile(template) as z:
            return sum(1 for n in z.namelist()
                       if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
    except Exception:
        return None


def design_note(designs, catalog_size=None):
    """One line on which template designs a deck was built from, or ``None``.

    Not a finding: repeating a design is legitimate and a deck of content
    slides *should* mostly be content slides.  What is worth saying out loud is
    the share, because the runs this exists for did not notice they had built
    fifteen slides on one design and neither did anything else in the report.
    """
    if not designs:
        return None
    counts = {}
    for design in designs:
        counts[design] = counts.get(design, 0) + 1
    ranked = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
    total = len(designs)
    spread = ", ".join(f"{n}x design {d}" for d, n in ranked[:4])
    if len(ranked) > 4:
        spread += f", +{len(ranked) - 4} more"
    note = (f"{total} slide(s) on {len(counts)} design(s): {spread}")
    if ranked[0][1] / total >= DESIGN_DOMINANCE:
        note += (f" -- {ranked[0][1] / total:.0%} of the deck is design "
                 f"{ranked[0][0]}")
        if catalog_size and catalog_size > len(counts):
            note += f", out of {catalog_size} the template offers"
        note += (". Fine if they are content slides; if it is the canvas "
                 "design, the deck is drawn from scratch on a template it "
                 "never uses -- an opener, a divider and a quote are designs")
    return note


def stray_template_images(deck, template, wanted=None):
    """``[(text, (kind, slide, content))]`` -- template pictures still on slides.

    ``.pptx`` only, and read straight from the zip: a graphic is identified by
    the checksum of its bytes, which is the one property that survives a build.
    Retexting a designed slide replaces its words and leaves its pictures, so a
    deck reports clean on text while a stock photo and a German process diagram
    are still sitting on slide 7.

    The second element is the structured key an allowlist matches on:
    ``("stray-image", slide, media_filename)``.
    """
    import xml.etree.ElementTree as ET
    import zipfile

    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    try:
        with zipfile.ZipFile(template) as z:
            known = {i.CRC for i in z.infolist() if i.filename.startswith("ppt/media/")}
        deck_zip = zipfile.ZipFile(deck)
    except Exception:
        return []

    out = []
    with deck_zip as z:
        crc = {i.filename: i.CRC for i in z.infolist()}
        try:
            root = ET.fromstring(z.read("ppt/presentation.xml"))
            size = root.find(f"{P}sldSz")
            slide_area = int(size.get("cx")) * int(size.get("cy"))
        except Exception:
            return []

        # Presentation order, not filename order.  A rebuilt deck routinely
        # stores slide 7 as slide10.xml, and numbering by filename reports
        # every finding against the wrong slide.
        try:
            deck_rels = {
                r.get("Id"): r.get("Target") for r in
                ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
                .findall(f"{REL}Relationship")
            }
            names = ["ppt/" + deck_rels[s.get(f"{R}id")].replace("../", "")
                     for s in root.find(f"{P}sldIdLst")]
        except Exception:
            names = sorted(
                (n for n in z.namelist()
                 if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1)),
            )

        for number, name in enumerate(names, start=1):
            if wanted and number not in wanted:
                continue
            seen = set()
            rels_name = f"ppt/slides/_rels/{name.rsplit('/', 1)[1]}.rels"
            try:
                rels = {
                    r.get("Id"): r.get("Target")
                    for r in ET.fromstring(z.read(rels_name)).findall(f"{REL}Relationship")
                }
                slide = ET.fromstring(z.read(name))
            except Exception:
                continue

            for pic in slide.iter(f"{P}pic"):
                blip = pic.find(f".//{A}blip")
                ext = pic.find(f".//{A}ext")
                if blip is None or ext is None:
                    continue
                target = rels.get(blip.get(f"{R}embed"), "")
                media = "ppt/" + target.replace("../", "")
                if crc.get(media) not in known:
                    continue
                try:
                    area = int(ext.get("cx")) * int(ext.get("cy"))
                except Exception:
                    continue
                if area < MIN_STRAY_IMAGE * slide_area:
                    continue
                # The same picture placed twice is one thing to remove.
                if media in seen:
                    continue
                seen.add(media)
                media_name = media.rsplit('/', 1)[1]
                out.append(
                    (f"slide {number}: a picture from the template is still on "
                     f"the slide ({media_name}, "
                     f"{100 * area / slide_area:.1f}% of it) -- retext replaces "
                     f"words, never pictures; remove it or use a design "
                     f"without one",
                     ("stray-image", number, media_name))
                )
    return out


#: A deck of at least this many slides with one design on more than
#: ``VARIETY_SHARE`` of them is reported as a finding, not a note.
VARIETY_MIN_SLIDES = 6
VARIETY_SHARE = 0.5


def variety_findings(designs, catalog_size=None):
    """``[(text, key)]`` when one design carries most of a deck.

    The note above said this for months and was read past every time: two
    builds in a row drew five or seven slides of the same three cards on the
    canvas design and shipped them.  A finding gets fixed; a design that
    really should repeat goes in the allowlist with its reason.
    """
    if len(designs) < VARIETY_MIN_SLIDES:
        return []
    top = max(set(designs), key=designs.count)
    share = designs.count(top) / len(designs)
    if share <= VARIETY_SHARE:
        return []
    offered = (f" of the {catalog_size} the template offers"
               if catalog_size else "")
    text = (f"deck: {designs.count(top)} of {len(designs)} slides are design "
            f"{top}. Pick designs{offered} whose content shape fits each "
            f"slide (catalog.md, Design index); keep the canvas for the "
            f"drawings nothing else fits")
    # Keyed on the first slide that uses it, so an allowlist entry can name
    # a real slide: {"slide": <that>, "kind": "variety", "content": "<design>"}.
    return [(text, ("variety", designs.index(top) + 1, str(top)))]
