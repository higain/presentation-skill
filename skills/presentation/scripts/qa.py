#!/usr/bin/env python3
"""
QA a finished deck without looking at it: overflow, leftovers, file integrity.

    python3 scripts/qa.py deck.pptx
    python3 scripts/qa.py deck.pptx --template Kite_brand_master.pptx
    python3 scripts/qa.py deck.odp --slides 3,7-9

The checks, in the order they catch real defects:

1. **Text overflow** -- every text shape is re-laid out at its own width and the
   height it wants is compared with the height it has.  This is the one thing
   the renders show that a text dump does not, so it is what makes the deck
   checkable when you cannot see the slides.
2. **Collision** -- two blocks of text printed over each other.  Not an
   overrun: both fit their own boxes, and the boxes were placed on top of one
   another to begin with.  This is what hand-placing labels around a diagram
   produces, and until it was measured every check here passed on decks whose
   renders showed two sentences through each other.
3. **Off the slide** -- a shape whose geometry hangs past the trim, which is
   what a diagram sized by eye does while every text check still passes.
4. **Leftover text** -- generic placeholders (lorem, TODO, xxx, [insert]), and,
   with ``--template``, any string still shared verbatim with the template you
   built from.  Retexting a designed slide by z-order index puts text in the
   wrong box and leaves the template's own words behind; this finds that.
5. **Markdown on a slide** -- ``~~struck~~`` and ``**bold**`` render as literal
   tildes and asterisks.  They arrive by pasting drafted wording into
   ``set_text`` unchanged.
6. **Contrast** -- each run's colour against the fill behind it, or the slide
   background.  A headline drawn in the default black onto a dark divider is
   the right words in the right box and invisible, and nothing else here sees
   it.
7. **Occlusion** -- an opaque shape covering the one beneath it, which is what
   a ``cut_*`` pair unpacked in the wrong order looks like on the slide.
8. **Stray template pictures** -- with ``--template``, a picture from the
   template still sitting on a slide.  ``retext`` replaces words and never
   pictures, so a design's stock photo survives a deck that reports clean.
9. **File integrity** -- runs ``office/validate.py`` on a ``.pptx``.  ODF needs
   no equivalent: LibreOffice writes it itself.

Exit codes: 0 clean, 1 issues found, 2 bad usage or unreadable file.

**This still does not replace looking at the slides.** It cannot see that a
diagram is off-centre, that a slide is half empty, or that a Venn's circles do
not overlap.  Render and look whenever you can -- see the QA section of
SKILL.md.
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import (  # noqa: E402
    PRESENTATION_EXTS, SofficeError, load_document, parse_slide_spec, shapes,
    slides,
)
from office.deckbuild import find_template  # noqa: E402
from office.draw import contrast_ratio  # noqa: E402
from office.template import (  # noqa: E402
    _bleed_key, colliding, ink_box, offslide, offslide_baseline, overflowing,
    walk_shapes,
)

PPTX_EXTS = {".pptx", ".potx"}

PLACEHOLDER = re.compile(
    r"\b(lorem|ipsum|todo|tbd|fixme|platzhalter|blindtext)\b"
    r"|\bx{3,}\b"
    r"|\[\s*(insert|insert [^\]]*|your [^\]]*|name|date|company)\s*\]",
    re.I,
)

# Text this long, still verbatim from the template, counts as leftover wording.
#
# This was 25 -- long enough that only a body paragraph could trip it -- on the
# theory that a designed deck reuses its template's short headers on purpose
# ("Agenda", "What we found") and flagging those would bury the real defect.
# Two runs measured against the delta template say otherwise: the label
# `"Navigator (1/2)"` -- 15 characters, a documented retext key of design 8 --
# survived on 13 of 15 slides in one deck and 9 of 10 in the other, in German,
# on a finished English deck, and both reports said clean.  Nothing else in
# either deck came within nine characters of the old threshold.
#
# At 6 the only other template strings either deck still held were the single
# digits "1" to "4", which is why a leftover must also contain a letter: a
# numbered step that happens to match a template's numbered step is a
# coincidence, not a leftover.
MIN_LEFTOVER = 6

# Markdown that reached a slide.  A slide is not a document: these render as
# literal tildes and asterisks, and they only ever arrive because the wording
# was drafted in markdown and pasted into set_text unchanged.
MARKDOWN = re.compile(r"~~.+?~~|\*\*.+?\*\*|(?<![\w*])\*[^*\s][^*]*\*(?![\w*])"
                      r"|(?:^|\n)\s{0,3}#{1,6}\s+\S")

# Below this WCAG ratio, text on its backdrop is unreadable at any size a slide
# uses.  4.5 is the AA threshold for body text; 3.0 is where a headline stops
# being legible at all, and it is the number worth failing a build over.
MIN_CONTRAST = 3.0

# Below this the text is not merely hard to read, it is not visible: black on
# navy, white on lime.  Always a mistake, never a style.
INVISIBLE = 1.8

# How much of a shape another one must cover before it counts as hiding it.
# A caption sitting on a panel overlaps it; a fill that swallows four fifths of
# a panel replaced it, which is what a mis-ordered cut_* produces.
OCCLUSION = 0.75

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


def _parse_slides(spec):
    return set(parse_slide_spec(spec))


def _norm(text):
    return " ".join(text.split())


def _page_texts(page):
    """``(address, text)`` for every shape on the page that carries text.

    Walks groups.  Stopping at the top level is how a deck gets reported clean
    with the template's own wording still on it: a designed slide keeps much of
    its text inside groups, and none of it is visible to ``shapes(page)``.
    """
    for address, shape in walk_shapes(page):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        if text:
            yield address, text


class _Baseline:
    """What the template already does, so a deck is not blamed for inheriting it.

    Same principle as ``offslide_baseline``: a brand template that sets white
    on orange, or stacks two shapes on its own design 17, hands both to every
    deck built from it.  Reporting them is reporting work nobody did, and a
    check that does that gets ignored -- including on the lines that matter.

    Colour pairs are baselined whatever their ratio.  A pairing the template
    ships is a house style decision you inherit, not a defect this build
    introduced; a pairing it does not ship -- black on a navy divider -- is
    yours, and still fires.
    """

    def __init__(self, strings=(), bleeds=(), contrast=(), occlusion=(),
                 collisions=()):
        self.strings = set(strings)
        self.bleeds = set(bleeds)
        self.contrast = set(contrast)
        self.occlusion = set(occlusion)
        self.collisions = set(collisions)


def _collision_key(reason):
    """The two texts a collision names, order-independent."""
    return frozenset(re.findall(r"'((?:[^']|\\')*)'", reason)[:2])


def _template_baseline(path):
    """Everything the template already does, in one document load."""
    if not path:
        return _Baseline()

    doc = load_document(path, read_only=True)
    try:
        strings, contrast, occlusion, collisions = set(), set(), set(), set()
        for page in slides(doc):
            for _, text in _page_texts(page):
                if len(text) >= MIN_LEFTOVER and any(c.isalpha() for c in text):
                    strings.add(text)
            contrast.update(pair for pair, *_ in _contrast_findings(page))
            occlusion.update(pair for pair, *_ in _occlusion_findings(page))
            collisions.update(_collision_key(reason)
                              for _, reason in colliding(doc, page))
    finally:
        doc.close(False)

    return _Baseline(strings, offslide_baseline(path), contrast, occlusion,
                     collisions)


def _geometry(shape):
    """``(x, y, w, h)`` or ``None`` for a shape that will not report it."""
    try:
        pos, size = shape.getPosition(), shape.getSize()
    except Exception:
        return None
    if size.Width <= 0 or size.Height <= 0:
        return None
    return pos.X, pos.Y, size.Width, size.Height


def _is_solid(style):
    """``FillStyle`` is a UNO enum; compare on its name, not its repr."""
    return getattr(style, "value", None) == "SOLID" or str(style).endswith("SOLID")


def _solid_fill(shape, max_transparence=40):
    """The shape's fill colour if it is opaque enough to hide what is under it."""
    try:
        if not _is_solid(shape.FillStyle):
            return None
        if int(shape.FillTransparence) > max_transparence:
            return None
        return int(shape.FillColor) & 0xFFFFFF
    except Exception:
        return None


def _page_bg(page):
    for holder in (page, getattr(page, "MasterPage", None)):
        if holder is None:
            continue
        try:
            background = holder.Background
            if background is None:
                continue
            if _is_solid(background.FillStyle):
                return int(background.FillColor) & 0xFFFFFF
        except Exception:
            continue
    return None


def _text_colors(shape):
    """``{colour: sample text}`` for every run in the shape that carries one.

    Reading ``shape.CharColor`` is not enough and usually returns -1
    ("automatic"): once a shape holds text the runs carry their own hard
    formatting, which is where ``draw.restyle`` and every template put it.
    """
    out = {}
    try:
        paragraphs = shape.createEnumeration()
    except Exception:
        return out
    while paragraphs.hasMoreElements():
        para = paragraphs.nextElement()
        if not hasattr(para, "createEnumeration"):
            continue
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            run = runs.nextElement()
            try:
                sample, color = run.getString().strip(), int(run.CharColor)
            except Exception:
                continue
            if sample and color >= 0:            # -1 is "automatic"
                out.setdefault(color & 0xFFFFFF, sample)
    return out


def _covers(outer, inner, share):
    """Does ``outer`` cover at least ``share`` of ``inner``'s area?"""
    ox, oy, ow, oh = outer
    ix, iy, iw, ih = inner
    over_w = min(ox + ow, ix + iw) - max(ox, ix)
    over_h = min(oy + oh, iy + ih) - max(oy, iy)
    if over_w <= 0 or over_h <= 0:
        return False
    return (over_w * over_h) >= share * (iw * ih)


def _contains(outer, x, y):
    ox, oy, ow, oh = outer
    return ox <= x <= ox + ow and oy <= y <= oy + oh


def _contrast_findings(page):
    """``((fg, bg), address, sample, ratio)`` for text that does not read.

    The backdrop is the shape's own fill, else the nearest opaque shape below
    it in z-order that contains its centre, else the slide background.
    """
    out = []
    slide_bg = _page_bg(page)
    tops = list(shapes(page))
    boxes = [(_geometry(s), _solid_fill(s)) for s in tops]

    for address, shape in walk_shapes(page):
        colors = _text_colors(shape)
        geometry = _geometry(shape)
        if not colors or geometry is None:
            continue

        backdrop = _solid_fill(shape)
        if backdrop is None:
            cx = geometry[0] + geometry[2] // 2
            cy = geometry[1] + geometry[3] // 2
            head = str(address).split(".")[0]
            index = int(head) if head.isdigit() else len(tops)
            for i in range(min(index, len(boxes)) - 1, -1, -1):
                other, fill = boxes[i]
                if fill is not None and other and _contains(other, cx, cy):
                    backdrop = fill
                    break
        if backdrop is None:
            backdrop = slide_bg
        if backdrop is None:
            continue

        for color, sample in colors.items():
            ratio = contrast_ratio(color, backdrop)
            if ratio < MIN_CONTRAST:
                out.append(((color, backdrop), address, sample, ratio))
    return out


def contrast_issues(page, number, baseline=frozenset()):
    """Text whose colour does not read on whatever is behind it.

    A headline drawn in the default black onto a dark divider passes every
    other check in this file -- it is the right shape, in the right place, with
    the right words, and invisible.
    """
    out = []
    for pair, address, sample, ratio in _contrast_findings(page):
        if pair in baseline:
            continue
        color, backdrop = pair
        # Two different defects, and conflating them is how a check earns the
        # habit of being ignored.  Under 1.8 the text is not there at all -- a
        # wrong colour, always a bug.  Above it the text is legible up close
        # and not from the back of a room, which is usually a brand colour at a
        # size it was not meant for: a judgement call, and the caller's to make.
        verdict = ("invisible -- wrong colour for this backdrop"
                   if ratio < INVISIBLE else
                   "too low to read from a seat -- fine for a rule or a "
                   "fill, not for words")
        out.append(
            f"slide {number} shape [{address}]: {sample[:40]!r} is "
            f"#{color:06X} on #{backdrop:06X} -- contrast {ratio:.1f}:1, "
            f"{verdict} (want {MIN_CONTRAST:.0f}:1)"
        )
    return out


def _occlusion_findings(page):
    """``((top_fill, low_fill), above, below)`` for a shape hiding an earlier one."""
    out = []
    tops = list(shapes(page))
    entries = []
    for shape in tops:
        geometry = _geometry(shape)
        try:
            has_text = bool(_norm(shape.getString()))
        except Exception:
            has_text = False
        entries.append((geometry, _solid_fill(shape), has_text))

    for above in range(len(entries) - 1, 0, -1):
        top_geo, top_fill, _ = entries[above]
        if top_geo is None or top_fill is None:
            continue
        for below in range(above):
            low_geo, low_fill, low_text = entries[below]
            if low_geo is None or (low_fill is None and not low_text):
                continue
            # A backdrop deliberately larger than what sits on it is fine;
            # only a shape that swallows a smaller one is the defect.
            if low_geo[2] * low_geo[3] > 4 * top_geo[2] * top_geo[3]:
                continue
            if _covers(top_geo, low_geo, OCCLUSION):
                out.append(((top_fill, low_fill), above, below))
                break
    return out


def occlusion_issues(page, number, baseline=frozenset()):
    """An opaque shape drawn on top of, and hiding, an earlier one.

    This is what a ``cut_*`` unpacked the wrong way round looks like on the
    slide: the band and the remainder swap, and a fill meant for a 20 mm strip
    covers the whole panel.  Nothing is off the slide and no text overflows,
    so this is the only check that sees it.
    """
    out = []
    for pair, above, below in _occlusion_findings(page):
        if pair in baseline:
            continue
        out.append(
            f"slide {number} shape [{above}]: an opaque #{pair[0]:06X} fill "
            f"covers {OCCLUSION:.0%}+ of shape [{below}] beneath it -- if that "
            f"was meant to be a band, a cut_* pair came back in the other order"
        )
    return out


def markdown_issues(page, number):
    """Markdown syntax that will render literally on the slide."""
    out = []
    for address, text in _page_texts(page):
        hit = MARKDOWN.search(text)
        if hit:
            out.append(
                f"slide {number} shape [{address}]: markdown "
                f"{hit.group(0)[:30]!r} will render as literal characters -- "
                f"style the run instead (weight=BOLD, italic=True)"
            )
    return out


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
        try:
            width, height = page.Width, page.Height
        except Exception:
            continue
        if width <= 0 or height <= 0:
            continue

        boxes = _painted_boxes(doc, page)
        if not boxes:
            out.append((number, "nothing on it at all"))
            continue

        # Overlapping shapes are counted twice, so this runs over 100% on a
        # busy slide.  That is harmless and deliberately not corrected: the
        # question is only ever "is this under a quarter", and double-counting
        # can push a slide *out* of that, never into it.
        area = sum(w * h for _, _, w, h in boxes) / (width * height)
        bands = [(y, y + h) for _, y, _, h in boxes]
        band, at = _widest_empty_band(bands, min(t for t, _ in bands), height)

        if area < SPARSE_AREA:
            out.append((number, f"paints {area:.0%} of the slide"))
        elif band >= SPARSE_BAND * height:
            out.append((number, f"{band / height:.0%} of its height is one "
                                f"empty band, from {at / height:.0%} down"))
    return out


def stray_template_images(deck, template, wanted=None):
    """Pictures still on a slide that came from the template, with their size.

    ``.pptx`` only, and read straight from the zip: a graphic is identified by
    the checksum of its bytes, which is the one property that survives a build.
    Retexting a designed slide replaces its words and leaves its pictures, so a
    deck reports clean on text while a stock photo and a German process diagram
    are still sitting on slide 7.
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
                out.append(
                    f"slide {number}: a picture from the template is still on "
                    f"the slide ({media.rsplit('/', 1)[1]}, "
                    f"{100 * area / slide_area:.1f}% of it) -- retext replaces "
                    f"words, never pictures; remove it or use a design "
                    f"without one"
                )
    return out


def check(path, template=None, wanted=None):
    """``(issues, slides_checked, sparse)`` -- all three ready to print.

    ``sparse`` is advisory and deliberately not part of ``issues``: see
    ``sparse_slides``.  A caller that fails a build on findings must not fail
    it on these.
    """
    issues = []
    leftovers = {}
    # Whatever the template already does -- its wording, the decorations it
    # bleeds past the trim on purpose, its brand colour pairings, the shapes it
    # stacks -- is duplicated into the deck unchanged, and is not a defect this
    # build introduced.  One document load for all four.
    base = _template_baseline(template)
    tpl_strings, tpl_bleeds = base.strings, base.bleeds

    doc = load_document(path, read_only=True)
    try:
        pages = slides(doc)
        checked = 0
        for number, page in enumerate(pages, start=1):
            if wanted and number not in wanted:
                continue
            checked += 1

            for address, reason in overflowing(doc, page):
                issues.append(f"slide {number} shape [{address}]: {reason}")

            for address, reason in colliding(doc, page):
                if _collision_key(reason) in base.collisions:
                    continue
                issues.append(f"slide {number} shape [{address}]: {reason}")

            # One table drawn too wide is 11 shapes over the edge by the
            # same amount.  Report the defect once, with a count -- eleven
            # lines for one mistake reads as eleven mistakes.
            bleeds = {}
            for address, reason in offslide(page):
                amount, direction = _bleed_key(reason)
                # Round: the same object's cells differ by a unit of rounding.
                key = (round(amount / 50) * 50 if amount else amount, direction)
                if (direction, amount) in tpl_bleeds:
                    continue
                bleeds.setdefault(key, []).append((address, reason))
            for (amount, direction), group in sorted(
                    bleeds.items(), key=lambda kv: -kv[0][0]):
                address, reason = group[0]
                more = (f" (and {len(group) - 1} more shape(s) by the same amount"
                        f" -- one object drawn too wide)" if len(group) > 1 else "")
                issues.append(f"slide {number} shape [{address}]: {reason}{more}")

            for address, text in _page_texts(page):
                hit = PLACEHOLDER.search(text)
                if hit:
                    issues.append(
                        f"slide {number} shape [{address}]: placeholder text "
                        f"{hit.group(0)!r} in {text[:60]!r}"
                    )
                elif text in tpl_strings:
                    leftovers.setdefault(text, []).append((number, address))

            issues.extend(markdown_issues(page, number))
            issues.extend(contrast_issues(page, number, base.contrast))
            issues.extend(occlusion_issues(page, number, base.occlusion))

        sparse = sparse_slides(doc, wanted)
    finally:
        doc.close(False)

    # One label nobody replaced, on thirteen slides, is one thing to fix.  It
    # is also the single most repeated defect these runs produce, so it is
    # reported with every slide it is on: the fix is one retext key, and the
    # list is what makes that obvious rather than thirteen near-identical
    # lines that read as thirteen separate mistakes.
    for text, places in sorted(leftovers.items(), key=lambda kv: -len(kv[1])):
        numbers = sorted({number for number, _ in places})
        where = (f"slide {numbers[0]}" if len(numbers) == 1
                 else f"slides {', '.join(str(n) for n in numbers)}")
        issues.append(
            f"{where} shape [{places[0][1]}]: still the template's own wording "
            f"-- {text[:60]!r}. Replace it with `retext`, or draw over the "
            f"design that carries it"
        )

    # Read from the zip, so it runs after the document is closed.
    if template and Path(path).suffix.lower() in PPTX_EXTS \
            and Path(template).suffix.lower() in PPTX_EXTS:
        issues.extend(stray_template_images(path, template, wanted))

    return issues, checked, sparse


def validate_pptx(path, template=None):
    """Run ``office/validate.py``, baselining against the template when given.

    ``--original`` suppresses schema findings the template already had, so what
    is left is yours.  It does not cover broken relationships or undeclared
    content types -- those are repaired at the source, in ``impress.save_as``.
    """
    script = Path(__file__).resolve().parent / "office" / "validate.py"
    cmd = [sys.executable, str(script), str(path)]
    if template and Path(template).suffix.lower() in PPTX_EXTS:
        cmd += ["--original", str(template)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    return proc.returncode, (proc.stdout + proc.stderr).strip()


def main():
    parser = argparse.ArgumentParser(
        description="Check a deck for overflow, leftover text and file defects"
    )
    parser.add_argument("input", help="Deck to check")
    parser.add_argument(
        "--template",
        help="The deck this was built from; flags text still shared with it",
    )
    parser.add_argument("--slides", help="Only these slides, e.g. 4,7-9")
    parser.add_argument(
        "--skip-validate", action="store_true",
        help="Skip the .pptx schema check",
    )
    args = parser.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"Error: {path} does not exist", file=sys.stderr)
        return 2
    if path.suffix.lower() not in PRESENTATION_EXTS:
        print(
            f"Error: {path} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        return 2
    # Same resolver build_deck uses, for the same reason: the template path in
    # the documentation is written relative to the skill, and a build script
    # runs next to the deck.
    if args.template:
        try:
            args.template = str(find_template(args.template))
        except FileNotFoundError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    wanted = _parse_slides(args.slides) if args.slides else None

    try:
        issues, checked, sparse = check(path, args.template, wanted)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    print(f"# QA {path.name}")
    print(f"\n{checked} slide(s) checked.\n")

    if issues:
        print(f"## {len(issues)} issue(s)\n")
        for issue in issues:
            print(f"- {issue}")
    else:
        print("No overflow, collisions, leftovers, markdown, contrast or "
              "occlusion defects.")

    if sparse:
        print(f"\n## {len(sparse)} slide(s) to look at\n")
        for number, reason in sparse:
            print(f"- slide {number}: {reason}")
        print(
            "\nNot findings -- a statement slide is sparse on purpose. These "
            "are the slides where that is worth confirming."
        )

    failed = bool(issues)

    if path.suffix.lower() in PPTX_EXTS and not args.skip_validate:
        code, output = validate_pptx(path, args.template)
        print("\n## File integrity\n")
        if code == 0:
            print("validate.py: passed")
        else:
            failed = True
            print(f"validate.py: FAILED (exit {code})")
            print(output)
            if not args.template:
                print(
                    "\nIf this deck came from a template, re-run with "
                    "`--template <template>` -- the template's own schema "
                    "defects otherwise read as yours."
                )

    print(
        "\nThis check cannot see the slides. Render them "
        "(`python3 scripts/render.py`) and look, whenever you can."
    )
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
