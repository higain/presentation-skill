"""
Reuse a template deck's slides: clone them, then retext them.

``duplicate_slide()`` can only insert a copy directly after its source, and
**Impress has no way to move a slide** -- ``page.Number`` is read-only,
``XDrawPages`` offers only insert/remove, and the ``.uno:MovePage*`` dispatches
silently do nothing.  So copies of a template slide can only ever form one
contiguous run beside that template, and a deck that interleaves layouts
(divider, chart, cards, divider, columns, ...) is **not buildable by
duplication at all**.

``clone_slide()`` is the way round it: it appends a blank page and rebuilds
every shape on it, so slides can be produced in any order you like.

    import sys; sys.path.insert(0, "<skill>/scripts")
    from office.impress import impress_document, slides, shapes, save_as, delete_slide
    from office.template import clone_slide, set_text

    with impress_document("template.pptx", save=False) as doc:
        tpl = slides(doc)[:9]                    # the template block
        page = clone_slide(doc, tpl[6])          # a copy of slide 7, at the end
        s = shapes(page)
        set_text(s[1], "Three of four are past their end date")
        for old in tpl:
            delete_slide(doc, old)               # drop the templates last
        save_as(doc, "out.pptx")                 # one .pptx save, see SKILL.md

Run ``scripts/probe.py template.pptx`` first: it prints the shape index,
geometry and character formatting of every slide, which is what you need to
know which ``s[i]`` to write to.
"""

import json
import re
from pathlib import Path

import uno

from office.impress import add_slide, notes_text, set_notes, shapes

# Properties clone_shape must not copy: identity, or geometry it sets itself.
# Name is copied deliberately -- duplicates are legal and keep probe.py output
# and shape_by_text() readable on a generated deck.
_SKIP = {
    "ZOrder", "LayerID", "LayerName", "Position", "Size",
    "BoundRect", "FrameRect", "Transformation", "UserDefinedAttributes",
    "InteropGrabBag", "NumberingRules", "Style", "Parent",
    # Writing ReplacementGraphic on an image kills the whole UNO bridge --
    # not an exception, the soffice process goes away.  Graphic carries the
    # picture anyway, so nothing is lost by skipping it.
    "ReplacementGraphic",
}

# Formatting that setString() throws away, and write_text() puts back.
CHAR_PROPS = (
    "CharFontName", "CharHeight", "CharWeight", "CharColor", "CharKerning",
    "CharPosture", "CharUnderline", "CharCaseMap", "CharStrikeout",
    "CharEscapement", "CharEscapementHeight", "CharContoured",
)
PARA_PROPS = (
    "ParaAdjust", "ParaLineSpacing", "ParaTopMargin", "ParaBottomMargin",
    "ParaLeftMargin", "ParaRightMargin", "ParaFirstLineIndent",
)

_BG_PROPS = ("FillStyle", "FillColor", "FillTransparence", "FillGradientName",
             "FillBitmapName", "FillHatchName")

_READONLY = 16  # com.sun.star.beans.PropertyAttribute.READONLY


# ------------------------------------------------------------ walking shapes

def _is_group(shape):
    try:
        return shape.getShapeType() == "com.sun.star.drawing.GroupShape"
    except Exception:
        return False


def _walk(container, prefix=()):
    """``(path, shape, parent)`` for every shape in ``container``, depth first.

    ``path`` is a tuple of indices: ``(3,)`` is the fourth top-level shape,
    ``(3, 1)`` the second shape inside it.  ``parent`` is the container to
    remove the shape from -- the page for a top-level shape, the group for a
    shape inside one.
    """
    try:
        count = container.getCount()
    except Exception:
        return
    for i in range(count):
        try:
            shape = container.getByIndex(i)
        except Exception:
            continue
        path = prefix + (i,)
        yield path, shape, container
        if _is_group(shape):
            yield from _walk(shape, path)


def _addr(path):
    """A readable address: ``3`` at the top level, ``'3.1'`` inside a group."""
    return path[0] if len(path) == 1 else ".".join(str(p) for p in path)


def addr_key(address):
    """Sort key for the addresses ``walk_shapes``, ``find_shapes`` and
    ``retext`` hand back -- ``3`` sorts before ``'3.1'`` before ``4``."""
    if isinstance(address, int):
        return (address,)
    return tuple(int(part) for part in str(address).split("."))


def walk_shapes(page):
    """``[(address, shape)]`` for every shape on ``page``, groups included.

    ``shapes(page)`` stops at the top level, and on a designed template that is
    routinely less than half the words on the slide: delta's phase-plan slide
    carries 14 top-level shapes and 12 more inside five groups, holding the
    phase labels and durations.  Anything that reads or rewrites a slide's text
    has to come through here, or it will report a slide clean while the
    template's own German is still sitting on it.
    """
    return [(_addr(path), shape) for path, shape, _ in _walk(page)]


# --------------------------------------------------------------------- text

def read_text(shape):
    """``[(text, char_props, para_props)]``, one entry per paragraph.

    Character formatting is taken from each paragraph's first run, which is
    all a uniformly formatted paragraph has.  A paragraph that mixes formats
    mid-line collapses to its first run's -- retext those by hand.
    """
    out = []
    enum = shape.createEnumeration()
    while enum.hasMoreElements():
        para = enum.nextElement()
        if not hasattr(para, "createEnumeration"):
            continue
        text, cprops, first = "", {}, True
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            run = runs.nextElement()
            text += run.getString()
            if first:
                cprops = {k: getattr(run, k) for k in CHAR_PROPS if hasattr(run, k)}
                first = False
        pprops = {k: getattr(para, k) for k in PARA_PROPS if hasattr(para, k)}
        out.append((text, cprops, pprops))
    return out


def write_text(shape, paras):
    """Replace a shape's text with ``[(text, char_props, para_props)]``."""
    shape.setString("\n".join(p[0] for p in paras))
    enum = shape.createEnumeration()
    i = 0
    while enum.hasMoreElements() and i < len(paras):
        para = enum.nextElement()
        if not hasattr(para, "createEnumeration"):
            continue
        _, cprops, pprops = paras[i]
        _apply(para, pprops)
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            _apply(runs.nextElement(), cprops)
        i += 1


def set_text(shape, *lines):
    """Retext a shape, keeping its paragraphs' formatting.

    ``setString()`` alone drops font, size, weight, colour and kerning back to
    the shape's defaults -- which on a template slide means losing the design.
    Pass one argument per paragraph.

    **Formatting is matched paragraph by paragraph**, and lines past the last
    template paragraph inherit that last one.  A designed body is often not
    uniform -- the house template's default content slide opens with a blue
    heading and continues in black -- and flattening everything onto the first
    paragraph's format turns that into a wall of blue.  Where every paragraph
    is formatted alike this is the same as taking the first.

    **Colour is the exception, and no code here can fix it: the .pptx export
    discards per-paragraph character colour** and writes one colour for the
    whole shape.  Measured on the house template -- three paragraphs written
    red/blue/green come back all one colour, while their per-paragraph *sizes
    and weights* survive intact.  An .odp keeps all of it.  So a heading and
    its body in different colours have to be **separate shapes** in a .pptx;
    ``office.draw.text`` makes that cheap.
    """
    paras = read_text(shape)
    if not paras:
        shape.setString("\n".join(lines))
        return
    write_text(shape, [
        (line, paras[min(i, len(paras) - 1)][1], paras[min(i, len(paras) - 1)][2])
        for i, line in enumerate(lines)
    ])


def _apply(target, props):
    for key, value in props.items():
        try:
            setattr(target, key, value)
        except Exception:
            pass


# ------------------------------------------------------------- measurement

def text_width(doc, page, text, like):
    """How wide ``text`` would be, laid out in the format of shape ``like``.

    **A shape's ``Size`` never re-measures after you retext it** -- not even
    with ``TextAutoGrowWidth``, which keeps whatever width the shape already
    had.  The only way to size a badge, pill or label to a new string is to lay
    that string out fresh, which is what this does on a throwaway shape.

    The result includes ``like``'s own left/right text insets, so it is the
    width the shape needs.  Template boxes usually carry a little slack beyond
    it; add your own if the design wants breathing room.
    """
    props = read_text(like)
    cprops = props[0][1] if props else {}
    tmp = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(tmp)
    try:
        tmp.TextWordWrap = False
        tmp.TextAutoGrowWidth = True
        tmp.TextLeftDistance = getattr(like, "TextLeftDistance", 0)
        tmp.TextRightDistance = getattr(like, "TextRightDistance", 0)
        tmp.setString(text)
        runs = tmp.createEnumeration().nextElement().createEnumeration()
        while runs.hasMoreElements():
            _apply(runs.nextElement(), cprops)
        return tmp.Size.Width
    finally:
        page.remove(tmp)


def text_overflow(doc, page, shape):
    """``(needed, available, line)`` height in 1/100 mm, or ``None``.

    Lays the shape's own text out fresh at the shape's own width on a throwaway
    box that is allowed to grow, and reports the height it wanted against the
    height it has, plus the height of a single line in the same format.

    This is the one defect visual QA catches that a text dump cannot, so it is
    also the fallback for an agent that cannot see the rendered slide.

    ``None`` when the shape holds no text, is not a text shape, or is set to
    grow its own height -- which makes overflow impossible by construction.

    Compare with care: a designed box is often sized to its text's cap height
    rather than the full line box, so a *single* line routinely measures taller
    than the shape that holds it without anything being wrong.  ``overflowing()``
    applies the threshold that accounts for this.
    """
    try:
        paras = read_text(shape)
    except Exception:
        return None
    if not any(text.strip() for text, _, _ in paras):
        return None
    if getattr(shape, "TextAutoGrowHeight", False):
        return None

    width = shape.getSize().Width
    if width <= 0:
        return None

    tmp = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(tmp)
    try:
        for prop in ("TextLeftDistance", "TextRightDistance",
                     "TextUpperDistance", "TextLowerDistance"):
            try:
                setattr(tmp, prop, getattr(shape, prop))
            except Exception:
                pass
        tmp.TextAutoGrowWidth = False
        tmp.TextWordWrap = True
        tmp.TextAutoGrowHeight = True
        resize(tmp, width=width)

        write_text(tmp, paras)
        needed = tmp.getSize().Height
        # One line in the same formatting, to judge how much of an overrun matters.
        text, cprops, pprops = paras[0]
        write_text(tmp, [("X", cprops, pprops)])
        line = tmp.getSize().Height
        return needed, shape.getSize().Height, line
    except Exception:
        return None
    finally:
        page.remove(tmp)


def overflowing(doc, page):
    """``[(address, reason)]`` for every shape whose text is visibly broken.

    Addresses are ``walk_shapes``'s: ``3`` for a top-level shape, ``'3.1'`` for
    one inside a group.

    **Text taller than its own box is not, by itself, a defect** -- and this is
    the trap that makes a naive overflow check useless.  Designed templates size
    a body box to one line and let real content flow on into the whitespace
    below; PowerPoint and Impress both draw it, nothing is clipped, and the
    slide looks exactly as intended.  Flagging that buries real defects in a
    list of things nobody should change.  (Checked against four decks built from
    one brand template: the naive rule fired 51 times, almost all benign.)

    So a shape is reported only when the overrun actually shows:

    * **off-slide** -- the text runs past the bottom edge of the slide.
    * **collision** -- it runs into another shape sitting below it.
    * **mid-word break** -- a single word is wider than the box, so it has to
      be broken across lines.  This is the narrow-box defect: axis labels like
      "BACKEND" rendering as "BAC / KEND".

    Each reason is a finished sentence, ready to print.

    **A word broken onto exactly two lines is below this check's resolution.**
    Laying text out on a throwaway shape reproduces the real layout to a few
    percent, which is not enough to separate a snug one-liner from a two-line
    break -- "QA" in a box it fills exactly measures the same as "DATA" in a box
    it just exceeds.  Flagging both would put false alarms on every short label,
    so the threshold sits above the ambiguity and catches breaks of three lines
    or more.  Marginal two-line breaks need the render.
    """
    page_height = page.Height

    # Only text can be collided *with* in a way anyone sees.  Text flowing over
    # the card or panel it sits on is how these templates are drawn.
    walked = walk_shapes(page)

    texts = []
    for i, shape in walked:
        try:
            if not " ".join(shape.getString().split()):
                continue
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        texts.append((i, pos.X, pos.Y, size.Width, size.Height))

    out = []
    for i, shape in walked:
        measured = text_overflow(doc, page, shape)
        if not measured:
            continue
        needed, available, line = measured
        if available <= 0 or line <= 0:
            continue

        pos, size = shape.getPosition(), shape.getSize()
        text = " ".join(shape.getString().split())

        wrapped = needed > line * 1.4

        # A lone word that wrapped had to be broken across lines -- there was no
        # space to break at.  Measuring the layout settles this; comparing the
        # word's width to the box does not, because a shape's text insets and a
        # label that simply overhangs its box both read as too wide while
        # rendering perfectly.
        if wrapped and text and " " not in text:
            out.append((
                i,
                f"{text[:24]!r} is broken across {max(2, round(needed / line))} "
                f"lines mid-word -- widen the shape or shorten the label"
            ))
            continue

        # Height defects.  A single line cannot be clipped: there is no second
        # line to lose, and a box set to its text's cap height measures short on
        # a slide that is perfectly fine.
        if not wrapped or needed <= available:
            continue

        bottom = pos.Y + needed
        if bottom > page_height:
            out.append((
                i,
                f"text runs {bottom - page_height} past the bottom of the slide"
            ))
            continue

        hit = next(
            (
                j for j, x, y, w, h in texts
                if j != i
                and x < pos.X + size.Width and pos.X < x + w      # overlap in x
                and y >= pos.Y + available and y < bottom          # starts in the overrun
            ),
            None,
        )
        if hit is not None:
            out.append((
                i,
                f"text overruns its box by {needed - available} and runs into "
                f"the text of shape [{hit}] below it"
            ))
    return out


def ink_box(doc, page, shape):
    """``(x, y, w, h, text)`` of the *ink* a text shape puts on the slide.

    A text box is nearly always bigger than the words in it: a caption box the
    width of a column holding two centred words paints a fifth of its own area,
    and the rest is empty.  So comparing shape boxes to find things that
    collide reports every slide ever designed.  This lays the text out at the
    shape's own width, measures the block it actually makes, and places that
    block inside the shape the way the shape's own alignment will place it.

    ``None`` for a shape with no text, one that is rotated (the rectangle no
    longer describes it), or one that cannot be measured.
    """
    try:
        text = " ".join(shape.getString().split())
    except Exception:
        return None
    if not text:
        return None
    try:
        if int(getattr(shape, "RotateAngle", 0) or 0) % 36000:
            return None
        pos, size = shape.getPosition(), shape.getSize()
    except Exception:
        return None
    if size.Width <= 0 or size.Height <= 0:
        return None

    height = size.Height
    measured = text_overflow(doc, page, shape)
    if measured:
        needed, available, line = measured
        # Only an overrun that *wrapped* is really on the slide.  A single line
        # measures taller than a box sized to its cap height without anything
        # being wrong -- and a drawn table is a column of those, so trusting the
        # raw number reports every pair of adjacent cells as a collision.  Same
        # threshold `overflowing` uses, for the same reason.
        height = (needed if needed > line * 1.4 else min(needed, available))

    width = _ink_width(doc, page, shape, size.Width)

    # Where the block sits inside the shape, per the shape's own alignment.
    x = pos.X + _align_offset(_para_adjust(shape), size.Width - width)
    y = pos.Y + _align_offset(_vert_adjust(shape), size.Height - height)
    return (x, y, max(width, 1), max(height, 1), text)


def _ink_width(doc, page, shape, box_width):
    """Width of the longest line the text lays out to, capped at the box."""
    try:
        paras = read_text(shape)
    except Exception:
        return box_width
    tmp = doc.createInstance("com.sun.star.drawing.TextShape")
    page.add(tmp)
    try:
        tmp.TextWordWrap = False
        tmp.TextAutoGrowWidth = True
        tmp.TextAutoGrowHeight = True
        widest = 0
        for text, cprops, pprops in paras:
            if not text.strip():
                continue
            write_text(tmp, [(text, cprops, pprops)])
            widest = max(widest, tmp.getSize().Width)
        return min(widest, box_width) if widest else box_width
    except Exception:
        return box_width
    finally:
        try:
            page.remove(tmp)
        except Exception:
            pass


# Both axes answer the same question -- does the block sit at the start of the
# space, in the middle of it, or at the end -- so both report it the same way.

# com.sun.star.style.ParagraphAdjust, by value: LEFT 0, RIGHT 1, BLOCK 2,
# CENTER 3, STRETCH 4.  Justified text fills the width, so its ink starts where
# left-aligned text does.
_PARA_ADJUST = {0: "start", 1: "end", 2: "start", 3: "middle", 4: "start"}


def _para_adjust(shape):
    """``'start' | 'middle' | 'end'`` across, for a shape's first paragraph."""
    try:
        enum = shape.createEnumeration()
        while enum.hasMoreElements():
            para = enum.nextElement()
            if hasattr(para, "createEnumeration"):
                return _PARA_ADJUST.get(int(para.ParaAdjust), "start")
    except Exception:
        pass
    return "start"


def _vert_adjust(shape):
    """``'start' | 'middle' | 'end'`` down.  A text shape defaults to centred."""
    try:
        return {"TOP": "start", "BOTTOM": "end"}.get(
            str(shape.TextVerticalAdjust.value), "middle")
    except Exception:
        return "middle"


def _align_offset(where, slack):
    """How far in from the leading edge the block starts."""
    slack = max(0, slack)
    return {"start": 0, "middle": slack // 2, "end": slack}[where]


# How much of the smaller block another must cover before it counts as a
# collision.  Laying text out on a throwaway shape reproduces the real layout to
# within a few percent, not exactly, so a threshold set down where that noise
# lives yields a finding that appears and disappears between identical runs.  A
# quarter of one block buried under another is well clear of it, and is what
# "printed over each other" looks like in a render.  Pairs that brush each other
# by a millimetre are below this check's resolution and need the eye.
COLLISION = 0.25

# A badge, a tick, a bullet glyph, a step number.  Templates place these on the
# corner of the tile they mark, so they overlap it by design -- six of the house
# template's own slides do it -- and reporting them buries the finding that
# matters.  Only blocks of actual words are compared.
MIN_COLLISION_CHARS = 4


def colliding(doc, page):
    """``[(address, reason)]`` for text drawn on top of other text.

    ``overflowing`` catches text that runs *out of its own box* into something
    below it.  This catches the other half, and the more common one once a
    build script starts drawing: two blocks placed at coordinates that overlap
    from the start -- a Venn label over the circle's own caption, a title long
    enough to reach the strapline beside it, a footnote under a table that grew
    a row.  Nothing overflows, nothing is off the slide, every other check
    passes, and the render shows two sentences printed through each other.

    Measured on the ink (see :func:`ink_box`), not on the boxes, or every
    designed slide reports.
    """
    # Two passes.  Measuring ink costs two UNO layout round trips per shape, so
    # the cheap shape rectangles go first: a pair whose *boxes* miss each other
    # entirely cannot have ink that meets, and on a well laid out slide that is
    # almost every pair.  Only the survivors get measured.
    candidates = []
    for address, shape in walk_shapes(page):
        try:
            text = " ".join(shape.getString().split())
            if len(text) < MIN_COLLISION_CHARS:
                continue
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if size.Width > 0 and size.Height > 0:
            # Grown by its own height above and below: text that overruns its
            # box is still on the slide, and the pair it collides with is one
            # the raw rectangles would have said could not meet.  Generous on
            # purpose -- this pass only decides what is worth measuring.
            candidates.append((address, shape,
                               (pos.X, pos.Y - size.Height,
                                size.Width, size.Height * 3)))

    touching = set()
    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a, b = candidates[i][2], candidates[j][2]
            if (min(a[0] + a[2], b[0] + b[2]) > max(a[0], b[0])
                    and min(a[1] + a[3], b[1] + b[3]) > max(a[1], b[1])):
                touching.add(i)
                touching.add(j)

    boxes = []
    for i in sorted(touching):
        address, shape, _ = candidates[i]
        ink = ink_box(doc, page, shape)
        if ink is not None:
            boxes.append((address, ink))

    out = []
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            (a_addr, a), (b_addr, b) = boxes[i], boxes[j]
            over_w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
            over_h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
            if over_w <= 0 or over_h <= 0:
                continue
            smaller = min(a[2] * a[3], b[2] * b[3])
            if not smaller or (over_w * over_h) / smaller < COLLISION:
                continue
            share = (over_w * over_h) / smaller
            out.append((
                a_addr,
                f"{a[4][:28]!r} and {b[4][:28]!r} in shape [{b_addr}] are "
                f"printed over each other -- {share:.0%} of the smaller block "
                f"is covered"
            ))
    return out


def offslide(page, tolerance=100):
    """``[(address, reason)]`` for every shape hanging off the edge of the slide.

    ``overflowing`` measures text against the shape that holds it, so it is
    blind to a shape that is simply in the wrong place: a Venn given a radius
    too large for its slot spills off the bottom of the page and every text
    check still passes.  Geometry is the other half of "is this slide on the
    slide", and it is the one a build script gets wrong when it computes
    positions from hardcoded millimetres.

    ``tolerance`` (default 1 mm) forgives the hairline bleed that rounding and
    designed templates both produce.  A template decoration that bleeds *far*
    past the trim on purpose still trips this, so ``qa.py`` baselines whatever
    the template itself does -- see ``offslide_baseline``.

    Returned amounts are exact, so a caller can baseline on them.
    """
    out = []
    for address, shape in walk_shapes(page):
        try:
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if size.Width <= 0 or size.Height <= 0:
            continue

        over = {
            "off the left edge": -(pos.X + tolerance),
            "above the top edge": -(pos.Y + tolerance),
            "off the right edge": pos.X + size.Width - page.Width - tolerance,
            "past the bottom edge": pos.Y + size.Height - page.Height - tolerance,
        }
        worst = max(over.items(), key=lambda kv: kv[1])
        if worst[1] <= 0:
            continue

        # A shape wider than the slide in both directions is a background, not
        # a mistake.  What matters is content that has walked off the page.
        if size.Width > page.Width and size.Height > page.Height:
            continue

        try:
            label = " ".join(shape.getString().split())[:32]
        except Exception:
            label = ""
        what = f"{label!r} " if label else ""
        out.append((
            address,
            f"{what}extends {int(worst[1])} {worst[0]} of the slide "
            f"-- it will be cut off"
        ))
    return out


def offslide_baseline(path, tolerance=100):
    """``{(direction, amount)}`` for the bleeds a template already has.

    The house template has a decoration on design 14 that hangs ~10 mm below
    the trim on purpose.  ``build_deck`` duplicates it unchanged, so it reappears
    in every deck built on that design with exactly the same overhang -- report
    it and the check cries wolf on work nobody did.  Keying on direction and
    exact amount is enough to recognise the duplicate without excusing a shape
    the build script really did push off the page.
    """
    from .impress import load_document, slides

    doc = load_document(str(path), read_only=True)
    try:
        seen = set()
        for page in slides(doc):
            for _, reason in offslide(page, tolerance):
                amount, direction = _bleed_key(reason)
                if amount is not None:
                    seen.add((direction, amount))
        return seen
    finally:
        doc.close(False)


def _bleed_key(reason):
    """``(amount, direction)`` parsed back out of an ``offslide`` reason."""
    hit = re.search(r"extends (\d+) (off the \w+ edge|above the top edge|"
                    r"past the bottom edge)", reason)
    return (int(hit.group(1)), hit.group(2)) if hit else (None, None)


def fit_width(doc, page, shape, slack=0):
    """Resize ``shape`` to the width of the text it currently holds.

    Anchors on its left edge and leaves its height alone.  Returns the new
    width, so a backdrop behind it can be sized to match.
    """
    width = text_width(doc, page, shape.getString(), shape) + slack
    resize(shape, width=width)
    return width


def shrink_to_fit(doc, page, shape, max_lines=None, min_size=18.0, step=0.92):
    """Shrink ``shape``'s type until its text fits the box it is in.

    Returns the point size it settled on, or ``None`` if the text already fit.

    **Reach for this on a template's title slot.** A designed headline box is
    sized to exactly the number of lines the designer drew it for, at the point
    size they chose, and a real title is longer than the placeholder every time.
    What happens then is not clipping, which you would notice: the text flows on
    past the bottom of the box, over whatever the master has drawn there.

    ``max_lines`` is the stricter question, and on a decorated divider it is the
    one that matters.  The house dividers (designs 1, 5, 7) sit their title in
    the gap between the two edges of a big white triangle -- **one** line of
    60pt fits that gap, and a second one crosses the decoration however small
    you make the type.  ``max_lines=1`` shrinks until the title is one line.
    Nothing on the page collides, so `qa.py` cannot see this: master decoration
    is not a shape on the slide.

    If the size it returns is embarrassing, the title is too long for the
    design rather than too big -- shorten it, or build the slide on a canvas.
    Adding a second text box beside the placeholder is not the fix.
    """
    def fits():
        measured = text_overflow(doc, page, shape)
        if not measured:
            return True
        needed, available, line = measured
        if max_lines is not None and line > 0:
            return needed <= line * (max_lines + 0.4)
        return needed <= available

    if fits():
        return None
    size = None
    for _ in range(24):
        paras = read_text(shape)
        current = max((p[1].get("CharHeight", 0) or 0) for p in paras)
        if not current:
            return None
        size = max(min_size, current * step)
        for _, cprops, _ in paras:
            cprops["CharHeight"] = size
        write_text(shape, paras)
        if fits() or size <= min_size:
            break
    return size


# ------------------------------------------------------------- geometry

def move(shape, x=None, y=None):
    """Reposition a shape.  The struct needs a round trip -- assigning to
    ``shape.Position.X`` silently does nothing."""
    point = shape.getPosition()
    if x is not None:
        point.X = x
    if y is not None:
        point.Y = y
    shape.setPosition(point)
    return shape


def resize(shape, width=None, height=None):
    """Resize a shape.  Same round trip as ``move``."""
    size = shape.getSize()
    if width is not None:
        size.Width = width
    if height is not None:
        size.Height = height
    shape.setSize(size)
    return shape


def set_bar(shape, baseline, height):
    """Set a chart bar's height, keeping its foot on ``baseline``.

    A bar grows upward: its ``y`` is ``baseline - height``, not its ``y``.
    Setting the height alone leaves the bar hanging from its old top edge.
    """
    height = round(height)
    resize(shape, height=height)
    move(shape, y=round(baseline) - height)


def scale_bars(bars, values, max_height=None):
    """Rescale a row of template bars to ``values``, on their shared baseline.

    Retexting a template chart is just moving rectangles -- there is no chart
    object to talk to -- and this is that arithmetic::

        bars = [shapes(page)[i] for i in (3, 5, 7, 9)]
        scale_bars(bars, [40, 48, 29, 34])

    The baseline and the tallest bar are read off the template, so the new
    chart keeps the proportions the designer drew.  ``max_height`` overrides
    what the tallest value maps to.  Negative values are not supported; a bar
    chart with a zero value gets a zero-height bar, which is what it should be.
    """
    if len(bars) != len(values):
        raise ValueError(f"{len(bars)} bars but {len(values)} values")
    if not bars:
        return
    baseline = max(b.getPosition().Y + b.getSize().Height for b in bars)
    if max_height is None:
        max_height = max(b.getSize().Height for b in bars)
    top = max(values)
    if top <= 0:
        raise ValueError("scale_bars needs at least one positive value")
    for bar, value in zip(bars, values):
        set_bar(bar, baseline, max(0, value) / top * max_height)


def fit_pill(doc, page, pill, label, slack=0):
    """Grow a pill or badge, and the label inside it, to fit the label's text.

    Padding comes from the template's own geometry -- the difference between
    the pill's width and the label's -- so nothing has to be guessed.  Both
    keep their left edge; a longer tag grows to the right.

    Returns the label's new width.
    """
    padding = pill.getSize().Width - label.getSize().Width
    width = fit_width(doc, page, label, slack)
    resize(pill, width=width + padding)
    return width


def set_char_color(shape, color):
    """Recolour every run in a shape -- for flipping a card to a dark fill.

    Set the colour on each paragraph *and* on its runs: a paragraph's own
    CharColor does not reach runs that carry their own, which is most of them
    in a template deck.
    """
    enum = shape.createEnumeration()
    while enum.hasMoreElements():
        para = enum.nextElement()
        try:
            para.CharColor = color
        except Exception:
            pass
        if not hasattr(para, "createEnumeration"):
            continue
        portions = para.createEnumeration()
        while portions.hasMoreElements():
            try:
                portions.nextElement().CharColor = color
            except Exception:
                pass
            continue
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            runs.nextElement().CharColor = color


def shape_by_text(page, prefix):
    """The first shape on ``page`` whose text starts with ``prefix``.

    Reaches inside groups, like ``retext`` and ``find_shapes``.
    """
    for _, shape in walk_shapes(page):
        try:
            if shape.getString().startswith(prefix):
                return shape
        except Exception:
            pass
    raise LookupError(f"no shape on {page.Name!r} starts with {prefix!r}")


# ------------------------------------------------------- retexting by text

def _norm(text):
    """Collapse whitespace, so a match does not hinge on line breaks."""
    return " ".join(text.split())


def _untruncate(text):
    """Drop a trailing ellipsis from a match key.

    Inventories print long text cut short.  The excerpt then gets copied back
    into a mapping with the ellipsis still attached, and the match fails on a
    string that is visibly right -- a failure models have burned whole sessions
    on, and one that costs three lines to make impossible.  A key is a
    substring anyway, so nothing is lost by trimming it.
    """
    for marker in ("...", "…"):
        if text.endswith(marker):
            return text[: -len(marker)].rstrip()
    return text


def _scan(page):
    """``({address: text}, {address: (shape, parent)})`` for the whole slide,
    groups included."""
    texts, entries = {}, {}
    for path, shape, parent in _walk(page):
        address = _addr(path)
        entries[address] = (shape, parent)
        try:
            texts[address] = _norm(shape.getString())
        except Exception:
            texts[address] = ""
    return texts, entries


def _shape_texts(page):
    """``{address: text}`` for every shape on the slide, groups included."""
    return _scan(page)[0]


def _inventory(texts):
    """The slide's own text, addressed -- so a failed match self-corrects."""
    ordered = sorted(texts.items(), key=lambda kv: addr_key(kv[0]))
    rows = [f"    [{i}] {t[:60]!r}" for i, t in ordered if t]
    empty = [str(i) for i, t in ordered if not t]
    out = ["  shapes carrying text:"] + (rows or ["    (none)"])
    if empty:
        out.append(f"  shapes without text: {', '.join(empty)}")
    if any(isinstance(i, str) for i, _ in ordered):
        out.append("  ([3.1] is a shape inside the group at [3].)")
    out.append("  Any distinctive substring of a line above works as a key.")
    return "\n".join(out)


def find_shapes(page, pattern, regex=False):
    """Every ``(address, shape)`` on ``page`` whose text matches ``pattern``.

    Substring match on whitespace-collapsed text; ``regex=True`` treats
    ``pattern`` as a case-insensitive regular expression.  Use it to hunt
    leftover template wording before shipping a deck.

    Reaches inside groups: an address of ``'3.1'`` is the second shape in the
    group at index 3.
    """
    out = []
    for address, shape in walk_shapes(page):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        if not text:
            continue
        hit = (re.search(pattern, text, re.I) if regex
               else _norm(pattern).lower() in text.lower())
        if hit:
            out.append((address, shape))
    return out


def retext(page, mapping, **kw):
    """Retext a slide by the text its shapes already carry, not by index.

    ``mapping`` goes from an excerpt of the **template's** wording to what
    should replace it::

        left = retext(page, {
            "Tagged cards for quick points": "Three of four are past due",
            "SPEED":                         "BACKEND",
            "Two shorter paragraphs read":   None,      # delete that shape
        })

    Counting z-order indices by hand is where template builds go wrong: a
    column whose body is split across two text boxes shifts every later index,
    and the mistake surfaces only as someone else's text in the wrong column.
    Matching on the words you can actually see in ``probe.py`` output removes
    the arithmetic.

    A key is matched against exact text first, then a prefix, then a substring
    -- the most specific tier that hits anything wins.  Matching nothing, or
    more than one shape, raises ``LookupError`` **before** anything is written,
    with the slide's own indexed text in the message.

    Values: a string is one paragraph, a list or tuple one entry per paragraph,
    ``""`` empties a shape but keeps the box, and ``None`` removes the shape
    from the slide -- which is what a leftover template paragraph usually wants.

    Keys reach **inside groups**, where designed templates keep a good half of
    their text; a shape there is addressed ``'3.1'`` rather than ``3``.

    A key ending in ``...`` or ``…`` has it stripped before matching, so a
    string pasted straight out of ``probe.py``'s truncated listing still finds
    its shape instead of failing on the ellipsis.

    **It changes words and nothing else** -- the template's size, weight,
    colour and alignment are the point of building on a template, so there are
    no styling keywords here.  When the new wording genuinely does not fit the
    slot the designer drew, ``shrink_to_fit(doc, page, shape)`` is the fix;
    ``set_char_color(shape, ...)`` and ``draw.restyle(shape, ...)`` change the
    rest, on a shape ``find_shapes`` hands you.

    Returns ``[(address, text)]`` for every shape left untouched that still
    carries text: the leftover-template-text check, for free.
    """
    if kw:
        raise TypeError(
            "retext() takes no styling keywords (got %s): it replaces words "
            "and keeps the template's own formatting, which is the reason to "
            "build on a template at all.\n"
            "  Too long for the slot -> shrink_to_fit(doc, page, shape, "
            "max_lines=1)\n"
            "  Different colour      -> set_char_color(shape, 0xRRGGBB)\n"
            "  Anything else         -> draw.restyle(shape, CharHeight=..., ...)\n"
            "  find_shapes(page, 'some words')[0][1] gives you the shape."
            % ", ".join(sorted(kw)))

    texts, entries = _scan(page)

    claimed = {}
    for key in mapping:
        want = _untruncate(_norm(key))
        if not want:
            raise ValueError("retext: keys must be a non-empty piece of the template's text")
        for hits in (
            [i for i, t in texts.items() if t == want],
            [i for i, t in texts.items() if t.startswith(want)],
            [i for i, t in texts.items() if want.lower() in t.lower()],
        ):
            if hits:
                break
        if not hits:
            raise LookupError(
                f"retext: no shape on {page.Name!r} carries {key!r}.\n"
                + _inventory(texts)
            )
        if len(hits) > 1:
            raise LookupError(
                f"retext: {key!r} matches {len(hits)} shapes on {page.Name!r} "
                f"(at {hits}). Use a longer, more specific excerpt.\n"
                + _inventory(texts)
            )
        address = hits[0]
        if address in claimed:
            raise LookupError(
                f"retext: {key!r} and {claimed[address]!r} both match shape "
                f"[{address}] on {page.Name!r}; give each its own excerpt.\n"
                + _inventory(texts)
            )
        claimed[address] = key

    for address, key in claimed.items():
        value = mapping[key]
        shape, parent = entries[address]
        if value is None:
            parent.remove(shape)          # the page, or the group it sits in
        elif isinstance(value, str):
            set_text(shape, value)
        else:
            set_text(shape, *value)

    return [(i, t) for i, t in sorted(texts.items(), key=lambda kv: addr_key(kv[0]))
            if i not in claimed and t]


# A pattern-collection template carries flags addressed to whoever builds from
# it -- an "example" marker, a fill-in note for a customer logo, a block of
# build instructions.  They are ordinary shapes, so nothing flags them, and they
# ship unless deleted.  Which strings those are is a fact about one template, so
# they live beside that template rather than in this module: the engine should
# not know any company's German.
_PROFILE = None
_PROFILE_PATH = (Path(__file__).resolve().parent.parent.parent
                 / "assets" / "company-template.profile.json")


def template_profile():
    """Per-template facts shipped beside the asset, or ``{}`` if there are none.

    Currently just the annotation strings.  Matched on the whole string, so a
    slide titled "Vorgehen zum Bilden eines Modells" is not mistaken for the
    "Vorgehen:" instruction block.
    """
    global _PROFILE
    if _PROFILE is None:
        try:
            _PROFILE = json.loads(_PROFILE_PATH.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _PROFILE = {}
    return _PROFILE


def strip_annotations(page, exact=None, prefixes=None, profile=None):
    """Delete a template's authoring annotations from a slide.

    Reaches inside groups -- a template's "example" flag is routinely in one.
    ``exact`` matches a shape's whole text (case-insensitively, whitespace
    collapsed); ``prefixes`` matches the start of it, for the multi-sentence
    build-instruction blocks.

    With neither given the strings come from the shipped template profile
    (``assets/company-template.profile.json``).  For another template pass
    ``profile={"exact": [...], "prefixes": [...]}``, or the two lists directly.

    Returns the texts it removed, so you can see what went.
    """
    if profile is None:
        profile = template_profile().get("annotations", {})
    exact = tuple(profile.get("exact", ())) if exact is None else exact
    prefixes = tuple(profile.get("prefixes", ())) if prefixes is None else prefixes
    if not exact and not prefixes:
        return []
    wanted = {_norm(t).lower() for t in exact}
    starts = tuple(_norm(p).lower() for p in prefixes)
    removed = []
    for _, shape, parent in list(_walk(page)):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        if not text:
            continue
        low = text.lower()
        if low in wanted or low.startswith(starts):
            try:
                parent.remove(shape)
            except Exception:
                continue
            removed.append(text)
    return removed


# ---------------------------------------------------------------- cloning

def copy_background(doc, src_page, dst_page):
    """Copy a slide's background fill.

    A slide made by ``add_slide()`` does **not** inherit the background of the
    slide you are copying -- it comes out white.  Impress needs a fresh
    ``drawing.Background`` instance; assigning the source's own does nothing.
    """
    src_bg = src_page.Background
    if src_bg is None:
        return None
    bg = doc.createInstance("com.sun.star.drawing.Background")
    for key in _BG_PROPS:
        try:
            bg.setPropertyValue(key, src_bg.getPropertyValue(key))
        except Exception:
            pass
    dst_page.Background = bg
    return bg


def _copy_table(src, new):
    """Copy a table's shape: row and column counts, then each cell's text."""
    old, fresh = src.Model, new.Model
    rows, cols = old.Rows.Count, old.Columns.Count
    if fresh.Rows.Count < rows:
        fresh.Rows.insertByIndex(fresh.Rows.Count, rows - fresh.Rows.Count)
    if fresh.Columns.Count < cols:
        fresh.Columns.insertByIndex(fresh.Columns.Count, cols - fresh.Columns.Count)
    for band, count in ((("Columns", "Width"), cols), (("Rows", "Height"), rows)):
        axis, prop = band
        for i in range(count):
            try:
                setattr(getattr(fresh, axis).getByIndex(i), prop,
                        getattr(getattr(old, axis).getByIndex(i), prop))
            except Exception:
                pass
    for r in range(rows):
        for c in range(cols):
            cell = old.getCellByPosition(c, r)
            target = fresh.getCellByPosition(c, r)
            paras = read_text(cell)
            if paras:
                write_text(target, paras)
            for key in ("FillStyle", "FillColor"):
                try:
                    target.setPropertyValue(key, cell.getPropertyValue(key))
                except Exception:
                    pass


_CHAR_PROPS = (
    "CharColor", "CharHeight", "CharFontName", "CharWeight", "CharPosture",
    "CharUnderline", "CharStrikeout", "CharContoured", "CharShadowed",
    "CharKerning", "CharEscapement", "CharCaseMap",
)

_TEXT_LAYOUT_PROPS = (
    "TextHorizontalAdjust", "TextVerticalAdjust", "TextAutoGrowHeight",
    "TextAutoGrowWidth", "TextWordWrap", "TextFitToSize",
    "TextLeftDistance", "TextRightDistance",
    "TextUpperDistance", "TextLowerDistance",
    "FillStyle", "FillColor", "LineStyle", "TextRotation",
)


def _is_placeholder(shape):
    """True for a presentation placeholder (title, outline, subtitle, ...)."""
    try:
        return shape.getShapeType().startswith("com.sun.star.presentation.")
    except Exception:
        return False


def _needs_flattening(shape):
    """True if cloning this placeholder as a placeholder would lose its colour.

    The .pptx export drops direct run formatting on a placeholder and writes
    the colour inherited from its presentation style instead.  That is
    invisible when a run's colour already *is* the style's -- ordinary black
    body text on a light master -- and destructive when it is not, which is
    exactly the case for white text on a dark divider.

    So compare the run against its own style: any run that overrides the style
    carries formatting the export cannot preserve, and must be rebuilt as a
    plain text shape.  Note the master is not part of this test -- a cloned
    placeholder loses the override whether or not the master changed.
    """
    if not _is_placeholder(shape) or not shape.getString():
        return False
    try:
        style_color = shape.Style.CharColor
    except Exception:
        return True     # no style to inherit from; keep the formatting we have
    for para in _paragraphs(shape):
        portion = _first_portion(para)
        if portion is None:
            continue
        try:
            if portion.CharColor != style_color:
                return True
        except Exception:
            continue
    return False


def _paragraphs(shape):
    """Every paragraph of a text-bearing shape, as a list."""
    out = []
    if not hasattr(shape, "createEnumeration"):
        return out
    enum = shape.createEnumeration()
    while enum.hasMoreElements():
        out.append(enum.nextElement())
    return out


def _first_portion(para):
    """The paragraph's first text run, or None for an empty paragraph."""
    if not hasattr(para, "createEnumeration"):
        return None
    enum = para.createEnumeration()
    return enum.nextElement() if enum.hasMoreElements() else None


def flatten_placeholder(doc, src, dst_page):
    """Rebuild a presentation placeholder as a plain ``drawing.TextShape``.

    Cloning a placeholder onto a page whose master differs from the source's
    silently loses its text colour.  The write does land -- read the portion
    back in the same session and it is the colour you set -- but LibreOffice's
    .pptx export discards direct run formatting on a *placeholder* and
    re-inherits the destination master's text style instead.  Measured on a
    dark section divider: the run leaves as ``<a:schemeClr val="dk1"/>``
    (black) where the source had ``lt1``, and reloading reports ``0x0``.
    Reassigning ``MasterPage``, rebinding ``Style`` to the source master's
    family, and forcing a non-theme colour all fail to prevent it.

    A plain text shape has no such inheritance, so the colour exports as an
    explicit ``<a:srgbClr val="ffffff"/>`` and survives.  The cost is that the
    shape stops being a placeholder: it no longer picks up later layout
    changes, and per-run formatting *within* a paragraph is flattened to that
    paragraph's first run.
    """
    new = doc.createInstance("com.sun.star.drawing.TextShape")
    dst_page.add(new)
    new.setPosition(src.getPosition())
    new.setSize(src.getSize())

    for prop in _TEXT_LAYOUT_PROPS:
        try:
            new.setPropertyValue(prop, src.getPropertyValue(prop))
        except Exception:
            pass

    new.setString(src.getString())

    # setString() collapses each paragraph to a single run, so carry the
    # source paragraph's first run's character formatting across the whole
    # of the matching destination paragraph.
    for src_para, dst_para in zip(_paragraphs(src), _paragraphs(new)):
        portion = _first_portion(src_para)
        if portion is None:
            continue
        for prop in _CHAR_PROPS:
            try:
                dst_para.setPropertyValue(prop, portion.getPropertyValue(prop))
            except Exception:
                pass
        for prop in ("ParaAdjust", "ParaLeftMargin", "ParaTopMargin",
                     "ParaBottomMargin", "ParaLineSpacing"):
            try:
                dst_para.setPropertyValue(prop, src_para.getPropertyValue(prop))
            except Exception:
                pass
    return new


def clone_shape(doc, src, dst_page):
    """Rebuild one shape on another slide, formatting and all.

    ``dst_page`` may be a slide or a group shape.
    """
    new = doc.createInstance(src.ShapeType)
    dst_page.add(new)

    if src.getShapeType() == "com.sun.star.drawing.TableShape":
        # Sweeping a table's properties crashes the UNO bridge outright, so
        # copy geometry and cells only.
        new.setPosition(src.getPosition())
        new.setSize(src.getSize())
        _copy_table(src, new)
        return new

    if src.getShapeType() == "com.sun.star.drawing.GroupShape":
        # A group's own properties are derived from its children, so clone
        # those into it and stop -- setting Position/Size on it would fight
        # the members that define it.
        for i in range(src.getCount()):
            clone_shape(doc, src.getByIndex(i), new)
        return new

    info = src.getPropertySetInfo()
    for prop in info.getProperties():
        if prop.Name in _SKIP or prop.Attributes & _READONLY:
            continue
        try:
            new.setPropertyValue(prop.Name, src.getPropertyValue(prop.Name))
        except Exception:
            pass

    # CustomShapeGeometry must land AFTER that sweep.  Set inside it, only
    # AdjustmentValues survives: the shape loses its Path and ViewBox and
    # renders as nothing at all -- invisible, with no error anywhere.
    try:
        new.CustomShapeGeometry = src.CustomShapeGeometry
    except Exception:
        pass

    # getPosition/getSize, not the properties: a TableShape has no Position
    # property at all and raises UnknownPropertyException for it.
    new.setPosition(src.getPosition())
    new.setSize(src.getSize())

    try:
        paras = read_text(src)
        if any(text for text, _, _ in paras):
            write_text(new, paras)
    except Exception:
        pass
    return new


def clone_slide(doc, src_page, index=None):
    """Copy a template slide to a new page, with its background and notes.

    ``index`` defaults to the end of the deck, and that is how to use it:
    **clone in final slide order and leave ``index`` alone.**  Unlike
    ``duplicate_slide()`` the copy can go anywhere, which is the whole point
    (see the module docstring) -- but inserting at a position while the
    template slides are still in the deck shifts those too, and what ends up
    where stops being predictable.  Clone order becomes deck order once the
    templates are deleted, which is all the control a build needs.

    Shapes are rebuilt property by property.  Measured pixel-identical to the
    source on a fixture carrying text, an image, a table and a group; for a
    slide with an OLE chart or embedded media, ``duplicate_slide()`` is the
    safer copy when slide order happens to allow it.

    **Cross-master clones.**  A new page comes up on the deck's default master,
    so cloning a slide that sits on a *different* master -- a dark section
    divider in a mostly-light template, say -- would otherwise lose both its
    background and its text colour.  The master is carried over here, and any
    presentation placeholder is flattened to a plain text shape so its colour
    survives the .pptx export (see ``flatten_placeholder`` for why that is
    necessary).  Slides already on the default master are untouched by this and
    keep their placeholders.
    """
    dst = add_slide(doc, index)
    # A new page comes up on whatever master its neighbour uses, which is not
    # necessarily the source's -- carry the source's over before anything else,
    # or a dark divider lands on the light master and loses its background.
    try:
        dst.MasterPage = src_page.MasterPage
    except Exception:
        pass

    copy_background(doc, src_page, dst)
    for shape in shapes(src_page):
        if _needs_flattening(shape):
            flatten_placeholder(doc, shape, dst)
        else:
            clone_shape(doc, shape, dst)
    try:
        notes = notes_text(src_page)
        if notes:
            set_notes(dst, notes)
    except Exception:
        pass
    return dst
