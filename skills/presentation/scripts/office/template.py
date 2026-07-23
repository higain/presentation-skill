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

import re

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
    """Retext a shape, keeping the formatting of its first paragraph.

    ``setString()`` alone drops font, size, weight, colour and kerning back to
    the shape's defaults -- which on a template slide means losing the design.
    Pass one argument per paragraph.
    """
    paras = read_text(shape)
    if not paras:
        shape.setString("\n".join(lines))
        return
    _, cprops, pprops = paras[0]
    write_text(shape, [(line, cprops, pprops) for line in lines])


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
    """``[(index, reason)]`` for every shape whose text is visibly broken.

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
    """
    page_height = page.Height

    # Only text can be collided *with* in a way anyone sees.  Text flowing over
    # the card or panel it sits on is how these templates are drawn.
    texts = []
    for i, shape in enumerate(shapes(page)):
        try:
            if not " ".join(shape.getString().split()):
                continue
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        texts.append((i, pos.X, pos.Y, size.Width, size.Height))

    out = []
    for i, shape in enumerate(shapes(page)):
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


def fit_width(doc, page, shape, slack=0):
    """Resize ``shape`` to the width of the text it currently holds.

    Anchors on its left edge and leaves its height alone.  Returns the new
    width, so a backdrop behind it can be sized to match.
    """
    width = text_width(doc, page, shape.getString(), shape) + slack
    resize(shape, width=width)
    return width


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


def set_char_color(shape, color):
    """Recolour every run in a shape -- for flipping a card to a dark fill."""
    enum = shape.createEnumeration()
    while enum.hasMoreElements():
        para = enum.nextElement()
        if not hasattr(para, "createEnumeration"):
            para.CharColor = color
            continue
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            runs.nextElement().CharColor = color


def shape_by_text(page, prefix):
    """The first shape on ``page`` whose text starts with ``prefix``."""
    for shape in shapes(page):
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


def _shape_texts(page):
    out = {}
    for i, shape in enumerate(shapes(page)):
        try:
            out[i] = _norm(shape.getString())
        except Exception:
            out[i] = ""
    return out


def _inventory(texts):
    """The slide's own text, indexed -- so a failed match self-corrects."""
    rows = [f"    [{i}] {t[:60]!r}" for i, t in sorted(texts.items()) if t]
    empty = [str(i) for i, t in sorted(texts.items()) if not t]
    out = ["  shapes carrying text:"] + (rows or ["    (none)"])
    if empty:
        out.append(f"  shapes without text: {', '.join(empty)}")
    return "\n".join(out)


def find_shapes(page, pattern, regex=False):
    """Every ``(index, shape)`` on ``page`` whose text matches ``pattern``.

    Substring match on whitespace-collapsed text; ``regex=True`` treats
    ``pattern`` as a case-insensitive regular expression.  Use it to hunt
    leftover template wording before shipping a deck.
    """
    out = []
    for i, shape in enumerate(shapes(page)):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        if not text:
            continue
        hit = (re.search(pattern, text, re.I) if regex
               else _norm(pattern).lower() in text.lower())
        if hit:
            out.append((i, shape))
    return out


def retext(page, mapping):
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

    Returns ``[(index, text)]`` for every shape left untouched that still
    carries text: the leftover-template-text check, for free.
    """
    by_index = dict(enumerate(shapes(page)))
    texts = _shape_texts(page)

    claimed = {}
    for key in mapping:
        want = _norm(key)
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
                f"(indices {hits}). Use a longer, more specific excerpt.\n"
                + _inventory(texts)
            )
        index = hits[0]
        if index in claimed:
            raise LookupError(
                f"retext: {key!r} and {claimed[index]!r} both match shape "
                f"[{index}] on {page.Name!r}; give each its own excerpt.\n"
                + _inventory(texts)
            )
        claimed[index] = key

    for index, key in claimed.items():
        value = mapping[key]
        if value is None:
            page.remove(by_index[index])
        elif isinstance(value, str):
            set_text(by_index[index], value)
        else:
            set_text(by_index[index], *value)

    return [(i, t) for i, t in sorted(texts.items()) if i not in claimed and t]


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

    ``index`` defaults to the end of the deck.  Unlike ``duplicate_slide()``
    the copy can go anywhere, which is the whole point -- see the module
    docstring.

    Shapes are rebuilt property by property.  Measured pixel-identical to the
    source on a fixture carrying text, an image, a table and a group; for a
    slide with an OLE chart or embedded media, ``duplicate_slide()`` is the
    safer copy when slide order happens to allow it.
    """
    dst = add_slide(doc, index)
    copy_background(doc, src_page, dst)
    for shape in shapes(src_page):
        clone_shape(doc, shape, dst)
    try:
        notes = notes_text(src_page)
        if notes:
            set_notes(dst, notes)
    except Exception:
        pass
    return dst
