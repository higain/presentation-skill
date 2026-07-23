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
