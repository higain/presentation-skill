"""Vector shapes on a slide: rectangles, ellipses, lines, arrows, text, tables.

``impress.py`` opens and retexts documents; this draws on them.  Without it the
only shape you can create is a text box, and a deck that needs a diagram gets a
bullet list instead -- which is the single largest quality difference between
runs of this skill.

Every function takes ``(doc, page, ...)`` and returns the UNO shape it made, so
you can keep styling it afterwards.  Geometry is **1/100 mm** throughout, like
the rest of the skill; ``mm()`` and ``pt()`` from ``office.impress`` convert.
Colours are plain ints: ``0x009EE0``.

    import sys; sys.path.insert(0, "<skill>/scripts")
    from office.impress import impress_document, slides, mm
    from office import draw as D

    with impress_document("deck.pptx") as doc:
        page = slides(doc)[0]
        D.rect(doc, page, mm(20), mm(40), mm(60), mm(30), fill=0x009EE0)
        D.text(doc, page, mm(20), mm(40), mm(60), mm(30), "Q3",
               size=24, color=D.WHITE, weight=D.BOLD, align="center")

**The one trap:** ``setString()`` on a shape re-inherits the drawing default
(Arial 18pt black) and discards character formatting set before it.  So text
goes in *first* and formatting is forced onto the runs afterwards.  ``text()``
does this for you; if you call ``setString()`` yourself, call ``restyle()``
after it, never before.
"""

from __future__ import annotations

import uno

from .impress import shapes
# Re-exported so a drawing script needs one import: these edit shapes that are
# already on the slide, which is the other half of the same job.
from .template import move, resize, set_text  # noqa: F401

# Enums and structs are resolved by name rather than imported: a
# ``from com.sun.star...`` import only works once a UNO bridge exists, and this
# module is imported before one does.
_E = uno.Enum


def Point(x, y):
    p = uno.createUnoStruct("com.sun.star.awt.Point")
    p.X, p.Y = int(x), int(y)
    return p


def Size(w, h):
    s = uno.createUnoStruct("com.sun.star.awt.Size")
    s.Width, s.Height = int(w), int(h)
    return s


ITALIC = _E("com.sun.star.awt.FontSlant", "ITALIC")
UPRIGHT = _E("com.sun.star.awt.FontSlant", "NONE")
FILL_SOLID = _E("com.sun.star.drawing.FillStyle", "SOLID")
FILL_NONE = _E("com.sun.star.drawing.FillStyle", "NONE")
LINE_SOLID = _E("com.sun.star.drawing.LineStyle", "SOLID")
LINE_DASH = _E("com.sun.star.drawing.LineStyle", "DASH")
LINE_NONE = _E("com.sun.star.drawing.LineStyle", "NONE")
VADJ = {k: _E("com.sun.star.drawing.TextVerticalAdjust", v)
        for k, v in (("top", "TOP"), ("center", "CENTER"),
                     ("middle", "CENTER"), ("bottom", "BOTTOM"))}
HADJ_BLOCK = _E("com.sun.star.drawing.TextHorizontalAdjust", "BLOCK")
PADJ = {k: _E("com.sun.star.style.ParagraphAdjust", v)
        for k, v in (("left", "LEFT"), ("center", "CENTER"),
                     ("right", "RIGHT"), ("justify", "BLOCK"))}

STRIKE_SINGLE, STRIKE_NONE = 1, 0

# Font weights, as CharWeight wants them.
THIN, NORM, BOLD = 50.0, 100.0, 150.0

# Neutrals only.  Brand colours belong to the template -- read them off the
# deck with probe.py, or take them from the layout catalog; do not invent them.
BLACK = 0x000000
WHITE = 0xFFFFFF
GREY_DK = 0x4A4A4A
GREY = 0x8A8A8A
GREY_BD = 0xC8C8C8    # borders and rules
GREY_MD = 0xDCDCDC
GREY_LT = 0xF2F2F2    # zebra rows, panel fills

DEFAULT_FONT = "Arial"


# ------------------------------------------------------------- primitives

def _new(doc, page, kind, x, y, w, h):
    shape = doc.createInstance("com.sun.star.drawing." + kind)
    page.add(shape)
    shape.setPosition(Point(x, y))
    shape.setSize(Size(w, h))
    return shape


def style_fill(shape, fill=None, transparence=0, line=None, line_width=0,
               dashed=False):
    """Fill and outline.  ``fill=None`` means no fill, ``line=None`` no outline."""
    if fill is None:
        shape.FillStyle = FILL_NONE
    else:
        shape.FillStyle = FILL_SOLID
        shape.FillColor = int(fill)
        shape.FillTransparence = int(transparence)
    if line is None:
        shape.LineStyle = LINE_NONE
    else:
        shape.LineStyle = LINE_DASH if dashed else LINE_SOLID
        shape.LineColor = int(line)
        shape.LineWidth = int(line_width or 35)
    return shape


def rect(doc, page, x, y, w, h, fill=None, line=None, transparence=0,
         line_width=0, radius=None, dashed=False):
    """A rectangle, optionally rounded (``radius`` in 1/100 mm)."""
    s = _new(doc, page, "RectangleShape", x, y, w, h)
    style_fill(s, fill, transparence, line, line_width, dashed)
    if radius:
        try:
            s.CornerRadius = int(radius)
        except Exception:
            pass                    # not every export format keeps it
    return s


def ellipse(doc, page, cx, cy, rx, ry=None, fill=None, line=None,
            transparence=0, line_width=0):
    """An ellipse given its **centre** and radii -- not its bounding box."""
    ry = rx if ry is None else ry
    s = _new(doc, page, "EllipseShape", cx - rx, cy - ry, 2 * rx, 2 * ry)
    return style_fill(s, fill, transparence, line, line_width)


def line(doc, page, x1, y1, x2, y2, color=GREY_BD, width=35, dashed=False):
    """A straight line from (x1, y1) to (x2, y2)."""
    s = _new(doc, page, "LineShape", min(x1, x2), min(y1, y2),
             abs(x2 - x1), abs(y2 - y1))
    s.LineStyle = LINE_DASH if dashed else LINE_SOLID
    s.LineColor = int(color)
    s.LineWidth = int(width)
    return s


def hline(doc, page, x1, x2, y, color=GREY_BD, width=35, dashed=False):
    """A horizontal rule -- the commonest line, and the one worth a shortcut."""
    return line(doc, page, x1, y, x2, y, color, width, dashed)


def vline(doc, page, x, y1, y2, color=GREY_BD, width=35, dashed=False):
    return line(doc, page, x, y1, x, y2, color, width, dashed)


def polygon(doc, page, points, fill=None, line=None, line_width=0):
    """A closed polygon through ``[(x, y), ...]`` in 1/100 mm."""
    s = doc.createInstance("com.sun.star.drawing.PolyPolygonShape")
    page.add(s)
    s.PolyPolygon = (tuple(Point(px, py) for px, py in points),)
    return style_fill(s, fill, 0, line, line_width)


def arrow(doc, page, x, y, w, h, color=GREY, direction="right"):
    """A solid triangular arrowhead filling the given box."""
    mx, my = x + w / 2, y + h / 2
    pts = {
        "right": [(x, y), (x + w, my), (x, y + h)],
        "left": [(x + w, y), (x, my), (x + w, y + h)],
        "down": [(x, y), (x + w, y), (mx, y + h)],
        "up": [(x, y + h), (x + w, y + h), (mx, y)],
    }[direction]
    return polygon(doc, page, pts + [pts[0]], fill=color)


def chevron(doc, page, x, y, w, h, fill=None, notch=None, line=None,
            line_width=0):
    """A right-pointing chevron block -- process steps, phase bars."""
    notch = int(h * 0.28) if notch is None else notch
    pts = [(x, y), (x + w - notch, y), (x + w, y + h / 2),
           (x + w - notch, y + h), (x, y + h), (x + notch, y + h / 2), (x, y)]
    return polygon(doc, page, pts, fill=fill, line=line, line_width=line_width)


# ------------------------------------------------------------------- text

def restyle(shape, **props):
    """Force character/paragraph properties onto every paragraph and run.

    Setting ``shape.CharHeight`` alone is not enough once a shape holds text:
    the runs carry their own hard formatting and win.  This walks them.
    """
    try:
        enum = shape.createEnumeration()
    except Exception:
        return shape
    while enum.hasMoreElements():
        para = enum.nextElement()
        if not hasattr(para, "createEnumeration"):
            continue
        for k, v in props.items():
            try:
                setattr(para, k, v)
            except Exception:
                pass
        runs = para.createEnumeration()
        while runs.hasMoreElements():
            run = runs.nextElement()
            for k, v in props.items():
                try:
                    setattr(run, k, v)
                except Exception:
                    pass
    return shape


def _char_props(size=14.0, color=BLACK, weight=NORM, align="left",
                italic=False, strike=False, font=DEFAULT_FONT, underline=None):
    props = {
        "CharFontName": font,
        "CharHeight": float(size),
        "CharColor": int(color),
        "CharWeight": float(weight),
        "CharPosture": ITALIC if italic else UPRIGHT,
        "CharStrikeout": STRIKE_SINGLE if strike else STRIKE_NONE,
        "ParaAdjust": PADJ[align],
    }
    if underline is not None:
        props["CharUnderline"] = 1 if underline else 0
    return props


def text(doc, page, x, y, w, h, *lines, **kw):
    """A styled text box.  One positional string per paragraph.

    Keywords: ``size`` (pt), ``color``, ``weight`` (``BOLD``/``NORM``),
    ``align`` (left/center/right/justify), ``valign`` (top/center/bottom),
    ``font``, ``italic``, ``strike``, ``underline``, ``pad`` (inner margin),
    ``wrap``, ``grow`` (autogrow height), and the ``fill``/``line``/
    ``transparence``/``line_width``/``radius`` of ``rect()`` for a boxed label.
    """
    fill = kw.pop("fill", None)
    line_c = kw.pop("line", None)
    transparence = kw.pop("transparence", 0)
    line_width = kw.pop("line_width", 0)
    pad = int(kw.pop("pad", 0))
    valign = kw.pop("valign", "center")
    wrap = kw.pop("wrap", True)
    grow = kw.pop("grow", False)

    s = _new(doc, page, "TextShape", x, y, w, h)
    style_fill(s, fill, transparence, line_c, line_width)

    # Text first: setString() on an empty shape re-inherits the drawing
    # default and throws away formatting applied before it.
    s.setString("\n".join(str(part) for part in lines))

    s.TextVerticalAdjust = VADJ[valign]
    s.TextHorizontalAdjust = HADJ_BLOCK
    s.TextAutoGrowHeight = bool(grow)
    s.TextAutoGrowWidth = False
    s.TextWordWrap = bool(wrap)
    for side in ("Left", "Right", "Upper", "Lower"):
        setattr(s, "Text%sDistance" % side, pad)

    props = _char_props(**kw)
    for k, v in props.items():
        try:
            setattr(s, k, v)
        except Exception:
            pass
    restyle(s, **props)
    return s


def chip(doc, page, x, y, w, h, fill, label="", color=WHITE, size=10.0,
         weight=BOLD, radius=None, line=None):
    """A small filled block with a centred label -- legend keys, badges, tags."""
    box = rect(doc, page, x, y, w, h, fill=fill, radius=radius, line=line)
    if label:
        text(doc, page, x, y, w, h, label, size=size, color=color,
             weight=weight, align="center", valign="center")
    return box


def bullets(doc, page, x, y, w, items, size=12.5, color=BLACK,
            marker_color=None, marker=None, gap=420, line_h=None,
            marker_size=170, font=DEFAULT_FONT):
    """A bulleted list, one text box per item, with a square marker.

    Drawn rather than typed so each item can be positioned and coloured.
    Returns the y just past the last item, so lists can be stacked.
    """
    line_h = line_h or int(size * 41)
    marker_color = BLACK if marker_color is None else marker_color
    cy = y
    for item in items:
        rows = est_lines(str(item), w - 800, size)
        bh = line_h * rows
        if marker:
            text(doc, page, x, cy, int(marker_size * 3), line_h, marker,
                 size=size, color=marker_color, weight=BOLD, valign="top",
                 font=font)
        else:
            rect(doc, page, x, cy + int(size * 15), marker_size, marker_size,
                 fill=marker_color)
        text(doc, page, x + 500, cy, w - 500, bh + 200, str(item),
             size=size, color=color, valign="top", font=font)
        cy += bh + gap
    return cy


def est_lines(s, width, size, font_factor=0.52):
    """Rough wrapped-line count.  Arial's mean glyph advance is ~0.52 em.

    An estimate, for laying out a list before the document can measure it.
    ``office.template.text_overflow`` measures for real -- use that to check.
    """
    per_char = size * 35.28 * font_factor      # pt -> 1/100 mm
    return max(1, int(len(s) * per_char / max(width, 1)) + 1)


# ------------------------------------------------------------------ table

def table(doc, page, x, y, widths, rows, row_h, header_h=None,
          head_fill=GREY_DK, head_color=WHITE, size=12.0, head_size=12.0,
          zebra=GREY_LT, body_fill=WHITE, align=None, pad=200,
          cell_colors=None, cell_weights=None, row_heights=None,
          rule=GREY_BD, font=DEFAULT_FONT):
    """Draw a table out of rectangles and text.  ``rows[0]`` is the header.

    A drawn table rather than a real ``TableShape``: it survives a .pptx
    round-trip unchanged, takes per-cell colour without fighting a table style,
    and needs no column-model bookkeeping.  Use a real table only when the
    recipient has to edit it as one.

    ``cell_colors``/``cell_weights`` are ``{(row, col): value}`` with row 1 the
    first body row.  Returns the y just past the last row.
    """
    align = align or ["left"] * len(widths)
    header_h = header_h or row_h
    cy = y
    for r, row in enumerate(rows):
        h = header_h if r == 0 else (row_heights[r - 1] if row_heights else row_h)
        cx = x
        for c, cell in enumerate(row):
            w = widths[c]
            head = r == 0
            rect(doc, page, cx, cy, w, h,
                 fill=head_fill if head else
                 (zebra if (zebra and r % 2 == 1) else body_fill))
            text(doc, page, cx, cy, w, h, str(cell),
                 size=head_size if head else size,
                 color=head_color if head else
                 (cell_colors or {}).get((r, c), BLACK),
                 weight=BOLD if head else
                 (cell_weights or {}).get((r, c), NORM),
                 align=align[c], valign="center", pad=pad, font=font)
            cx += w
        cy += h
    if rule is not None:
        hline(doc, page, x, x + sum(widths), cy, rule, 30)
    return cy


# ------------------------------------------------------------------- venn

def venn_geometry(cx, cy, r, n=3):
    """Circle centres for a symmetric Venn of radius ``r`` about ``(cx, cy)``.

    For ``n=3`` the side of the centre triangle is ``r``, which puts each
    centre on a vertex of the three-way overlap and leaves that overlap a
    Reuleaux triangle of inradius 0.42 r -- enough room for a label.
    """
    if n == 2:
        return [(cx - r / 2, cy), (cx + r / 2, cy)]
    d = r / 3 ** 0.5
    h = r * (3 ** 0.5) / 2
    return [(cx, cy - d), (cx - r / 2, cy - d + h), (cx + r / 2, cy - d + h)]


def venn(doc, page, cx, cy, r, colors, labels=None, transparence=76,
         line_width=45, label_size=None, label_color=None, centre=None,
         centre_color=None, centre_text=None, centre_size=None):
    """Two or three overlapping circles.  Returns the centres, so you can
    label the regions yourself.

    ``colors`` is one colour per circle and decides how many are drawn --
    pass a prefix of it to build the diagram up over successive slides.
    ``labels``, if given, are placed just outside each circle.
    ``centre`` fills the all-way overlap with that colour.
    """
    n = len(colors)
    centres = venn_geometry(cx, cy, r, n)
    for i, color in enumerate(colors):
        ex, ey = centres[i]
        ellipse(doc, page, ex, ey, r, fill=color, transparence=transparence,
                line=color, line_width=line_width)
    if labels:
        box_w = r * 2.2
        for i, label in enumerate(labels[:n]):
            # Push the label out along the ray from the figure's centre through
            # this circle's centre, so it clears the circle instead of landing
            # inside it -- and so two- and three-circle figures need no cases.
            ex, ey = centres[i]
            dx, dy = ex - cx, ey - cy
            norm = (dx * dx + dy * dy) ** 0.5 or 1
            sx = ex + dx / norm * r * 1.25
            sy = ey + dy / norm * r * 1.25
            # A one-word label wants to be big; "EXTERNAL COMMUNICATION" wants
            # to fit in two lines rather than break mid-word off the slide.
            size = label_size or max(
                7.0, min(r / 145.0, box_w * 2 / (max(len(str(label)), 1) * 19.5))
            )
            text(doc, page, sx - box_w / 2, sy - r * 0.36, box_w, r * 0.72,
                 label, size=size, weight=BOLD, align="center",
                 valign="center", color=(label_color or colors[i]))
    if centre is not None:
        cr = r * 0.41
        ellipse(doc, page, cx, cy, cr, fill=centre, line=centre,
                line_width=line_width)
        if centre_text:
            if isinstance(centre_text, str):
                centre_text = [centre_text]
            box_w = cr * 1.90
            # Size the centre label to the centre circle, not to the figure:
            # the widest line has to fit across it on one line.  A fraction of
            # the outer radius reads far too small once the circles are modest.
            widest = max(len(str(part)) for part in centre_text) or 1
            size = centre_size or max(7.0, min(r / 145.0, box_w / (widest * 18.3)))
            text(doc, page, cx - box_w / 2, cy - cr * 0.60, box_w, cr * 1.2,
                 *centre_text, size=size,
                 color=(centre_color or WHITE), weight=BOLD, align="center",
                 valign="center")
    return centres


def pair_spots(cx, cy, r):
    """Centres of the three two-of-three-only regions of a 3-circle Venn,
    in the order (0+1), (0+2), (1+2) -- where a "any two is fine" label goes."""
    c = venn_geometry(cx, cy, r, 3)
    out = []
    for a, b, other in ((0, 1, 2), (0, 2, 1), (1, 2, 0)):
        mx = (c[a][0] + c[b][0]) / 2
        my = (c[a][1] + c[b][1]) / 2
        dx, dy = mx - c[other][0], my - c[other][1]
        norm = (dx * dx + dy * dy) ** 0.5 or 1
        out.append((mx + dx / norm * r * 0.28, my + dy / norm * r * 0.28))
    return out


# ------------------------------------------------- editing existing shapes

def walk(shape):
    """A shape and, recursively, every shape inside it.

    Template slides put half their text inside groups, where ``shapes(page)``
    cannot see it.  ``office.template.retext`` and ``qa.py`` walk groups for
    you; use this when you need the shapes themselves.
    """
    yield shape
    try:
        count = shape.getCount()
    except Exception:
        return
    for i in range(count):
        yield from walk(shape.getByIndex(i))


def walk_page(page):
    """Every shape on a slide, groups included, depth first."""
    for top in shapes(page):
        yield from walk(top)


def canvas(page, keep=(), remove=()):
    """Strip a template slide back to a blank drawing area.

    Pass ``keep`` as shape indices to leave in place (title, topic label) and
    ``remove`` as the ones to drop; giving only ``keep`` removes everything
    else.  The point is to reuse a designed slide's background, master and
    chrome while drawing your own content on it -- which is how you get a
    diagram onto a template that has no design for one.

    Returns the shapes left on the page.
    """
    current = shapes(page)
    if keep:
        doomed = [s for i, s in enumerate(current) if i not in set(keep)]
    else:
        doomed = [current[i] for i in remove if 0 <= i < len(current)]
    for shape in doomed:
        page.remove(shape)
    return shapes(page)


__all__ = [
    "Point", "Size", "BLACK", "WHITE", "GREY", "GREY_DK", "GREY_BD",
    "GREY_MD", "GREY_LT", "BOLD", "NORM", "THIN", "DEFAULT_FONT",
    "rect", "ellipse", "line", "hline", "vline", "polygon", "arrow",
    "chevron", "text", "chip", "bullets", "table", "venn", "venn_geometry",
    "pair_spots", "restyle", "style_fill", "walk", "walk_page", "move",
    "resize", "canvas", "est_lines", "set_text",
]
