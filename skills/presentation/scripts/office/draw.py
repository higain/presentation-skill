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

import functools
import sys

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
                     ("centre", "CENTER"), ("right", "RIGHT"),
                     ("justify", "BLOCK"))}


def _adjust(table, value, default, kind):
    """Look ``value`` up in ``PADJ``/``VADJ``, forgiving ``None`` and case.

    ``align=None`` reads as "leave it alone" at every call site that writes it
    -- ``align="center" if numeric else None`` -- and a bare ``KeyError: None``
    from four frames down says nothing about which argument was wrong.
    """
    if value is None:
        value = default
    try:
        return table[str(value).lower()]
    except KeyError:
        raise ValueError(
            "%s=%r is not one of %s" % (kind, value, "/".join(sorted(table)))
        ) from None

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

# Ceilings for labels a diagram sizes for itself.  These are typographic, not
# geometric: a label scaled to the figure's radius reads as a second title on a
# large diagram and collides with whatever the caller drew beneath it.  Body
# text on a brand template is 12-14pt and a slide title 16-20pt, so a label
# that fits sits between them, and shrinks from there when the space is tight.
LABEL_MAX_PT = 18.0
CENTRE_MAX_PT = 16.0


# ------------------------------------------------------------------- layout

_SPLIT_HELP = (
    "cut_top/cut_bottom/cut_left/cut_right return a PAIR of boxes, not a box.\n"
    "  unpack it:      head, body = area.cut_top(mm(14))\n"
    "  or take one:    area.top(mm(14))    area.bottom(mm(14))\n"
    "                  area.left(mm(60))   area.right(mm(60))\n"
    "Those return a single Box and cannot come out in the wrong order."
)


class Split(tuple):
    """The two boxes a ``cut_*`` returns, in reading order.

    Top-to-bottom for ``cut_top``/``cut_bottom``, left-to-right for
    ``cut_left``/``cut_right`` -- so the pair always unpacks the way the slide
    reads, and *which half is the band changes with the direction*::

        head, body = area.cut_top(mm(14))      # head is the band
        body, foot = area.cut_bottom(mm(14))   # foot is the band

    That is the trap this class exists to make loud.  Getting it backwards is
    silent: you fill the whole panel instead of a 14 mm strip of it, and no
    check catches it because nothing is off the slide.  If you only want one of
    the two, name it -- ``area.top(mm(14))``, ``area.bottom(mm(14))`` -- or use
    ``.band`` / ``.rest``, which mean the same thing in every direction.
    """

    # No __slots__: a tuple subclass cannot have a nonempty one, and the
    # band index has to live somewhere.

    def __new__(cls, first, second, band_at):
        self = super().__new__(cls, (first, second))
        self._band_at = band_at
        return self

    @property
    def band(self):
        """The strip that was cut off -- whichever end it came from."""
        return self[self._band_at]

    @property
    def rest(self):
        """What is left after the band."""
        return self[1 - self._band_at]

    def __repr__(self):
        return "Split(band=%r, rest=%r)" % (self.band, self.rest)

    def __getattr__(self, name):
        raise AttributeError(
            "a pair of boxes has no %r.\n%s" % (name, _SPLIT_HELP)
        )


class Box:
    """A rectangle on the slide -- ``x, y, w, h`` in 1/100 mm.

    **Pass it whole.** Every drawing call takes a ``Box`` where its geometry
    goes, so there is nothing to splat and nothing to get out of order::

        D.rect(doc, page, box, fill=0xEEEEEE)

    It exists because the alternative is arithmetic on hardcoded millimetres,
    and that is where built decks go wrong: a run of ``mm(30 + i * 70)``
    silently assumes a slide width, drifts out of alignment when a column is
    added, and leaves a third of the slide empty because nobody subtracted.
    Ask the area for the pieces you want instead::

        area = D.canvas(page, keep=(0, 1, 2, 3))    # the free area, as a Box
        head, body = area.cut_top(mm(14), gap=mm(6))
        for col, step in zip(body.cols(4, gap=mm(6)), steps):
            D.rect(doc, page, col, fill=0xEEEEEE, radius=300)
            D.text(doc, page, col, step, size=13, align="center")

    Slicing a box always consumes all of it, so the content fills the area the
    template left -- which is the difference between a slide that looks placed
    and one that floats in the top half.

    **The one thing to get right:** ``cut_top`` and friends return *two* boxes
    (a :class:`Split`), so they must be unpacked -- ``head, body = ...``.  When
    you want only one of the two, ask for it by name instead::

        area.top(mm(14))     area.bottom(mm(14))
        area.left(mm(60))    area.right(mm(60))

    Those return a single ``Box``, go straight into a drawing call, and have no
    order to get backwards.  Reach for them first.
    """

    __slots__ = ("x", "y", "w", "h")

    def __init__(self, x, y, w, h):
        self.x, self.y, self.w, self.h = int(x), int(y), int(w), int(h)

    def __iter__(self):
        return iter((self.x, self.y, self.w, self.h))

    def __repr__(self):
        return "Box(x=%d, y=%d, w=%d, h=%d)  # %.0f,%.0fmm %.0fx%.0fmm" % (
            self.x, self.y, self.w, self.h,
            self.x / 100, self.y / 100, self.w / 100, self.h / 100)

    @property
    def x2(self):
        return self.x + self.w

    @property
    def y2(self):
        return self.y + self.h

    @property
    def cx(self):
        return self.x + self.w // 2

    @property
    def cy(self):
        return self.y + self.h // 2

    @property
    def centre(self):
        """``(cx, cy)`` -- unpacks into ``ellipse`` and ``venn``."""
        return (self.cx, self.cy)

    def inset(self, all=0, x=None, y=None, left=None, top=None,
              right=None, bottom=None):
        """A smaller box inside this one.  ``all`` unless overridden per side."""
        l = left if left is not None else (x if x is not None else all)
        r = right if right is not None else (x if x is not None else all)
        t = top if top is not None else (y if y is not None else all)
        b = bottom if bottom is not None else (y if y is not None else all)
        return Box(self.x + l, self.y + t,
                   max(0, self.w - l - r), max(0, self.h - t - b))

    def cols(self, n, gap=0):
        """Split into ``n`` equal columns separated by ``gap``."""
        if n <= 0:
            return []
        each = (self.w - gap * (n - 1)) / n
        return [Box(self.x + i * (each + gap), self.y, each, self.h)
                for i in range(n)]

    def rows(self, n, gap=0):
        """Split into ``n`` equal rows separated by ``gap``."""
        if n <= 0:
            return []
        each = (self.h - gap * (n - 1)) / n
        return [Box(self.x, self.y + i * (each + gap), self.w, each)
                for i in range(n)]

    def grid(self, rows, cols, gap=0, row_gap=None):
        """``[[Box, ...], ...]`` -- ``rows`` lists of ``cols`` boxes."""
        return [r.cols(cols, gap)
                for r in self.rows(rows, gap if row_gap is None else row_gap)]

    # -- single slices.  Unambiguous: there is no order to get wrong. --

    def top(self, h):
        """The band ``h`` tall off the top, as one ``Box``."""
        return Box(self.x, self.y, self.w, min(h, self.h))

    def bottom(self, h):
        """The band ``h`` tall off the bottom, as one ``Box``."""
        h = min(h, self.h)
        return Box(self.x, self.y2 - h, self.w, h)

    def left(self, w):
        """The band ``w`` wide off the left, as one ``Box``."""
        return Box(self.x, self.y, min(w, self.w), self.h)

    def right(self, w):
        """The band ``w`` wide off the right, as one ``Box``."""
        w = min(w, self.w)
        return Box(self.x2 - w, self.y, w, self.h)

    # -- paired slices.  Always in reading order; see ``Split``. --

    def cut_top(self, h, gap=0):
        """``Split(band, rest)`` -- a band ``h`` tall off the top, then the rest.

        The band comes **first**.  For just one of the two use ``top(h)`` or
        ``bottom(h)``, which cannot be got the wrong way round.
        """
        return Split(Box(self.x, self.y, self.w, h),
                     Box(self.x, self.y + h + gap, self.w,
                         max(0, self.h - h - gap)),
                     band_at=0)

    def cut_bottom(self, h, gap=0):
        """``Split(rest, band)`` -- top-to-bottom order, so the band is **second**."""
        return Split(Box(self.x, self.y, self.w, max(0, self.h - h - gap)),
                     Box(self.x, self.y2 - h, self.w, h),
                     band_at=1)

    def cut_left(self, w, gap=0):
        """``Split(band, rest)`` -- the band is **first**."""
        return Split(Box(self.x, self.y, w, self.h),
                     Box(self.x + w + gap, self.y,
                         max(0, self.w - w - gap), self.h),
                     band_at=0)

    def cut_right(self, w, gap=0):
        """``Split(rest, band)`` -- left-to-right, so the band is **second**."""
        return Split(Box(self.x, self.y, max(0, self.w - w - gap), self.h),
                     Box(self.x2 - w, self.y, w, self.h),
                     band_at=1)

    def square(self, scale=1.0):
        """The largest centred square that fits, times ``scale``.

        For anything drawn from a radius -- a Venn, a pie, a ring.
        ``r = box.square().w / 2`` never runs off the area.
        """
        side = min(self.w, self.h) * scale
        return Box(self.cx - side / 2, self.cy - side / 2, side, side)


def page_box(page):
    """The whole slide as a ``Box``.  ``page_box(page).inset(mm(15))`` is a
    reasonable margin on any slide size -- read it, never hardcode 28000."""
    return Box(0, 0, page.Width, page.Height)


# ------------------------------------------------------------- primitives

def _num(value, arg, fn):
    """Coordinates are numbers in 1/100 mm.

    A ``Box`` reaching here is a box in a slot that is not the geometry slot --
    every drawing call takes one whole ``Box`` in place of its leading
    coordinates (see :func:`takes_box`), so a second one has nowhere to go.
    A ``Split`` or a list of boxes is the other half of the same mistake.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        if isinstance(value, Split):
            raise TypeError(
                "%s(): %s= got a pair of boxes.  You splatted a cut_* result, "
                "and `*box.cut_top(h)` expands to TWO boxes, not to x, y, w, h."
                "\n%s" % (fn, arg, _SPLIT_HELP))
        if isinstance(value, Box):
            raise TypeError(
                "%s(): %s= got a Box.  Pass the box on its own, in place of "
                "all the coordinates -- `%s(doc, page, box, ...)` -- not one "
                "box per argument." % (fn, arg, fn))
        if (isinstance(value, (list, tuple)) and value
                and isinstance(value[0], Box)):
            raise TypeError(
                "%s(): %s= got a list of %d Box(es).  `.cols()`/`.rows()` "
                "return a list and `.grid()` returns a list of rows -- iterate "
                "it, one box per shape:\n"
                "    for cell in area.cols(3, gap=mm(6)):\n"
                "        D.rect(doc, page, cell, fill=0xEEEEEE)\n"
                "    for row in area.grid(2, 3, gap=mm(6)):\n"
                "        for cell in row:\n"
                "            ..." % (fn, arg, len(value)))
        raise TypeError(
            "%s(): %s= must be a number in 1/100 mm (use mm(20)), got %s"
            % (fn, arg, type(value).__name__))
    return value


def _xywh(fn, x, y, w, h):
    return (_num(x, "x", fn), _num(y, "y", fn),
            _num(w, "w", fn), _num(h, "h", fn))


# A ``Box`` is what the layout helpers return, and it is what a caller has in
# hand at every drawing call.  Requiring `*box` to splat it bought nothing and
# cost whole build cycles: forgetting the star raises deep inside the callee,
# where the message names an argument the caller never wrote.  So every
# drawing call takes the box itself, and the splat still works.
_BOX_PARTS = {
    "x": lambda b: b.x, "y": lambda b: b.y,
    "w": lambda b: b.w, "h": lambda b: b.h,
    "x2": lambda b: b.x2, "y2": lambda b: b.y2,
    "cx": lambda b: b.cx, "cy": lambda b: b.cy,
    "rx": lambda b: b.w // 2, "ry": lambda b: b.h // 2,
}


def takes_box(*names, at=2):
    """Accept one ``Box`` in place of the coordinates named by ``names``.

    ``D.rect(doc, page, box, fill=...)`` and
    ``D.rect(doc, page, *box, fill=...)`` both work, and so does the long form
    with four numbers.  ``names`` says what the box means to this call:
    ``"x", "y", "w", "h"`` for a rectangle, ``"cx", "cy", "rx", "ry"`` for an
    ellipse, which is why passing a box to ``ellipse`` gives its bounding box
    rather than requiring centre-and-radius arithmetic.
    """
    def decorate(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kw):
            if len(args) > at and not isinstance(args[at], (int, float, Box)):
                # A list of boxes here is arity-checked by Python before any
                # coordinate is looked at, so the useful message has to be
                # raised now or it never gets raised at all.
                _num(args[at], names[0], fn.__name__)
            if len(args) > at and isinstance(args[at], Box):
                if len(args) > at + 1 and isinstance(args[at + 1], Box):
                    raise TypeError(
                        "%s(): got two boxes where one goes.  You splatted a "
                        "cut_* result, and `*box.cut_top(h)` expands to TWO "
                        "boxes.\n%s" % (fn.__name__, _SPLIT_HELP))
                box = args[at]
                args = (args[:at] + tuple(_BOX_PARTS[n](box) for n in names)
                        + args[at + 1:])
            return fn(*args, **kw)
        return wrapper
    return decorate


def _new(doc, page, kind, x, y, w, h):
    x, y, w, h = _xywh(kind[:-5].lower(), x, y, w, h)
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


@takes_box("x", "y", "w", "h")
def rect(doc, page, x, y, w, h, fill=None, line=None, transparence=0,
         line_width=0, radius=None, dashed=False):
    """A rectangle, optionally rounded (``radius`` in 1/100 mm).

    Takes a ``Box`` in place of ``x, y, w, h``::

        D.rect(doc, page, area.top(mm(14)), fill=0x009EE0)
    """
    s = _new(doc, page, "RectangleShape", x, y, w, h)
    style_fill(s, fill, transparence, line, line_width, dashed)
    if radius:
        try:
            s.CornerRadius = int(radius)
        except Exception:
            pass                    # not every export format keeps it
    return s


_venn_hinted = set()

_VENN_HELP = """\
[draw] three ellipses on one slide -- if this is a Venn, use D.venn() instead:

    D.venn(doc, page, *area.centre, area.w * 0.15,
           [0x009EE0, 0xE8730C, 0x6B3FA0],
           labels=["PRIVATE DATA", "UNTRUSTED CONTENT", "EXTERNAL COMMS"],
           centre=0xC00000, centre_text="DATA EXFILTRATION")

venn() places the circles so they actually overlap, sizes the figure to the
area, and puts each label outside its own lobe.  Hand-placed circles come out
tangent or overlapping their labels, and no automated check catches either --
that is a rendering defect, and it needs eyes.  D.venn_fit(box) if you want the
geometry without the drawing."""


_drawing_venn = False


def _venn_hint(page):
    """Nudge once per slide when a third ellipse lands on it.  See ``venn``."""
    if _drawing_venn:
        return
    try:
        key = page.Name
        count = sum(1 for s in shapes(page)
                    if s.getShapeType() == "com.sun.star.drawing.EllipseShape")
    except Exception:
        return
    if count == 3 and key not in _venn_hinted:
        _venn_hinted.add(key)
        print(_VENN_HELP, file=sys.stderr)


@takes_box("cx", "cy", "rx", "ry")
def ellipse(doc, page, cx, cy, rx, ry=None, fill=None, line=None,
            transparence=0, line_width=0):
    """An ellipse given its **centre** and radii -- not its bounding box.

    Pass a ``Box`` instead and it *is* read as the bounding box, which is
    usually what you have::

        D.ellipse(doc, page, area.square(0.6), fill=0x009EE0)

    For a Venn diagram use :func:`venn`, not three of these: it computes the
    overlap and the label placement, which is the part that gets drawn wrong.
    """
    ry = rx if ry is None else ry
    cx, cy = _num(cx, "cx", "ellipse"), _num(cy, "cy", "ellipse")
    rx, ry = _num(rx, "rx", "ellipse"), _num(ry, "ry", "ellipse")
    s = _new(doc, page, "EllipseShape", cx - rx, cy - ry, 2 * rx, 2 * ry)
    style_fill(s, fill, transparence, line, line_width)
    _venn_hint(page)
    return s


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


@takes_box("x", "y", "w", "h")
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


@takes_box("x", "y", "w", "h")
def chevron(doc, page, x, y, w, h, fill=None, notch=None, line=None,
            line_width=0):
    """A right-pointing chevron block -- process steps, phase bars."""
    notch = int(h * 0.28) if notch is None else notch
    pts = [(x, y), (x + w - notch, y), (x + w, y + h / 2),
           (x + w - notch, y + h), (x, y + h), (x + notch, y + h / 2), (x, y)]
    return polygon(doc, page, pts, fill=fill, line=line, line_width=line_width)


# -------------------------------------------------------------- contrast

def _rel_luminance(color):
    """WCAG relative luminance of a 0xRRGGBB int."""
    def channel(v):
        v /= 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    color = int(color)
    return (0.2126 * channel((color >> 16) & 255)
            + 0.7152 * channel((color >> 8) & 255)
            + 0.0722 * channel(color & 255))


def contrast_ratio(a, b):
    """WCAG contrast ratio between two colours, 1.0 (same) to 21.0 (black/white).

    Below 3.0 is unreadable at any size a slide uses.  ``qa.py`` flags it, but
    a build script can check before it draws.
    """
    la, lb = _rel_luminance(a), _rel_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def bg_color(page):
    """The slide's own background colour as an int, or ``None`` if it has no
    flat one (a gradient, a bitmap, or nothing of its own and none inherited).

    A dark section divider gets its colour from here, not from a shape, so a
    script that clears the slide and draws text on it has no other way to know
    it is drawing onto navy.  That is exactly how a deck ships an invisible
    headline: ``text()`` used to default to black regardless.
    """
    for holder in (page, getattr(page, "MasterPage", None)):
        if holder is None:
            continue
        try:
            background = holder.Background
            if background is None:
                continue
            if background.FillStyle == FILL_SOLID:
                return int(background.FillColor)
        except Exception:
            continue
    return None


def readable_on(background, dark=BLACK, light=WHITE):
    """``dark`` or ``light``, whichever reads on ``background``."""
    if background is None:
        return dark
    return light if contrast_ratio(background, light) >= \
        contrast_ratio(background, dark) else dark


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
        "ParaAdjust": _adjust(PADJ, align, "left", "align"),
    }
    if underline is not None:
        props["CharUnderline"] = 1 if underline else 0
    return props


def _text_height(lines, w, size, weight=NORM, pad=0, line_spacing=1.22):
    """Height a text box needs for ``lines`` wrapped to ``w``, in 1/100 mm.

    Deliberately generous by a quarter line: a box measured exactly to its
    estimate is the one ``qa.py`` reports as overflowing, because LibreOffice
    lays the text out slightly differently than this arithmetic predicts.
    """
    count = 0
    for part in lines:
        count += len(wrap_words(str(part), max(w - 2 * pad, 1), size, weight))
    count = max(count, 1)
    return int((count + 0.25) * size * 35.28 * line_spacing) + 2 * pad


@takes_box("x", "y", "w", "h")
def text(doc, page, x, y, w, h=None, *lines, **kw):
    """A styled text box.  One positional string per paragraph.

    ``h`` is optional: omit it (or pass ``None``) and the box is sized to the
    text it holds at ``size``, wrapped to ``w``.  So all of these work::

        D.text(doc, page, box, "Heading", size=20)       # box gives w and h
        D.text(doc, page, *box, "Heading", size=20)
        D.text(doc, page, box.x, box.y, box.w, "Heading", size=20)

    ``color`` defaults to whatever reads on the slide's background -- white on
    a dark divider, black on a white body slide.  Pass it explicitly to
    override.  Nothing else guesses: a headline drawn in the default black onto
    navy is invisible, renders clean through every automated check, and is the
    single defect most likely to survive to the deck you present.

    Keywords: ``size`` (pt), ``color``, ``weight`` (``BOLD``/``NORM``),
    ``align`` (left/center/right/justify), ``valign`` (top/center/bottom),
    ``font``, ``italic``, ``strike``, ``underline``, ``pad`` (inner margin),
    ``wrap``, ``grow`` (autogrow height), and the ``fill``/``line``/
    ``transparence``/``line_width``/``radius`` of ``rect()`` for a boxed label.
    """
    # `D.text(doc, page, x, y, w, "the text")` -- h omitted, first line in its
    # place.  Without this the signature error names `h`, which sends the
    # caller looking for a missing number instead of a missing height.
    if isinstance(h, str):
        lines = (h,) + lines
        h = None

    fill = kw.pop("fill", None)
    line_c = kw.pop("line", None)
    transparence = kw.pop("transparence", 0)
    line_width = kw.pop("line_width", 0)
    pad = int(kw.pop("pad", 0))
    valign = kw.pop("valign", "center")
    wrap = kw.pop("wrap", True)
    grow = kw.pop("grow", False)

    if kw.get("color") is None:
        kw["color"] = readable_on(fill if fill is not None else bg_color(page))

    if h is None:
        h = _text_height(lines, _num(w, "w", "text"),
                         float(kw.get("size", 14.0)),
                         float(kw.get("weight", NORM)), pad)

    s = _new(doc, page, "TextShape", x, y, w, h)
    style_fill(s, fill, transparence, line_c, line_width)

    # Text first: setString() on an empty shape re-inherits the drawing
    # default and throws away formatting applied before it.
    s.setString("\n".join(str(part) for part in lines))

    s.TextVerticalAdjust = _adjust(VADJ, valign, "center", "valign")
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


@takes_box("x", "y", "w", "h")
def chip(doc, page, x, y, w, h, fill, label="", color=WHITE, size=10.0,
         weight=BOLD, radius=None, line=None):
    """A small filled block with a centred label -- legend keys, badges, tags."""
    box = rect(doc, page, x, y, w, h, fill=fill, radius=radius, line=line)
    if label:
        text(doc, page, x, y, w, h, label, size=size, color=color,
             weight=weight, align="center", valign="center")
    return box


@takes_box("x", "y", "w")
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


# Mean glyph advance in em, by character class.  A single 0.52 average is right
# for a sentence of mixed-case prose and badly wrong for the short bold
# uppercase labels that diagrams are made of: "FORBIDDEN" set at a size derived
# from 0.52 overruns its box by a third and breaks mid-word.  Caps run ~0.72 em,
# and bold adds ~6% on top.
_NARROW = set("ijltfrI.,;:'\"!|()[]{}/\\-")
_WIDE = set("mwMW@%")


def advance_em(s, weight=NORM):
    """Mean glyph advance of ``s``, in em, for Arial at ``weight``.

    Used to size a label to its box.  Accurate to a few percent on the short
    strings diagrams carry, which is what matters -- for a paragraph the error
    averages out and any of these constants would do.
    """
    if not s:
        return 0.52
    total = 0.0
    for ch in s:
        if ch == " ":
            total += 0.28
        elif ch in _NARROW:
            total += 0.26
        elif ch in _WIDE:
            total += 0.85
        elif ch.isupper():
            total += 0.72
        elif ch.isdigit():
            total += 0.56
        else:
            total += 0.52
    return total / len(s)


def text_width_est(s, size, weight=NORM):
    """Estimated rendered width of ``s`` at ``size`` pt, in 1/100 mm."""
    return len(s) * size * 35.28 * advance_em(s, weight)


def fit_size(s, width, weight=NORM, max_size=44.0, min_size=7.0, slack=0.94):
    """Largest point size at which ``s`` fits ``width`` on one line.

    ``slack`` keeps the result just inside the box rather than exactly on it --
    without it every auto-sized label lands within a percent of its own width
    and half of them wrap, which is the mid-word break ``qa.py`` reports.
    """
    if not s:
        return max_size
    per_pt = len(s) * 35.28 * advance_em(s, weight)
    return max(min_size, min(max_size, width * slack / max(per_pt, 1e-6)))


def wrap_words(s, width, size, weight=NORM):
    """Greedy word wrap of ``s`` to ``width`` at ``size`` pt.  ``[line, ...]``."""
    out, cur = [], ""
    for word in str(s).split():
        trial = (cur + " " + word).strip()
        if cur and text_width_est(trial, size, weight) > width:
            out.append(cur)
            cur = word
        else:
            cur = trial
    if cur:
        out.append(cur)
    return out or [""]


def fit_lines(s, width, height, weight=NORM, max_size=44.0, min_size=7.0,
              line_spacing=1.22):
    """Size ``s`` to fill ``width`` x ``height``, wrapping on spaces.

    Returns ``(size, lines)``.  Unlike ``fit_size`` this may use several lines,
    so a two-word label gets a readable size instead of being shrunk until it
    fits on one -- and it never breaks a word, because it only ever wraps at a
    space.  A single word that cannot fit at ``min_size`` is returned anyway;
    widen the box.
    """
    longest = max(str(s).split() or [""], key=len)
    ceiling = min(max_size, fit_size(longest, width, weight, max_size, min_size))
    size = ceiling
    while size > min_size:
        lines = wrap_words(s, width, size, weight)
        if len(lines) * size * 35.28 * line_spacing <= height:
            return size, lines
        size -= 0.5
    return max(size, min_size), wrap_words(s, width, max(size, min_size), weight)


def est_lines(s, width, size, font_factor=None):
    """Rough wrapped-line count, for laying out a list before the document can
    measure it.  ``office.template.text_overflow`` measures for real -- use that
    to check.  ``font_factor`` overrides the per-character estimate.
    """
    factor = advance_em(s) if font_factor is None else font_factor
    per_char = size * 35.28 * factor           # pt -> 1/100 mm
    return max(1, int(len(s) * per_char / max(width, 1)) + 1)


# ------------------------------------------------------------------ table

def table(doc, page, x=None, y=None, widths=(), rows=(), row_h=None,
          header_h=None,
          head_fill=GREY_DK, head_color=WHITE, size=12.0, head_size=12.0,
          zebra=GREY_LT, body_fill=WHITE, align=None, pad=200,
          cell_colors=None, cell_weights=None, row_heights=None,
          rule=GREY_BD, font=DEFAULT_FONT, box=None, **kw):
    """Draw a table out of rectangles and text.  ``rows[0]`` is the header.

    A drawn table rather than a real ``TableShape``: it survives a .pptx
    round-trip unchanged, takes per-cell colour without fighting a table style,
    and needs no column-model bookkeeping.  Use a real table only when the
    recipient has to edit it as one.

    ``cell_colors``/``cell_weights`` are ``{(row, col): value}`` with row 1 the
    first body row.  Returns the y just past the last row.

    **Pass ``box`` and give ``widths`` as proportions** -- ``[3, 2, 2, 2, 2]``,
    or fractions summing to 1 -- and the columns are scaled to fit it exactly::

        D.table(doc, page, widths=[3, 2, 2, 2, 2], rows=rows,
                row_h=mm(14), box=area, head_fill=0x009EE0)

    Absolute widths are easy to get wrong by the width of one column, and the
    result is a last column over the edge of the slide that every text check
    passes.  Proportions cannot overflow.
    """
    # ``x`` is where a Box naturally lands when the call is written positionally.
    if isinstance(x, Box):
        box, x = x, None
    head_color = kw.pop("head_text_color", head_color)   # the other obvious name
    if kw:
        raise TypeError(
            "table(): unknown keyword(s) %s.  Colours are head_fill, "
            "head_color, body_fill, zebra, rule and cell_colors={(row, col): "
            "colour}; sizes are size, head_size, row_h, header_h, row_heights."
            % ", ".join(sorted(kw)))
    if box is not None:
        x = box.x if x is None else x
        y = box.y if y is None else y
        total = float(sum(widths)) or 1.0
        widths = [box.w * (w / total) for w in widths]
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


def _clear_push(ux, uy, r, w, h):
    """Distance along the unit ray ``(ux, uy)`` at which a ``w`` x ``h`` label
    centred there stops overlapping a circle of radius ``r``.

    Pushing a fixed multiple of ``r``, as this used to, is what puts a wide
    horizontal label back inside the circle it belongs to -- the label extends
    sideways far further than the ray travels.  Solved instead from the
    rectangle's nearest point to the circle's centre, which is exact and needs
    no cases for two- versus three-circle figures.

    ``w`` and ``h`` are the *text's* extent, not its box: a centred label in a
    generous box would otherwise be shoved out by its own empty margins.
    """
    hw, hh = w / 2, h / 2
    gap = r * 1.06                     # a hair of air, so it does not just kiss
    lo, hi = r * 0.5, r * 4
    for _ in range(24):
        mid = (lo + hi) / 2
        dx = max(0.0, abs(ux * mid) - hw)
        dy = max(0.0, abs(uy * mid) - hh)
        if dx * dx + dy * dy >= gap * gap:
            hi = mid
        else:
            lo = mid
    return hi


def _venn_labels(cx, cy, r, n, labels, label_size=None):
    """Where each outer label goes: ``[(x, y, w, h, size, lines), ...]``.

    Pure geometry, so ``venn_bounds`` can ask the same question without drawing
    anything -- which is what lets ``venn(box=...)`` size itself to fit.
    """
    out = []
    box_w, box_h = r * 1.7, r * 0.62
    for i, label in enumerate(labels[:n]):
        # Push the label out along the ray from the figure's centre through this
        # circle's centre, so it clears the circle instead of landing inside it
        # -- and so two- and three-circle figures need no cases.
        ex, ey = venn_geometry(cx, cy, r, n)[i]
        dx, dy = ex - cx, ey - cy
        norm = (dx * dx + dy * dy) ** 0.5 or 1
        ux, uy = dx / norm, dy / norm
        # A diagram label belongs to the deck's type scale, not to the figure's
        # radius: sizing it as a fraction of r gives a large Venn 28pt labels
        # that dwarf a 16pt slide title and shove each other onto the footnote.
        # Fit it to the box, but never above body-plus-a-bit.
        size, lines = fit_lines(label, box_w, box_h, BOLD,
                               max_size=LABEL_MAX_PT)
        if label_size:
            size, lines = label_size, wrap_words(label, box_w, label_size, BOLD)
        ink_w = max(text_width_est(ln, size, BOLD) for ln in lines)
        ink_h = len(lines) * size * 35.28 * 1.22
        push = _clear_push(ux, uy, r, ink_w, ink_h)
        sx, sy = ex + ux * push, ey + uy * push
        out.append((sx - box_w / 2, sy - box_h / 2, box_w, box_h, size, lines))
    return out


def venn_bounds(cx, cy, r, n=3, labels=None, label_size=None):
    """The ``Box`` a ``venn`` would actually occupy, labels included.

    The circles alone span ``3r`` by ``2.87r``; the labels take it well past
    that, which is how a figure that looked sized to its slot ends up sitting
    on the footnote underneath.  Measure rather than assume -- or pass ``box``
    to ``venn`` and let it solve for ``r``.
    """
    c = venn_geometry(cx, cy, r, n)
    xs = [e[0] - r for e in c] + [e[0] + r for e in c]
    ys = [e[1] - r for e in c] + [e[1] + r for e in c]
    for lx, ly, lw, lh, size, lines in _venn_labels(cx, cy, r, n, labels or [],
                                                    label_size):
        # The box is wider than its centred text; bound the ink, not the box.
        ink_w = max(text_width_est(ln, size, BOLD) for ln in lines)
        ink_h = len(lines) * size * 35.28 * 1.22
        xs += [lx + lw / 2 - ink_w / 2, lx + lw / 2 + ink_w / 2]
        ys += [ly + lh / 2 - ink_h / 2, ly + lh / 2 + ink_h / 2]
    return Box(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))


def venn_fit(box, n=3, labels=None, label_size=None):
    """``(cx, cy, r)`` for the largest ``venn`` whose labels still fit ``box``.

    ``D.venn(doc, page, *D.venn_fit(area, labels=names), colors, labels=names)``
    -- or just pass ``box=area`` to ``venn``, which does this for you.
    """
    lo, hi = min(box.w, box.h) / 60.0, min(box.w, box.h)
    for _ in range(30):
        r = (lo + hi) / 2
        b = venn_bounds(box.cx, box.cy, r, n, labels, label_size)
        if b.w <= box.w and b.h <= box.h:
            lo = r
        else:
            hi = r
    r = lo
    # Re-centre: the labelled figure is not symmetric about the circles.
    b = venn_bounds(box.cx, box.cy, r, n, labels, label_size)
    return (int(box.cx + (box.cx - b.cx)), int(box.cy + (box.cy - b.cy)), int(r))


def venn(doc, page, cx=None, cy=None, r=None, colors=(), labels=None,
         transparence=76, line_width=45, label_size=None, label_color=None,
         centre=None, centre_color=None, centre_text=None, centre_size=None,
         box=None):
    """Two or three overlapping circles.  Returns the centres, so you can
    label the regions yourself.

    ``colors`` is one colour per circle and decides how many are drawn --
    pass a prefix of it to build the diagram up over successive slides.
    ``labels``, if given, are placed just outside each circle.
    ``centre`` fills the all-way overlap with that colour.

    **Give it the area, not a radius** -- a ``Box``, positionally or as
    ``box=``, typically the one ``canvas`` returned::

        D.venn(doc, page, area, [0x009EE0, 0xE8730C, 0x6B3FA0],
               labels=["PRIVATE DATA", "UNTRUSTED CONTENT", "EXTERNAL COMMS"],
               centre=0xC00000, centre_text="DATA EXFILTRATION")

    The figure then sizes and centres itself to fill it, labels included.  A
    radius picked by eye -- ``mm(40)`` -- either wastes half the area or pushes
    the labels onto whatever else is on the slide, and no automated check sees
    the second one.
    """
    global _drawing_venn
    if isinstance(cx, Box):
        # `venn(doc, page, area, colors, labels)` -- one box stands in for all
        # three of cx, cy, r, so any positional after it has shifted down two.
        box = cx
        if colors == () and cy is not None:
            colors = cy
        if labels is None and r is not None:
            labels = r
        cx = cy = r = None
    n = len(colors)
    if box is not None:
        cx, cy, r = venn_fit(box, n, labels, label_size)
    centres = venn_geometry(cx, cy, r, n)
    _drawing_venn = True
    try:
        for i, color in enumerate(colors):
            ex, ey = centres[i]
            ellipse(doc, page, ex, ey, r, fill=color,
                    transparence=transparence, line=color,
                    line_width=line_width)
    finally:
        _drawing_venn = False
    if labels:
        for i, (lx, ly, lw, lh, size, lines) in enumerate(
                _venn_labels(cx, cy, r, n, labels, label_size)):
            text(doc, page, lx, ly, lw, lh, *lines, size=size, weight=BOLD,
                 align="center", valign="center",
                 color=(label_color or colors[i]))
    if centre is not None:
        joined = " ".join(str(p) for p in (
            [centre_text] if isinstance(centre_text, str) else (centre_text or [])
        ))
        # 0.41 r is the inradius of the three-way overlap, so the disc sits
        # exactly inside it.  A long label at that size shrinks to nothing --
        # better to let the disc grow a little into the pair regions than to
        # set "DATA EXFILTRATION" at 7pt.  Capped where it would start hiding
        # the two-way overlaps that the figure exists to show.
        cr = r * 0.41
        if joined and not centre_size:
            while cr < r * 0.54 and fit_lines(
                    joined, cr * 1.38, cr * 1.38, BOLD,
                    max_size=CENTRE_MAX_PT)[0] < 11.0:
                cr *= 1.04
        ellipse(doc, page, cx, cy, cr, fill=centre, line=centre,
                line_width=line_width)
        if joined:
            # The label has to fit *inside a circle*, so the usable box is the
            # inscribed square, not the diameter.  Wrap on spaces and take the
            # size from the longest word: sizing the whole string to one line
            # makes "DATA EXFILTRATION" unreadably small, and sizing it without
            # regard to wrapping breaks "FORBIDDEN" mid-word.
            box_w = box_h = cr * 1.38
            size, lines = fit_lines(joined, box_w, box_h, BOLD,
                                    max_size=CENTRE_MAX_PT)
            if centre_size:
                size = centre_size
                lines = wrap_words(joined, box_w, size, BOLD)
            text(doc, page, cx - box_w / 2, cy - box_h / 2, box_w, box_h,
                 *lines, size=size,
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

    **Returns the free area as a** ``Box``, measured from the shapes it just
    removed -- the body placeholder's own rectangle, which is the designer's
    answer to where content goes on this template.  Draw inside it and the
    margins match the rest of the deck::

        area = canvas(page, keep=(0, 1, 2, 3))
        left, right = area.cols(2, gap=mm(8))

    Tiny shapes (a think-cell stub is 3 x 3) are ignored when measuring, and if
    nothing substantial was removed the page inset by the kept chrome's own
    left margin is returned instead.
    """
    current = shapes(page)
    if keep:
        doomed = [s for i, s in enumerate(current) if i not in set(keep)]
    else:
        doomed = [current[i] for i in remove if 0 <= i < len(current)]

    spans = []
    for shape in doomed:
        try:
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if size.Width > page.Width // 20 and size.Height > page.Height // 20:
            spans.append((pos.X, pos.Y, pos.X + size.Width, pos.Y + size.Height))

    for shape in doomed:
        page.remove(shape)

    if spans:
        return Box(min(s[0] for s in spans), min(s[1] for s in spans),
                   max(s[2] for s in spans) - min(s[0] for s in spans),
                   max(s[3] for s in spans) - min(s[1] for s in spans))

    # Nothing sizeable was removed: fall back to the page, inset by whatever
    # margin the surviving chrome sits on, and below the lowest of it.
    left, top = page.Width // 16, page.Height // 8
    for shape in shapes(page):
        try:
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if size.Width < page.Width // 20:
            continue
        left = min(left, pos.X)
        if pos.Y + size.Height < page.Height // 2:
            top = max(top, pos.Y + size.Height)
    return Box(left, top, page.Width - 2 * left,
               page.Height - top - page.Height // 12)


__all__ = [
    "Point", "Size", "BLACK", "WHITE", "GREY", "GREY_DK", "GREY_BD",
    "GREY_MD", "GREY_LT", "BOLD", "NORM", "THIN", "DEFAULT_FONT",
    "rect", "ellipse", "line", "hline", "vline", "polygon", "arrow",
    "chevron", "text", "chip", "bullets", "table", "venn", "venn_geometry",
    "venn_fit", "venn_bounds",
    "pair_spots", "restyle", "style_fill", "walk", "walk_page", "move",
    "resize", "canvas", "est_lines", "set_text", "Box", "Split", "page_box",
    "advance_em", "text_width_est", "fit_size", "fit_lines", "wrap_words",
    "contrast_ratio", "bg_color", "readable_on",
]
