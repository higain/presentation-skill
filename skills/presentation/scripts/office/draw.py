"""Vector shapes on a slide: rectangles, ellipses, lines, arrows, text, tables.

``impress.py`` opens and retexts documents; this draws on them.  Without it the
only shape you can create is a text box, and a deck that needs a diagram gets a
bullet list instead -- which is the single largest quality difference between
runs of this skill.

Every function takes ``(doc, page, ...)`` and returns the UNO shape it made, so
you can keep styling it afterwards.  Geometry is **1/100 mm** throughout, like
the rest of the skill; ``mm()`` and ``pt()`` from ``office.impress`` convert.
Colours are plain ints: ``0x2A7DE1``.

    import sys; sys.path.insert(0, "<skill>/scripts")
    from office.impress import impress_document, slides, mm
    from office import draw as D

    with impress_document("deck.pptx") as doc:
        page = slides(doc)[0]
        D.rect(doc, page, mm(20), mm(40), mm(60), mm(30), fill=0x2A7DE1)
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
import inspect
import sys

import uno

from .impress import shapes
# Re-exported so a drawing script needs one import: these edit shapes that are
# already on the slide, which is the other half of the same job.
from .template import (  # noqa: F401
    move, palette, palette_roles, resize, set_text, text_width, type_scale,
    walk_shapes,
)
from .template import _is_group, _walk as _template_walk

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

# Below this, text is legible on the monitor it was built on and not from a
# seat in the room.  The floor of *last resort*: it applies to a deck whose
# template nothing is known about.  Where a profile exists the number comes
# from the template instead -- see ``min_readable`` -- because 9pt is one
# company's smallest role and another's is 12.
MIN_READABLE_PT = 9.0

# The sizes a build script would otherwise pick one at a time.  Read from the
# template profile, which counts them off the asset; empty for a deck with no
# template, where nothing is stopping you from choosing a scale of your own --
# see references/design.md -- as long as it *is* a scale.
_SCALE_FALLBACK = {"display": 40.0, "title": 20.0, "subtitle": 16.0,
                   "body": 14.0, "label": 12.0, "caption": 10.0,
                   "footnote": 9.0}


def min_readable(template=None):
    """The smallest type size this deck's template itself sets, in points.

    A floor taken from the design rather than invented here: text under the
    smallest size a designer ever used is text the designer would not have set.
    ``MIN_READABLE_PT`` when nothing is known about the template, which is the
    honest answer for a deck built from nothing.

    Deliberately not capped.  A template whose smallest role is 18pt really
    does mean that 12pt is off-style on it, and a check that quietly clamped to
    9 would report nothing on exactly the templates with the most to say.
    """
    scale = type_scale(template=template)
    return min(scale.values()) if scale else MIN_READABLE_PT


def size_for(role):
    """The point size of a named role in the deck's type scale.

    ``D.text(..., role="footnote")`` is the usual way in; this is for the
    arithmetic around it -- how tall a band has to be, what a caption under a
    figure should match.  Unknown names raise with the list of real ones,
    because a silent default is how a scale stops being a scale.
    """
    scale = type_scale() or _SCALE_FALLBACK
    if role not in scale:
        raise ValueError(
            "size_for(): no role %r in this deck's type scale.\n"
            "  It has: %s\n"
            "  Or pass size= in points and skip the scale."
            % (role, ", ".join(sorted(scale))))
    return float(scale[role])


def color(name):
    """A colour from the template's own palette, by name.

    ``D.color("cyan")``, ``D.color("accent")``, ``D.color("alert")``.  Unknown
    names raise with the list, and a deck built from no template has no palette
    at all -- read the colours off the deck with ``probe.py`` rather than
    inventing them, which is how four decks went out in a corporate orange the
    corporation does not use.
    """
    names = palette()
    if name in names:
        return names[name]
    roles = palette_roles()
    if name in roles and not isinstance(roles[name], list):
        return roles[name]
    raise ValueError(
        "color(): no colour %r in this template's palette.\n"
        "  Colours: %s\n"
        "  Roles: %s\n"
        "  %s"
        % (name, ", ".join(sorted(names)) or "(none -- this deck's template "
                                             "has no pack)",
           ", ".join(sorted(k for k, v in roles.items()
                            if not isinstance(v, list))) or "(none)",
           "series() gives the ordered colours for the nth thing on a figure."
           if names else
           "Run the `presentation-setup` skill on this template to count its "
           "palette, or read the colours off the deck with probe.py -- do not "
           "invent one."))


def series(n=None):
    """The palette's colours in the order to use them, longest first.

    ``D.series(3)`` for a three-lobed Venn, three cards, three bars.  Cycles if
    you ask for more than the palette has, so the fourth thing is not silently
    black.
    """
    colours = palette_roles().get("series") or list(palette().values())
    if not colours:
        raise ValueError(
            "series(): this deck has no palette -- there is no template "
            "profile to take one from. Pick the colours yourself (probe.py "
            "prints the ones a deck already uses) or see references/design.md.")
    if n is None:
        return list(colours)
    return [colours[i % len(colours)] for i in range(n)]


# ------------------------------------------------------------------- layout

_SPLIT_HELP = (
    "cut_top/cut_bottom/cut_left/cut_right return a PAIR of boxes, not a box.\n"
    "  unpack it:      head, body = area.cut_top(mm(14))\n"
    "  or take one:    area.top(mm(14))    area.bottom(mm(14))\n"
    "                  area.left(mm(60))   area.right(mm(60))\n"
    "Those return a single Box and cannot come out in the wrong order."
)

# What a caller reached for on a Box, and what this one is called.  Every entry
# is a name an agent run actually wrote before the build died on a bare
# "'Box' object has no attribute ...", which says nothing about what to write
# instead -- the kwarg table below does the same job for keywords, and the
# misses it covered cost one build each against three or four for the misses
# it did not.
_BOX_ALIASES = {
    "move": "inset(left=…, top=…), or Box(box.x + dx, box.y + dy, box.w, box.h)",
    "shift": "inset(left=…, top=…), or build a Box at the offset you want",
    "offset": "inset(…) -- it takes left/top/right/bottom separately",
    "translate": "inset(left=…, top=…)",
    "pad": "inset(all=mm(6))",
    "shrink": "inset(all=mm(6))",
    "split": "cols(n, gap=…), rows(n, gap=…) or grid(rows, cols, gap=…)",
    "split_vertical": "rows(n, gap=…) -- n boxes stacked down the page",
    "split_horizontal": "cols(n, gap=…) -- n boxes across the page",
    "split_v": "rows(n, gap=…)",
    "split_h": "cols(n, gap=…)",
    "columns": "cols(n, gap=…)",
    "column": "cols(n, gap=…)[i]",
    "row": "rows(n, gap=…)[i]",
    "divide": "cols(n, gap=…) or rows(n, gap=…)",
    "subdivide": "grid(rows, cols, gap=…)",
    "width": "w",
    "height": "h",
    "left_edge": "x",
    "right_edge": "x2",
    "top_edge": "y",
    "bottom_edge": "y2",
    "center": "centre, or box.cx / box.cy",
    "center_x": "cx",
    "center_y": "cy",
    "middle": "centre",
    "area": "box.w * box.h",
    "band": "cut_top(…).band -- a Box is not a pair, it has no band",
    "rest": "cut_top(…).rest -- a Box is not a pair, it has no rest",
}

_BOX_HELP = (
    "  parts:     box.x  box.y  box.w  box.h  box.x2  box.y2  box.cx  box.cy"
    "  box.centre\n"
    "  slice:     box.cols(3, gap=mm(6))   box.rows(2)   box.grid(2, 3)\n"
    "             box.top(mm(14))   box.bottom(mm(14))   box.left(mm(60))"
    "   box.right(mm(60))\n"
    "             box.inset(all=mm(6))   box.square()\n"
    "  pairs:     head, body = box.cut_top(mm(14), gap=mm(6))   "
    "-- also cut_bottom/cut_left/cut_right"
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

    def __getitem__(self, index):
        # `canvas(page)[1]` and `area[0]` both get written, by callers who
        # expect the pair a cut_* returns.  Python's "'Box' object is not
        # subscriptable" says nothing about what to write instead.
        raise TypeError(
            "a Box is one rectangle, not a sequence -- there is no [%r].\n"
            "  Its parts are named:   box.x  box.y  box.w  box.h  "
            "box.x2  box.y2  box.cx  box.cy\n"
            "  Pass it whole:         D.rect(doc, page, box, fill=0xEEEEEE)\n"
            "  Slice it for pieces:   box.top(mm(14))   box.cols(3, gap=mm(6))"
            "   head, body = box.cut_top(mm(14))\n"
            "  canvas() returns ONE Box, not a pair." % (index,))

    def __repr__(self):
        return "Box(x=%d, y=%d, w=%d, h=%d)  # %.0f,%.0fmm %.0fx%.0fmm" % (
            self.x, self.y, self.w, self.h,
            self.x / 100, self.y / 100, self.w / 100, self.h / 100)

    def __getattr__(self, name):
        # Only reached on a miss -- __slots__ and the properties are found
        # first.  Dunders go back the normal way: copy, pickle and inspect all
        # probe for methods that are meant to be absent, and answering those
        # with a wall of layout advice breaks them.
        if name.startswith("__") and name.endswith("__"):
            raise AttributeError(name)
        hint = _BOX_ALIASES.get(name)
        raise AttributeError(
            "a Box has no %r.%s\n%s"
            % (name, ("  Use %s" % hint) if hint else "", _BOX_HELP))

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
                # A box AND a coordinate keyword: the box already supplies all
                # four, so Python would raise "got multiple values for argument
                # 'y'" from inside the callee, naming an argument the caller
                # never wrote positionally.  Say what actually happened.
                clash = [n for n in names if n in kw]
                if clash:
                    raise TypeError(
                        "%s(): got the area as a Box and %s= as well.  The box "
                        "already supplies %s -- there is nothing left for %s to "
                        "mean.\n"
                        "  To move or resize the area, slice the box:\n"
                        "      %s(doc, page, area.bottom(mm(20)), ...)   # a band\n"
                        "      %s(doc, page, area.inset(mm(6)), ...)     # padded\n"
                        "      %s(doc, page, D.Box(area.x, area.y + mm(60), "
                        "area.w, mm(20)), ...)\n"
                        "  Or drop the box and pass all four numbers."
                        % (fn.__name__, "=, ".join(clash), ", ".join(names),
                           clash[0], fn.__name__, fn.__name__, fn.__name__))
                box = args[at]
                args = (args[:at] + tuple(_BOX_PARTS[n](box) for n in names)
                        + args[at + 1:])
            _check_kwargs(fn, kw)
            return fn(*args, **kw)
        return wrapper
    return decorate


# Names a caller reaches for that this module spells differently.  Each one was
# written by a model building a real deck; the bare Python error ("unexpected
# keyword argument 'opacity'") says which word is wrong and never which word is
# right, so the next guess is as likely to be wrong as the first.
_KWARG_ALIASES = {
    "opacity": "transparence= (0-100, opaque to invisible -- not a 0-1 alpha)",
    "alpha": "transparence= (0-100, opaque to invisible)",
    "transparency": "transparence=",
    "colour": "color=",
    "fill_color": "fill=",
    "fill_colour": "fill=",
    "fillcolor": "fill=",
    "background": "fill=",
    "bg": "fill=",
    "line_color": "line=",
    "linecolor": "line=",
    "border": "line=",
    "border_color": "line=",
    "border_width": "line_width=",
    "stroke": "line=",
    "stroke_width": "line_width=",
    "font_size": "size=",
    "fontsize": "size=",
    "font_name": "font=",
    "bold": "weight=D.BOLD",
    "corner_radius": "radius=",
    "rounded": "radius= (in 1/100 mm)",
    "padding": "pad=",
    "text_color": "color=",
    "halign": "align=",
    "vertical_align": "valign=",
    "row_height": "row_h=",
    "header_fill": "head_fill=",
    "columns": "widths= (proportions, with the area as the box)",
    # There is no "nudge it a bit" keyword anywhere in here: the geometry comes
    # from the Box, so an offset is a different Box.  Every one of these was
    # written by a run that had a Box in hand and reached for a keyword.
    "y_offset": "nothing -- move the Box instead: box.inset(top=mm(8)), or "
                "head, body = box.cut_top(mm(8))",
    "x_offset": "nothing -- move the Box instead: box.inset(left=mm(8))",
    "offset": "nothing -- move the Box instead: box.inset(left=…, top=…)",
    "dx": "nothing -- move the Box instead: box.inset(left=…)",
    "dy": "nothing -- move the Box instead: box.inset(top=…)",
    "top": "nothing -- the y comes from the Box you pass",
    "margin": "pad= (inside the shape), or box.inset(all=…) (outside it)",
    "line_spacing": "nothing -- one string per paragraph, not a spacing knob",
    "italics": "italic=",
    "underlined": "underline=",
    "strikethrough": "strike=",
    "weight_bold": "weight=D.BOLD",
    "font_weight": "weight= (D.BOLD or D.NORM)",
    "font_family": "font=",
    "text_align": "align=",
    "justify": "align='justify'",
    "auto_fit": "grow= (autogrow height), or D.fit_size to pick the size",
    "shrink_to_fit": "nothing here -- D.fit_size/D.fit_lines measure, and "
                     "office.template.shrink_to_fit resizes a template's shape",
}


@functools.lru_cache(maxsize=None)
def _accepted(fn):
    params = inspect.signature(fn).parameters
    if any(p.kind is inspect.Parameter.VAR_KEYWORD for p in params.values()):
        return None                      # takes **kw and checks them itself
    return frozenset(params)


def _check_kwargs(fn, kw, label=None, extra=()):
    """Refuse an unknown keyword with the name that was meant instead.

    ``label`` names the function the caller actually wrote, and ``extra`` are
    keywords it consumed before delegating -- both for a wrapper like
    ``text()``, which pops its own and hands the rest to ``_char_props``.
    """
    accepted = _accepted(fn)
    if accepted is None:
        return
    accepted = accepted | frozenset(extra)
    fn_name = label or fn.__name__
    unknown = [k for k in kw if k not in accepted]
    if not unknown:
        return
    lines = []
    for name in unknown:
        hint = _KWARG_ALIASES.get(name.lower())
        lines.append("  %s= -> %s" % (name, hint) if hint
                     else "  %s= is not one of them" % name)
    raise TypeError(
        "%s(): unknown keyword(s) %s.\n%s\n  It accepts: %s"
        % (fn_name, ", ".join(sorted(unknown)), "\n".join(lines),
           ", ".join(sorted(n for n in accepted
                            if n not in ("doc", "page")))))


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

        D.rect(doc, page, area.top(mm(14)), fill=0x2A7DE1)
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

    D.venn(doc, page, area, D.series(3),
           labels=["PRIVATE DATA", "UNTRUSTED CONTENT", "EXTERNAL COMMS"],
           centre=D.color("alert"), centre_text="EXFILTRATION")

venn() places the circles so they actually overlap, sizes the figure to the
area, and puts each label against its own lobe.  Hand-placed circles come out
tangent or overlapping their labels, and no automated check catches either --
that is a rendering defect, and it needs eyes.  D.venn_fit(box) if you want the
geometry without the drawing."""

_VENN_RADIUS_HELP = """\
[draw] D.venn() was given a radius rather than an area.  Pass the Box:

    D.venn(doc, page, area, D.series(3), labels=[...])

With the area it solves for the largest figure whose labels still fit and
centres the result; with a radius it draws exactly what you asked for, and what
gets asked for is a fraction of the space -- one run's hero slide put a Venn in
a quarter of its canvas and shipped, because nothing that cannot see the slide
can tell a small diagram from a small diagram on purpose."""


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

        D.ellipse(doc, page, area.square(0.6), fill=0x2A7DE1)

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

    ``role`` takes the size from the deck's own type scale --
    ``role="footnote"``, ``role="body"``, ``role="caption"`` -- so a footnote is
    the same size on every slide of the deck instead of 9pt here and 12pt three
    slides later.  ``D.size_for(role)`` is the number, for the arithmetic
    around it; an explicit ``size=`` still wins.

    Keywords: ``role``, ``size`` (pt), ``color``, ``weight`` (``BOLD``/``NORM``),
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

    _OWN = ("fill", "line", "transparence", "line_width", "pad", "valign",
            "wrap", "grow", "role")
    # Everything not consumed here goes to _char_props, and an unknown name
    # would raise from in there, three frames down, naming a private function.
    _check_kwargs(_char_props, kw, label="text", extra=_OWN)

    # A named size from the deck's scale.  An explicit size= still wins, so a
    # figure that has measured its own label is not overruled.
    role = kw.pop("role", None)
    if role is not None and kw.get("size") is None:
        kw["size"] = size_for(role)

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

    # Before the string, not after it.  A new TextShape grows its own height by
    # default, and a shape that has already grown does not shrink back when the
    # property is turned off a line later -- so the box a caller measured, cut
    # and handed over came back a different size, and everything positioned
    # against its floor (a card's note, the row under it) was left hanging
    # outside.  Set here, the box keeps the height it was given, and text that
    # does not fit is reported by qa.py rather than silently rearranging the
    # slide.
    s.TextAutoGrowHeight = bool(grow)
    s.TextAutoGrowWidth = False
    s.TextWordWrap = bool(wrap)

    # Text after: setString() on an empty shape re-inherits the drawing
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
# An em dash is a full em and an en dash half of one; counted as ordinary
# letters they were the difference between a line that fits and a line
# LibreOffice re-broke, on the one slide made of a single sentence.
_EM = {"—": 1.0, "–": 0.5, "…": 1.0}


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
        elif ch in _EM:
            total += _EM[ch]
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

MIN_COL_W = 800          # 8 mm: narrower than this and a word breaks mid-word


def _column_widths(widths, box, x, page):
    """Column widths in 1/100 mm, and the check that they can possibly work.

    With a ``box`` the widths are proportions and cannot overflow.  Without
    one they are absolute, and absolute widths are the one place in this module
    where a plausible-looking number silently wrecks the slide: ``[80, 60, 40,
    120]`` reads like points or pixels, is 0.8 mm to 1.2 mm here, and produces
    a table every label breaks mid-word inside.  Measured on one such build:
    54 QA findings, all of them from two calls shaped like that.

    So absolute widths are checked against the slide before anything is drawn.
    """
    if box is not None:
        total = float(sum(widths)) or 1.0
        return [box.w * (w / total) for w in widths]

    widths = [_num(w, "widths[%d]" % i, "table") for i, w in enumerate(widths)]
    span = sum(widths)
    slide_w = getattr(page, "Width", 0) or 0
    narrow = [i for i, w in enumerate(widths) if w < MIN_COL_W]
    if narrow or (slide_w and span < slide_w * 0.2):
        raise ValueError(
            "table(): these column widths cannot hold text -- %s "
            "(%s in 1/100 mm, so column %d is %.1f mm wide).  Widths here are "
            "1/100 mm like every other coordinate, not points and not "
            "percentages.\n"
            "  Pass the area and give widths as PROPORTIONS -- they are then "
            "scaled to fit it exactly and cannot overflow:\n"
            "      D.table(doc, page, area, widths=[3, 2, 2, 5], rows=rows)\n"
            "  Or, for absolute widths, use mm(): widths=[mm(60), mm(40)]."
            % (widths, span, (narrow or [0])[0] + 1,
               widths[(narrow or [0])[0]] / 100.0 if widths else 0.0))
    if slide_w and x is not None and x + span > slide_w * 1.02:
        raise ValueError(
            "table(): the columns run %.0f mm past the right edge of the slide "
            "(x=%d + widths=%d against a slide %d wide).  Nothing would report "
            "this -- every cell holds its own text perfectly, the last column "
            "is simply over the trim.\n"
            "  Pass the area and give widths as proportions instead:\n"
            "      D.table(doc, page, area, widths=[3, 2, 2, 5], rows=rows)"
            % ((x + span - slide_w) / 100.0, x, span, slide_w))
    return widths


@takes_box("x", "y", "w", "h")
def table(doc, page, x=None, y=None, w=None, h=None, widths=(), rows=(),
          row_h=None, header_h=None,
          head_fill=GREY_DK, head_color=WHITE, size=12.0, head_size=12.0,
          zebra=GREY_LT, body_fill=WHITE, align=None, pad=200,
          cell_colors=None, cell_weights=None, row_heights=None,
          rule=GREY_BD, font=DEFAULT_FONT, box=None, **kw):
    """Draw a table out of rectangles and text.  ``rows[0]`` is the header.

    A drawn table rather than a real ``TableShape``: it survives a .pptx
    round-trip unchanged, takes per-cell colour without fighting a table style,
    and needs no column-model bookkeeping.  Use a real table only when the
    recipient has to edit it as one.

    **Give it the area and proportional widths**, like every other call here::

        D.table(doc, page, area, widths=[3, 2, 2, 5], rows=rows,
                head_fill=0x2A7DE1)

    The columns are then scaled to the box exactly, ``row_h`` defaults to the
    box height divided between the rows, and nothing can land off the slide.
    Absolute widths still work -- ``widths=[mm(60), mm(40), mm(40)]`` with an
    ``x, y`` -- but they are checked against the slide first, because a number
    that looks like points is 25x too small and every cell then breaks its
    own label mid-word.

    ``cell_colors``/``cell_weights`` are ``{(row, col): value}`` with row 1 the
    first body row.  Returns the y just past the last row.
    """
    # The pre-box signature was ``table(doc, page, x, y, widths, rows, ...)``,
    # so a sequence in the ``w`` slot is that call, not a width.
    if isinstance(w, (list, tuple)):
        widths, rows, w, h = w, (h if h is not None else rows), None, None
    if isinstance(x, Box):                    # box passed as a bare positional
        box, x = x, None
    head_color = kw.pop("head_text_color", head_color)   # the other obvious name
    if kw:
        raise TypeError(
            "table(): unknown keyword(s) %s.  Colours are head_fill, "
            "head_color, body_fill, zebra, rule and cell_colors={(row, col): "
            "colour}; sizes are size, head_size, row_h, header_h, row_heights."
            % ", ".join(sorted(kw)))
    if box is None and None not in (x, y, w, h):
        box = Box(x, y, w, h)                 # the @takes_box form
    if not rows:
        raise ValueError("table(): no rows to draw")
    if not widths:
        widths = [1] * len(rows[0])           # even columns, the obvious default
    if box is not None:
        x = box.x if x is None else x
        y = box.y if y is None else y
    if box is None and (x is None or y is None):
        raise TypeError(
            "table(): nowhere to draw it.  Pass the area -- "
            "D.table(doc, page, area, widths=[3, 2, 2], rows=rows) -- or an "
            "x and y with absolute widths.")
    widths = _column_widths(widths, box, x, page)
    if row_h is None and row_heights is None:
        if box is None:
            raise TypeError(
                "table(): give it the area -- D.table(doc, page, area, "
                "widths=[3, 2, 2], rows=rows) -- or say how tall a row is "
                "with row_h=mm(12).")
        # Fill the box: a table drawn to two-thirds of its area and a dead
        # band underneath is the commonest composition defect in a built deck.
        row_h = box.h / max(len(rows), 1)
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


# ------------------------------------------------------------ composition
#
# Four layouts, because four builds of the same plan produced four hand-rolled
# versions of each and every one of them got the rhythm wrong somewhere: cards
# whose content sat in the top third with the bottom half bare, step arrows
# drawn as 5 mm stubs between tiles, a statement slide hand-broken with "\n"
# into "it --", "an", "email.", and one genuinely good row of big numbers that
# existed only because a model chose to build it out of rect and text.
#
# None of this is beyond a build script.  All of it is arithmetic nobody should
# do twice, and a slide's rhythm is exactly the thing that comes out different
# every time it is done by hand.

#: Inner margin of a card, and the gap between cards, when none is given.
CARD_PAD, CARD_GAP = 500, 600

#: Below this ratio an accent colour is used as a fill, never as text.  Same
#: number qa.py fails a build on, checked here so the build never makes the
#: mistake rather than reporting it afterwards.
MIN_ACCENT_CONTRAST = 3.0

#: Line widths are estimated to a few percent, and LibreOffice lays the text
#: out for real.  Wrapping to the full box means every line that lands within a
#: percent of it gets re-broken on the slide -- so a statement balanced to
#: three lines renders as four, with the orphan this exists to prevent.
WRAP_SLACK = 0.94


def _as_item(item):
    """One card's content as a dict, from whatever shape it arrived in."""
    if isinstance(item, dict):
        return dict(item)
    if isinstance(item, str):
        return {"body": item}
    parts = list(item) + [None] * 3
    return {"title": parts[0], "body": parts[1], "note": parts[2]}


def _head_metrics(item, inner_w, title_size, numbered):
    """``(chip_w, chip_gap, text_w, band_h, gap_h)`` for a card's head band.

    One function, used by the measuring pass and the drawing pass alike.  When
    they disagreed -- the measure assuming a one-line title, the drawing
    wrapping it to two -- every card whose title wrapped grew past the height
    it had been given, and the note pinned to its floor was printed outside the
    panel.
    """
    line_h = int(title_size * 35.28 * 1.35)
    chip_w = line_h if numbered else 0
    chip_gap = int(chip_w * 0.45)
    text_w = inner_w - chip_w - chip_gap
    title_h = (_text_height([item["title"]], text_w, title_size, BOLD)
               if item.get("title") else line_h)
    return chip_w, chip_gap, text_w, max(line_h, title_h), \
        int(title_size * 35.28 * 0.35)


def _card_height(item, w, title_size, size, note_size, pad, numbered):
    """Height this card needs for its own content, in 1/100 mm."""
    inner = w - 2 * pad
    total = 2 * pad
    if numbered or item.get("title"):
        _, _, _, band_h, gap_h = _head_metrics(item, inner, title_size, numbered)
        total += band_h + gap_h
    if item.get("body"):
        total += _text_height([item["body"]], inner, size)
    if item.get("note"):
        total += int(note_size * 35.28 * 1.5)
    return total


#: How much of its panel the tallest card's content may fill when ``cards``
#: grows the type.  The rest is air under the text -- enough that the panel
#: reads as a panel, not so much that it reads as empty.
CARD_FILL = 0.6

#: ``height="auto"``: the share of a card its text should fill, and the
#: least share of the area the row should take.
CARD_TEXT_SHARE = 0.75
CARD_MIN_SHARE = 0.5

#: ``(body, title)`` roles ``cards`` tries, smallest first.
CARD_TYPE_LADDER = (("label", "subtitle"), ("body", "subtitle"),
                    ("subtitle", "title"))


def _longest_word_fits(items, inner, size, title_size):
    for item in items:
        for key, pt, weight in (("body", size, NORM),
                                ("title", title_size, BOLD)):
            for word in str(item.get(key) or "").split():
                if text_width_est(word, pt, weight) > inner:
                    return False
    return True


def _grow_card_type(items, cell_w, offered, size, title_size, note_size, pad,
                    numbered):
    """``(size, title_size)``: the largest rung of the scale that fits well.

    A size the caller passed is kept; only the missing one is chosen.
    """
    chosen = None
    for body_role, title_role in CARD_TYPE_LADDER:
        try:
            trial = (size if size is not None else size_for(body_role),
                     title_size if title_size is not None
                     else size_for(title_role))
        except ValueError:                   # a scale without that role
            continue
        if chosen is None:
            chosen = trial                   # the smallest is the floor
            continue
        needed = max(_card_height(item, cell_w, trial[1], trial[0],
                                  note_size, pad, numbered) for item in items)
        if (needed > offered * CARD_FILL
                or not _longest_word_fits(items, cell_w - 2 * pad, *trial)):
            break
        chosen = trial
    return chosen


def _draw_card(doc, page, box, item, fill, accent, number, title_size, size,
               note_size, pad, radius, color, title_color, note_color, line):
    """One card: panel, optional number chip, title, body, and a pinned note."""
    if fill is not None or line is not None:
        rect(doc, page, box, fill=fill, radius=radius, line=line)
    on = fill if fill is not None else bg_color(page)
    inner = box.inset(pad)

    # The accent titles a card only if it can be read on the card.  Brand
    # colours are chosen against white, and the panel fill is not white: the
    # house cyan on #F2F2F2 is 2.7:1, which qa.py reports and a room cannot
    # read.  The colour still marks the card -- as the chip's fill, with white
    # on it -- which is where it was legible in the first place.
    head = title_color
    if head is None:
        head = (accent if accent is not None
                and contrast_ratio(accent, on) >= MIN_ACCENT_CONTRAST
                else readable_on(on))

    if number is not None or item.get("title"):
        # Measured, not assumed: a two-word title fits one line and a six-word
        # one does not, and a title that overruns a fixed band is drawn
        # straight through the body under it -- which is what the first cut of
        # this function did on every step of a four-step flow.
        chip_w, chip_gap, text_w, band_h, gap_h = _head_metrics(
            item, inner.w, title_size, number is not None)
        band, inner = inner.cut_top(band_h, gap=gap_h)
        if number is not None:
            chip(doc, page,
                 Box(band.x, band.y, chip_w, int(title_size * 35.28 * 1.35)),
                 fill=accent if accent is not None else GREY_DK,
                 label=str(number), size=title_size, radius=radius)
        if item.get("title"):
            text(doc, page, Box(band.x + chip_w + chip_gap, band.y, text_w,
                                band.h),
                 item["title"], size=title_size, weight=BOLD, color=head,
                 valign="top")

    if item.get("note"):                      # pinned to the floor, not floated
        inner, foot = inner.cut_bottom(int(note_size * 35.28 * 1.5))
        text(doc, page, foot, item["note"], size=note_size,
             color=note_color, weight=BOLD, valign="bottom")

    if item.get("body"):
        text(doc, page, inner, item["body"], size=size,
             color=color if color is not None else readable_on(on),
             valign="top")
    return box


@takes_box("x", "y", "w", "h")
def cards(doc, page, x, y, w, h, items, cols=None, gap=None, fill=GREY_LT,
          accent=None, numbered=False, height="auto", size=None,
          title_size=None, note_size=None, pad=None, radius=None, color=None,
          title_color=None, note_color=GREY, line=None):
    """A row (or grid) of cards, all built the same way.  Returns their boxes.

    Each item is a string, a ``(title, body)`` or ``(title, body, note)``
    tuple, or a dict with any of ``title``, ``body``, ``note``, ``fill`` and
    ``accent``::

        D.cards(doc, page, area, [
            ("Inventory", "Classify every agent against A / B / C.",
             "Owner: ‹name›  ·  By: ‹date›"),
            ("Gate", "A/B/C classification becomes mandatory in design review."),
            ("Accept the trade-off", "Cap capability where all three coincide.",
             "Board decision  ·  today"),
        ], numbered=True, accent=[0x2A7DE1, 0xE8730C, 0x6B3FA0])

    ``height`` is the one that matters.  ``"auto"`` (the default) sizes the
    row so the tallest card's text fills about three quarters of it, never
    less than half the area, and centres the row: no panel of grey under two
    lines of text, and no thin strip floating in a half-empty slide -- the two
    things every vision review of generated cards reported.  ``"equal"`` fills
    the area, which is right when the cards hold a lot.  ``"content"`` gives every
    card the height of the *longest* one's content and centres the row in the
    area -- use it when one card has three lines and another has ten, or the
    short ones ship with half a panel of nothing under the text, which is the
    commonest thing wrong with a finished-looking slide.

    ``note`` sits on the card's floor rather than under the body, so a row of
    cards has its owner/date line on one baseline instead of three.
    """
    if not items:
        raise ValueError("cards(): nothing to draw -- items is empty")
    if height not in ("auto", "equal", "content"):
        raise ValueError("cards(): height= is 'auto', 'equal' (fill the area) "
                         "or 'content' (size to the text), got %r" % (height,))
    items = [_as_item(item) for item in items]
    gap = CARD_GAP if gap is None else gap
    pad = CARD_PAD if pad is None else pad
    note_size = size_for("caption") if note_size is None else note_size
    cols = len(items) if cols is None else max(1, int(cols))
    rows_n = -(-len(items) // cols)
    area = Box(x, y, w, h)

    cell_w = (area.w - gap * (cols - 1)) / cols
    offered = (area.h - gap * (rows_n - 1)) / rows_n
    # From the deck's own scale, so a card's body is a size the rest of the
    # deck uses.  Unless the caller chose, the cards grow up that scale while
    # the tallest still fills under CARD_FILL of its panel and the longest
    # word still fits a line: label-sized words in a half-slide panel are the
    # "hollow cards" every vision review of a generated deck flags first.
    if size is None or title_size is None:
        size, title_size = _grow_card_type(
            items, int(cell_w), offered, size, title_size, note_size,
            pad, numbered)
    if cell_w < 2 * pad + 1000:               # 10 mm of usable card
        raise ValueError(
            "cards(): %d cards across %.0f mm leaves %.0f mm each, which is "
            "narrower than their own padding. Use cols= to wrap them onto "
            "rows, or give them the wider box."
            % (cols, area.w / 100, cell_w / 100))

    needed = max(_card_height(item, int(cell_w), title_size, size,
                              note_size, pad, numbered)
                 for item in items)
    if needed > offered * 1.15:
        # Drawing anyway puts the body outside its own panel, which reads as
        # a card with a coloured hat on and is not something QA measures.
        raise ValueError(
            "cards(): the tallest card needs %.0f mm and the box gives each "
            "row %.0f mm. Give cards() a taller box, fewer words, or a "
            "smaller size=/title_size=." % (needed / 100, offered / 100))
    if height in ("content", "auto"):
        row_h = min(needed, offered)
        if height == "auto":
            row_h = min(offered, max(needed / CARD_TEXT_SHARE,
                                     offered * CARD_MIN_SHARE))
        used = row_h * rows_n + gap * (rows_n - 1)
        area = Box(area.x, int(area.y + (area.h - used) / 2), area.w, int(used))

    tones = [item.get("accent",
                      accent[i % len(accent)]
                      if isinstance(accent, (list, tuple)) else accent)
             for i, item in enumerate(items)]
    fills = [item.get("fill", fill) for item in items]

    # One decision for the whole row, not one per card.  A row where two
    # titles are black because their accent is too pale on the panel and the
    # third is purple because its accent happens to clear 3:1 reads as three
    # kinds of card -- and "peers styled differently" is a defect no
    # measurement reports, because each card on its own is fine.
    heads = [title_color[i % len(title_color)]
             if isinstance(title_color, (list, tuple)) else title_color
             for i in range(len(items))]
    if title_color is None:
        page_bg = bg_color(page)
        readable = all(
            tone is None
            or contrast_ratio(tone, f if f is not None else page_bg)
            >= MIN_ACCENT_CONTRAST
            for tone, f in zip(tones, fills))
        heads = [None if readable
                 else readable_on(f if f is not None else page_bg)
                 for f in fills]

    out = []
    for r, band in enumerate(area.rows(rows_n, gap=gap)):
        for c, cell in enumerate(band.cols(cols, gap=gap)):
            i = r * cols + c
            if i >= len(items):
                break
            out.append(_draw_card(
                doc, page, cell, items[i], fills[i], tones[i],
                i + 1 if numbered else None, title_size, size, note_size, pad,
                radius, color, heads[i], note_color, line))
    return out


@takes_box("x", "y", "w", "h")
def steps(doc, page, x, y, w, h, items, captions=None, numbered=True,
          accent=None, fill=GREY_LT, gap=None, size=None, title_size=None,
          caption_size=None, arrow_color=None, pad=None, radius=None):
    """A left-to-right flow: equal tiles, real arrows between them, captions
    under them.  Returns the tile boxes.

    ``steps`` takes the same items as :func:`cards`; ``captions`` is an
    optional label per step, drawn *below* the row and colour-matched to
    ``accent``, which is how a flow says which property each step belongs to
    without a legend::

        D.steps(doc, page, area,
                [("A request arrives", "Nothing is decided yet"),
                 ("Agent reads hidden text", "…"),
                 ("Agent reads your last 50 emails", "…"),
                 ("The answer goes out", "…")],
                captions=["outside", "untrusted content", "private data",
                          "external comms"],
                accent=[0x4A4A4A, 0xE8730C, 0x2A7DE1, 0x6B3FA0])

    The arrows are drawn in the gap between tiles and sized to it -- a flow
    whose arrows are 5 mm stubs floating in a 20 mm gap is the hand-rolled
    version of this.
    """
    if not items:
        raise ValueError("steps(): nothing to draw -- items is empty")
    gap = CARD_GAP * 2 if gap is None else gap    # room for the arrowhead
    caption_size = size_for("caption") if caption_size is None else caption_size
    area = Box(x, y, w, h)
    if captions:
        if len(captions) != len(items):
            raise ValueError("steps(): %d step(s) but %d caption(s) -- one "
                             "each, or none" % (len(items), len(captions)))
        area, foot = area.cut_bottom(int(caption_size * 35.28 * 2.2),
                                     gap=int(caption_size * 35.28 * 0.5))

    # Content height, not the full band: a flow's tiles hold a line or two
    # each, and four tiles stretched to fill a half-slide are four panels of
    # grey with the words along their top edge.
    boxes = cards(doc, page, area, items, gap=gap, fill=fill, accent=accent,
                  numbered=numbered, size=size, title_size=title_size,
                  pad=pad, radius=radius, height="content")

    for left, right in zip(boxes, boxes[1:]):
        span = right.x - left.x2
        head = min(span * 0.8, area.h * 0.22)
        arrow(doc, page,
              Box(int(left.x2 + (span - head) / 2), int(left.cy - head / 2),
                  int(head), int(head)),
              color=arrow_color if arrow_color is not None else GREY)

    if captions:
        # Under the tiles, wherever the tiles ended up -- ``height="content"``
        # centres them in the band, and a caption left at the foot of the band
        # is a label floating half a slide below the thing it labels.
        foot = Box(foot.x, int(boxes[0].y2 + caption_size * 35.28 * 0.6),
                   foot.w, foot.h)
        for i, (box, caption) in enumerate(zip(boxes, captions)):
            tone = accent[i % len(accent)] if isinstance(accent, (list, tuple)) \
                else accent
            text(doc, page, Box(box.x, foot.y, box.w, foot.h), caption,
                 size=caption_size, weight=BOLD, align="center", valign="top",
                 color=tone if tone is not None else GREY_DK)
    return boxes


@takes_box("x", "y", "w", "h")
def kpis(doc, page, x, y, w, h, items, color=None, caption_color=GREY_DK,
         gap=None, max_size=54.0, caption_size=None, align="center"):
    """A row of big numbers with captions under them.  Returns their boxes.

    ``items`` are ``(value, caption)`` pairs::

        D.kpis(doc, page, area, [("~47,000", "downloads of the affected
                                  package"), ("3 hours", "it sat on PyPI"),
                                 ("dozens", "frameworks depend on it")])

    Every value is set at **one** size -- the largest at which the longest of
    them still fits its column -- because three numbers at three sizes read as
    three levels of importance rather than three facts.
    """
    if not items:
        raise ValueError("kpis(): nothing to draw -- items is empty")
    items = [(item, "") if isinstance(item, str)
             else (tuple(item) + ("",))[:2] for item in items]
    gap = CARD_GAP * 2 if gap is None else gap
    caption_size = size_for("caption") if caption_size is None else caption_size
    area = Box(x, y, w, h)
    columns = area.cols(len(items), gap=gap)

    value_size = min(max_size, min(
        fit_size(str(value), columns[0].w, BOLD, max_size=max_size)
        for value, _ in items))
    band = int(value_size * 35.28 * 1.35)

    # Centred in the area rather than parked at its top: a number and its
    # caption need about a third of the height they are usually handed, and the
    # other two thirds are the bare band that reads as an unfinished slide.
    caption_h = max((_text_height([caption], columns[0].w, caption_size)
                     for _, caption in items if caption), default=0)
    gap_h = int(value_size * 35.28 * 0.2)
    block = min(area.h, band + (gap_h + caption_h if caption_h else 0))
    top = int(area.y + (area.h - block) / 2)
    columns = [Box(c.x, top, c.w, block) for c in columns]

    out = []
    for column, (value, caption) in zip(columns, items):
        head, rest = column.cut_top(band, gap=gap_h)
        text(doc, page, head, str(value), size=value_size, weight=BOLD,
             align=align, valign="bottom",
             color=color if color is not None else readable_on(bg_color(page)))
        if caption:
            text(doc, page, rest, str(caption), size=caption_size,
                 align=align, valign="top", color=caption_color)
        out.append(column)
    return out


def balance_lines(s, width, size, weight=NORM, max_lines=None):
    """Wrap ``s`` to ``width`` with the lines evened up.  ``[line, ...]``.

    A greedy wrap fills each line to the brim and leaves the remainder on the
    last one, which is how a statement slide ends on a line reading ``an``.
    This wraps to the *narrowest* width that still needs the same number of
    lines, so the slack is shared out instead of all landing at the end.
    """
    lines = wrap_words(s, width, size, weight)
    if len(lines) < 2:
        return lines
    if max_lines and len(lines) > max_lines:
        return lines                          # caller decides: smaller, or wider
    target, low, high = len(lines), width // 4, width
    while high - low > width // 100 + 1:
        mid = (low + high) // 2
        if len(wrap_words(s, mid, size, weight)) <= target:
            high = mid
        else:
            low = mid
    return wrap_words(s, high, size, weight)


@takes_box("x", "y", "w", "h")
def statement(doc, page, x, y, w, h, *paragraphs, max_lines=3, size=None,
              max_size=44.0, min_size=16.0, weight=BOLD, color=None,
              align="center", valign="center", font=DEFAULT_FONT):
    """One sentence, as large as it fits, broken where it reads.  The shape.

    For the slide that is a single line of type::

        D.statement(doc, page, area,
                    "We are deploying software that reliably follows "
                    "instructions from anyone who can get text in front of it.")

    It picks the largest size at which the text fits ``max_lines`` and the box,
    then evens the lines up (see :func:`balance_lines`) and writes them as
    separate paragraphs so LibreOffice cannot re-break them somewhere else.

    Pass several strings for several paragraphs.  Do **not** put ``\\n`` in one
    -- that is hand-breaking, and hand-breaking is what leaves a slide ending
    on a line that reads ``an``.
    """
    if not paragraphs:
        raise ValueError("statement(): nothing to write")
    for part in paragraphs:
        if "\n" in str(part):
            raise ValueError(
                "statement(): the text contains a newline. This call chooses "
                "the size and the line breaks together, and a hand-broken "
                "string fights it -- a centred statement broken by hand is "
                "how a slide ends on a line reading 'an'.\n"
                "  Pass the sentence whole; pass SEVERAL strings for several "
                "paragraphs:\n"
                "      D.statement(doc, page, area, first, second)")

    area = Box(x, y, w, h)
    inside = area.w * WRAP_SLACK
    chosen, lines = min_size, None
    trial = max_size if size is None else float(size)
    while trial >= min_size:
        wrapped = []
        for part in paragraphs:
            wrapped += balance_lines(str(part), inside, trial, weight)
        if (len(wrapped) <= max_lines * len(paragraphs)
                and len(wrapped) * trial * 35.28 * 1.22 <= area.h):
            chosen, lines = trial, wrapped
            break
        if size is not None:                  # a size was asked for: keep it
            chosen, lines = trial, wrapped
            break
        trial -= 1.0
    if lines is None:                         # nothing fits: smallest, wrapped
        lines = []
        for part in paragraphs:
            lines += balance_lines(str(part), inside, min_size, weight)
        chosen = min_size

    # Then check it against the layout engine rather than against the estimate.
    # The estimate is good to a few percent, and a few percent is exactly the
    # margin a balanced line sits on: the first version of this call put three
    # measured lines on the slide and LibreOffice re-broke the last one, which
    # is the orphan the whole function exists to avoid.  Two round trips per
    # attempt, on the one slide in a deck that is a single sentence.
    while True:
        shape = text(doc, page, area, *lines, size=chosen, weight=weight,
                     align=align, valign=valign, color=color, font=font)
        try:
            widest = max(text_width(doc, page, part, shape) for part in lines)
        except Exception:                     # no layout available: trust the sums
            return shape
        if widest <= area.w or chosen <= min_size:
            return shape
        # The breaks are balanced and only the type is a shade too big, so the
        # type gives way and the breaks stay.  A point either way is invisible;
        # a re-broken line is not.
        page.remove(shape)
        chosen = max(min_size, chosen * min(0.98, area.w / widest))


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


#: Air between a lobe's edge and the label that names it, as a share of r.
VENN_LABEL_GAP = 0.08


def _venn_labels(cx, cy, r, n, labels, label_size=None):
    """Where each outer label goes: ``[(x, y, w, h, size, lines), ...]``.

    **Anchored to the lobe, not pushed along a ray.**  Each label is set
    against the edge of the circle it names -- above it, below it, or beside
    it, whichever the lobe's direction from the figure's centre is mostly --
    and offset by half its own ink so that it touches rather than overlaps.

    The version before this solved for the distance along the ray at which a
    label centred there stopped overlapping the circle.  That is geometrically
    correct and reads wrong: a wide label on a diagonal has to travel a long
    way before its near corner clears, so the two lower labels of a three-lobe
    Venn ended up far out to the lower left and lower right, close to nothing,
    with the figure marooned in the middle.  Anchoring instead puts
    ``UNTRUSTED CONTENT`` against the side of the circle it belongs to, and
    turns the wasted diagonal into width -- which a 16:9 slide has and height
    it does not.

    Pure geometry, so ``venn_bounds`` can ask the same question without drawing
    anything, which is what lets ``venn(box=...)`` size itself to fit.
    """
    out = []
    box_w, box_h = r * 1.7, r * 0.62
    for i, label in enumerate(labels[:n]):
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
        gap = r * VENN_LABEL_GAP
        if abs(uy) >= abs(ux):                # over or under its own lobe
            sx = ex
            sy = ey + (1 if uy > 0 else -1) * (r + gap + ink_h / 2)
        else:                                 # beside it, on its own centre line
            sx = ex + (1 if ux > 0 else -1) * (r + gap + ink_w / 2)
            sy = ey
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


def _centre_fit(parts, fixed, box):
    """``(size, lines)`` for a Venn's centre label inside a square of ``box``.

    ``fixed`` keeps the caller's own line breaks and sizes to the longest of
    them; otherwise the string is wrapped on spaces and sized to fit both ways.
    """
    if not parts:
        return CENTRE_MAX_PT, []
    if not fixed:
        return fit_lines(" ".join(parts), box, box, BOLD,
                         max_size=CENTRE_MAX_PT)
    widest = min(fit_size(line, box, BOLD, max_size=CENTRE_MAX_PT)
                 for line in parts)
    tall = box / max(len(parts) * 35.28 * 1.22, 1e-6)
    return min(widest, tall, CENTRE_MAX_PT), list(parts)


def venn(doc, page, cx=None, cy=None, r=None, colors=(), labels=None,
         transparence=76, line_width=45, label_size=None, label_color=None,
         centre=None, centre_color=None, centre_text=None, centre_size=None,
         box=None):
    """Two or three overlapping circles.  Returns the centres, so you can
    label the regions yourself.

    ``colors`` is one colour per circle and decides how many are drawn --
    pass a prefix of it to build the diagram up over successive slides.
    ``labels``, if given, are set against the edge of each circle.
    ``centre`` fills the all-way overlap with that colour, and ``centre_text``
    names it: a string wraps, a **list is one line per entry**.

    **Give it the area, not a radius** -- a ``Box``, positionally or as
    ``box=``, typically the one ``canvas`` returned::

        D.venn(doc, page, area, D.series(3),
               labels=["PRIVATE DATA", "UNTRUSTED CONTENT", "EXTERNAL COMMS"],
               centre=D.color("alert"), centre_text="EXFILTRATION")

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
    if n not in (2, 3):
        raise ValueError(
            "venn(): %d colours -- a Venn here has two or three lobes. For "
            "more sets draw circles with ellipse(), or use cards()." % n)
    if box is not None:
        cx, cy, r = venn_fit(box, n, labels, label_size)
    elif "radius" not in _venn_hinted:
        _venn_hinted.add("radius")
        print(_VENN_RADIUS_HELP, file=sys.stderr)
    centres = venn_geometry(cx, cy, r, n)
    # Held for the whole call, not just the lobes: the centre disc is an
    # ellipse too, so a two-lobe Venn with a filled centre is three ellipses on
    # the slide and used to trigger this function's own "if this is a Venn, use
    # D.venn()" nudge -- from inside D.venn().
    _drawing_venn = True
    try:
        return _venn_body(doc, page, cx, cy, r, n, centres, colors, labels,
                          transparence, line_width, label_size, label_color,
                          centre, centre_color, centre_text, centre_size)
    finally:
        _drawing_venn = False


def _venn_body(doc, page, cx, cy, r, n, centres, colors, labels, transparence,
               line_width, label_size, label_color, centre, centre_color,
               centre_text, centre_size):
    """The drawing half of :func:`venn`, with the hint suppressed around it."""
    for i, color in enumerate(colors):
        ex, ey = centres[i]
        ellipse(doc, page, ex, ey, r, fill=color,
                transparence=transparence, line=color,
                line_width=line_width)
    if labels:
        for i, (lx, ly, lw, lh, size, lines) in enumerate(
                _venn_labels(cx, cy, r, n, labels, label_size)):
            text(doc, page, lx, ly, lw, lh, *lines, size=size, weight=BOLD,
                 align="center", valign="center",
                 color=(label_color or colors[i]))
    if centre is not None:
        # A string wraps where it fits; a list is the lines, one each, because
        # where a two-word label breaks is a wording decision and the caller is
        # the one holding it.
        fixed = not isinstance(centre_text, str) and bool(centre_text)
        parts = ([centre_text] if isinstance(centre_text, str)
                 else [str(p) for p in (centre_text or [])])
        joined = " ".join(parts)

        # 0.41 r is the inradius of the three-way overlap, so the disc sits
        # exactly inside it.  A long label at that size shrinks to nothing --
        # better to let the disc grow a little into the pair regions than to
        # set "DATA EXFILTRATION" at 7pt.  Capped where it would start hiding
        # the two-way overlaps that the figure exists to show.
        cr = r * 0.41
        if joined and not centre_size:
            while cr < r * 0.58 and _centre_fit(
                    parts, fixed, cr * 1.38)[0] < MIN_READABLE_PT + 2.0:
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
            size, lines = _centre_fit(parts, fixed, box_w)
            if centre_size:
                size = centre_size
                lines = parts if fixed else wrap_words(joined, box_w, size, BOLD)
            elif size < MIN_READABLE_PT and "centre" not in _venn_hinted:
                # The disc has already grown as far as it can without eating
                # the pair regions the figure exists to show, so the only thing
                # left to change is the wording -- which is the caller's, not
                # ours.  "DATA EXFILTRATION" inside a 35 mm figure sizes to
                # about 7pt however the disc is drawn.
                _venn_hinted.add("centre")
                print(
                    "[draw] the Venn centre label %r sizes to %.1fpt inside the "
                    "centre disc, under the %.0fpt floor -- readable on your "
                    "monitor, not from a seat.\n"
                    "  The disc cannot grow further without covering the "
                    "two-way overlaps the figure is for, so shorten the label:\n"
                    "      centre_text=\"EXFILTRATION\"      # one short word\n"
                    "      centre_text=[\"DATA\", \"LEAK\"]     # a list is one "
                    "line per entry\n"
                    "  Or name the centre in a caption under the figure and "
                    "leave the disc plain."
                    % (joined, size, MIN_READABLE_PT), file=sys.stderr)
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
    """A shape and, recursively, every shape inside it if it is a group.

    Template slides put half their text inside groups, where ``shapes(page)``
    cannot see it.  ``office.template.retext`` and ``qa.py`` walk groups for
    you; use this when you need the shapes themselves.  The same walk as
    ``office.template.walk_shapes``: groups only, never into a table's cells.
    """
    yield shape
    if _is_group(shape):
        for _path, inner, _parent in _template_walk(shape):
            yield inner


def walk_page(page):
    """Every shape on a slide, groups included, depth first."""
    for _address, shape in walk_shapes(page):
        yield shape


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
    if not keep and not remove:
        raise TypeError(
            "canvas(): say what to keep.  With neither keep= nor remove= this "
            "would delete nothing and hand you an area guessed from the page, "
            "so you would draw ON TOP of the template's own text -- which "
            "renders as a slide in two languages and reports as 'still the "
            "template's own wording'.\n"
            "  Run probe.py on the design and keep its chrome:\n"
            "      area = D.canvas(page, keep=(0, 1, 2, 3))   # title, topic, source\n"
            "  The catalog entry for each design lists the tuple that works.\n"
            "  For the whole slide with nothing removed, that is "
            "D.page_box(page).")
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
        # Clamped to the slide: a generated template (PptxGenJS output, for
        # one) can carry body boxes a few millimetres wider than the page, and
        # a free area measured off them puts everything drawn in it off the
        # right edge.  The slide is the outer limit whatever the shapes say.
        x0 = max(0, min(s[0] for s in spans))
        y0 = max(0, min(s[1] for s in spans))
        x1, y1 = max(s[2] for s in spans), max(s[3] for s in spans)
        if x1 > page.Width:                   # overhangs: mirror the left margin
            x1 = page.Width - x0
        if y1 > page.Height:
            y1 = page.Height - y0
        return Box(x0, y0, x1 - x0, y1 - y0)

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


def __getattr__(name):
    """Name the module a missing ``D.<something>`` actually lives in.

    ``office.draw`` is imported as ``D`` and is the module a build script has
    in its hands, so it is where every slide-shaped verb gets tried first --
    ``D.find_shapes``, ``D.shrink_to_fit``, ``D.retext``.  Those are real
    functions in ``office.template``, and Python's own "module 'office.draw'
    has no attribute 'find_shapes'" gives no hint that they exist at all: one
    run spent four builds on two of them, guessing at spellings.

    Only called on a miss, so it costs nothing until something is already
    wrong.
    """
    if name.startswith("__"):
        raise AttributeError(name)

    from . import impress, template            # lazy: only on a miss
    if not name.startswith("_"):
        found = [m for m in (template, impress) if hasattr(m, name)]
        # Where it is *defined*, not where it is also visible: template imports
        # add_slide from impress, so a plain first-hit wins sends the caller to
        # the wrong module -- which still works, and teaches the wrong home for
        # everything else they look for next.
        home = next(
            (m for m in found
             if getattr(getattr(m, name), "__module__", None) == m.__name__),
            found[0] if found else None,
        )
        if home is not None:
            raise AttributeError(
                "office.draw has no %r -- it is in %s.\n  from %s import %s"
                % (name, home.__name__, home.__name__, name))

    close = [n for n in __all__ if name.lower() in n.lower()
             or n.lower() in name.lower()]
    raise AttributeError(
        "office.draw has no %r.%s\n"
        "  draw is shapes and text on a slide; office.template is the one that"
        " reads and rewrites\n  what a template already put there "
        "(retext, find_shapes, shrink_to_fit, strip_annotations)."
        % (name, ("  Closest here: %s" % ", ".join(sorted(close)[:6]))
           if close else ""))


__all__ = [
    "Point", "Size", "BLACK", "WHITE", "GREY", "GREY_DK", "GREY_BD",
    "GREY_MD", "GREY_LT", "BOLD", "NORM", "THIN", "DEFAULT_FONT",
    "rect", "ellipse", "line", "hline", "vline", "polygon", "arrow",
    "chevron", "text", "chip", "bullets", "table",
    "cards", "steps", "kpis", "statement", "balance_lines",
    "venn", "venn_geometry",
    "venn_fit", "venn_bounds",
    "pair_spots", "restyle", "style_fill", "walk", "walk_page", "move",
    "resize", "canvas", "est_lines", "set_text", "Box", "Split", "page_box",
    "advance_em", "text_width_est", "fit_size", "fit_lines", "wrap_words",
    "contrast_ratio", "bg_color", "readable_on",
    "palette", "palette_roles", "color", "series", "type_scale", "size_for",
    "MIN_READABLE_PT",
]
