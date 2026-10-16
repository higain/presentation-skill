"""Colour and stacking: text contrast against what is behind it, and occlusion."""

from __future__ import annotations

from ..impress import shapes
from ..draw import contrast_ratio
from ..template import walk_shapes
from ..textenum import enumerate_text
from .text import _norm

# Below this WCAG ratio, text on its backdrop is unreadable at any size a slide
# uses.  4.5 is the AA threshold for body text; 3.0 is where a headline stops
# being legible at all, and it is the number worth failing a build over.
MIN_CONTRAST = 3.0

#: What an unset slide background renders as.  See ``_page_bg``.
WHITE_BG = 0xFFFFFF

# Below this the text is not merely hard to read, it is not visible: black on
# navy, white on lime.  Always a mistake, never a style.
INVISIBLE = 1.8

# How much of a shape another one must cover before it counts as hiding it.
# A caption sitting on a panel overlaps it; a fill that swallows four fifths of
# a panel replaced it, which is what a mis-ordered cut_* produces.
OCCLUSION = 0.75


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


def _fill_layer(shape):
    """``(colour, opacity)`` of a shape's solid fill, or ``None``.

    Opacity is ``1 - transparence``.  A tinted panel at 85 % transparence is
    still a panel: white text on it reads as white on near-white, and treating
    the fill as absent compared the text with the slide instead.
    """
    try:
        if not _is_solid(shape.FillStyle):
            return None
        alpha = 1 - int(shape.FillTransparence) / 100
        if alpha <= 0.02:
            return None
        return int(shape.FillColor) & 0xFFFFFF, alpha
    except Exception:
        return None


def _blend(top, alpha, under):
    """``top`` at ``alpha`` over ``under``, both ints."""
    out = 0
    for shift in (16, 8, 0):
        t, u = (top >> shift) & 255, (under >> shift) & 255
        out |= int(round(t * alpha + u * (1 - alpha))) << shift
    return out


def _page_bg(page):
    """The colour behind everything on this page.

    Falls back to **white** when neither the page nor its master sets one,
    because that is what an unset background renders as -- in LibreOffice, in
    PowerPoint, and on the projector.  Returning ``None`` instead meant text
    sitting on the bare slide had no backdrop to be checked against, so the
    contrast check covered exactly the slides with a coloured panel and skipped
    every ordinary white one: a label in the brand's own lime is 2.15:1 on
    white, and it reported clean.
    """
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
    return WHITE_BG


def _text_colors(shape):
    """``{colour: sample text}`` for every run in the shape that carries one.

    Reading ``shape.CharColor`` is not enough and usually returns -1
    ("automatic"): once a shape holds text the runs carry their own hard
    formatting, which is where ``draw.restyle`` and every template put it.
    """
    out = {}
    try:
        paragraphs = list(enumerate_text(shape))
    except Exception:
        return out
    for para in paragraphs:
        if not hasattr(para, "createEnumeration"):
            continue
        for run in enumerate_text(para):
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
    boxes = [(_geometry(s), _fill_layer(s)) for s in tops]

    for address, shape in walk_shapes(page):
        colors = _text_colors(shape)
        geometry = _geometry(shape)
        if not colors or geometry is None:
            continue

        # The colour under the text's centre: the slide, then every filled
        # shape below the text in z-order, translucent ones blended in.
        cx = geometry[0] + geometry[2] // 2
        cy = geometry[1] + geometry[3] // 2
        head = str(address).split(".")[0]
        index = int(head) if head.isdigit() else len(tops)
        backdrop = slide_bg
        for other, layer in boxes[:min(index, len(boxes))]:
            if layer is not None and other and _contains(other, cx, cy):
                backdrop = _blend(layer[0], layer[1], backdrop) \
                    if backdrop is not None else layer[0]
        own = _fill_layer(shape)
        if own is not None:
            backdrop = _blend(own[0], own[1], backdrop) \
                if backdrop is not None else own[0]
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


def _is_picture(shape):
    try:
        return shape.ShapeType.endswith("GraphicObjectShape")
    except Exception:
        return False


def buried_text_issues(page, number):
    """Text drawn *under* a picture that hides it.

    Overlap, contrast and the vision-agent all missed this one.  A divider's
    title placeholder sat beneath a full-slide image; the build wrote three
    section names into it, every check read the words straight out of the
    file, and the reviewer who said "no headline" was overruled with "the
    probe shows the title is there".  It was there, and nobody could see it.
    """
    out = []
    tops = list(shapes(page))
    for below, shape in enumerate(tops):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        geo = _geometry(shape)
        if not text or geo is None:
            continue
        for above in range(below + 1, len(tops)):
            pic = _geometry(tops[above])
            if (pic is not None and _is_picture(tops[above])
                    and _covers(pic, geo, OCCLUSION)):
                out.append(
                    f"slide {number} shape [{below}]: {text[:40]!r} is under "
                    f"picture [{above}] and does not show -- write the text "
                    f"into a shape above the picture (on a divider, the "
                    f"design's own label), never into the one beneath it")
                break
    return out
