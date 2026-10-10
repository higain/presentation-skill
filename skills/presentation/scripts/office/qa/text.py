"""Text on a slide: placeholders, markdown, emoji, type below the template's floor."""

from __future__ import annotations

import re

from ..draw import min_readable
from ..impress import placeholder
from ..template import walk_shapes

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
# Two runs measured against a corporate template say otherwise: a **15-character
# design label** -- a documented retext key of the canvas design, of the "Topic
# (1 of 2)" kind -- survived on 13 of 15 slides in one deck and 9 of 10 in the
# other, in the template's language, on a finished deck in another, and both
# reports said clean.  Nothing else in either deck came within nine characters
# of the old threshold.
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


#: Colour-font emoji on a corporate slide.  The pictograph blocks, anything
#: carrying variation-selector-16 (which is what forces emoji presentation on a
#: character that would otherwise be a glyph), and the handful of older
#: codepoints that render in colour without it.
#:
#: Deliberately NOT the dingbats as a block: ✓ and ✗ are ordinary monochrome
#: glyphs that a table has every right to use, and they set in the deck's own
#: font.  ✅ and ❌ are pictures, they arrive at whatever size and colour the
#: emoji font feels like, and they render differently in LibreOffice and
#: PowerPoint -- one run put ✅ ⚠️ ❌ in a status column and coloured the word
#: beside them red for "safe".
EMOJI = re.compile(
    "[\U0001F000-\U0001FAFF]"            # pictographs, faces, transport, flags
    "|️"                            # variation selector-16: "draw as emoji"
    "|[✅❌❎⭐⭕"   # heavy tick, cross, star, circle
    "⛔⚠❗❓❔"     # no-entry, warning, bang, question
    "‼⁉⬆⬇❤]"    # !!, !?, arrows, heart
)


def _text_sizes(shape):
    """``{size: sample text}`` for every run in the shape, in points.

    Same enumeration as ``_text_colors`` and for the same reason: the size that
    matters is the run's, not the shape's.
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
                sample, size = run.getString().strip(), float(run.CharHeight)
            except Exception:
                continue
            if sample and size > 0:
                out.setdefault(round(size, 1), sample)
    return out


def tiny_type_issues(page, number, baseline=frozenset()):
    """Text set smaller than the deck's own smallest role.

    The floor is ``draw.min_readable()`` -- the smallest size **this deck's
    template** itself types at, so the limit is the design's own rather than a
    number invented here, and it differs per template.  Anything the *template*
    set that small is baselined out: a pattern collection's think-cell stubs
    and corner marks are 3 and 5pt by design.

    Short strings are ignored. A lone digit or a units mark at 6pt is a
    diagram's furniture; a sentence at 6pt is a footnote nobody in the room can
    read, and both of the runs this was written for produced one.
    """
    # Clustered by size: one caption set too small is usually a row of them,
    # and a line per cell reads as twelve mistakes instead of one decision.
    floor = min_readable()
    seen = {}
    for address, shape in walk_shapes(page):
        for size, sample in _text_sizes(shape).items():
            if size >= floor or len(sample) < 12:
                continue
            if _norm(sample) in baseline:
                continue
            seen.setdefault(size, [address, sample, 0])[2] += 1
    out = []
    for size, (address, sample, count) in sorted(seen.items()):
        more = (f" (and {count - 1} more run(s) at the same size)"
                if count > 1 else "")
        out.append(
            f"slide {number} shape [{address}]: {size:g}pt text -- under "
            f"the {floor:g}pt floor this template types at, "
            f"legible on a monitor and not from a seat: {sample[:40]!r}{more}. "
            f"Use the scale: D.text(..., role='footnote')"
        )
    return out


def emoji_issues(page, number, baseline=frozenset()):
    """Colour-font emoji on a slide.  See :data:`EMOJI`."""
    seen = {}
    for address, text in _page_texts(page):
        if text in baseline:
            continue
        hit = EMOJI.search(text)
        if hit:
            seen.setdefault(hit.group(0), [address, text, 0])[2] += 1
    out = []
    for glyph, (address, text, count) in seen.items():
        more = (f" (and {count - 1} more shape(s) with it)" if count > 1 else "")
        out.append(
            f"slide {number} shape [{address}]: emoji {glyph!r} in "
            f"{text[:40]!r}{more} -- a colour-font picture on a corporate "
            f"slide, sized and coloured by the emoji font rather than by the "
            f"template, and drawn differently by LibreOffice and PowerPoint. "
            f"Use a D.chip, a coloured D.rect, or the monochrome ✓ / ✗"
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


#: A slide with at least this many text shapes is a content slide, not a
#: divider whose headline lives in a display-type shape of its own.
TITLED_SLIDE_TEXTS = 3


def missing_title_issues(page, number):
    """A content slide whose title placeholder is empty.

    An empty placeholder renders as nothing, so the slide shows no headline --
    and nothing else here sees it: the ghost deck then reads the nearest small
    label as the title.  ``retext`` cannot fill an empty placeholder, which is
    how it happens; ``set_title`` can.
    """
    title = placeholder(page, "Title")
    if title is None:
        return []
    try:
        if title.getString().strip():
            return []
    except Exception:
        return []
    if sum(1 for _ in _page_texts(page)) < TITLED_SLIDE_TEXTS:
        return []
    return [f"slide {number}: the title placeholder is empty, so the slide "
            f"shows no headline. Write it with "
            f"`office.template.set_title(page, '...')` -- `retext` cannot "
            f"reach an empty shape"]
