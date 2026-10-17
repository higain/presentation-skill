"""Does the deck say what the source says -- no less, and nothing invented?

Two directions, both on normalised text (quotes, dashes, spacing and case
evened out):

* **source -> deck**: every piece of the source that is content -- a slide
  heading, text in quotation marks, a table cell, a plain list item, a
  speaker-note line -- has to be on the slides or in the notes.  Present
  means: the normalised text occurs, or at least ``WORDS_PRESENT`` of its
  words do *and* every one of its numbers does.  A shortened quote and a
  dropped "> 90 %" both fail that.
* **deck -> source**: every sentence of six words or more on a slide has to
  come from the source.  Fewer than ``WORDS_SOURCED`` of its content words
  anywhere in the source is invented text.  The template's own wording is
  not the build's and is skipped.

What counts as content in a Markdown source: headings of the slide sections
(``## 3 · Title`` or ``### Folie 3: Title``), quoted text („…“, "…", »…«),
table body cells, plain list items, and ``>`` lines (speaker notes, compared
with the notes).  A list item that opens with a bold label ending in a colon
-- ``**Struktur:** …`` -- is an instruction to the builder, not slide text,
and only the quotes inside it count.

A field run shortened a quote, dropped a key figure and wrote a paragraph of
its own to fill an empty band; every other check passed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

#: Share of a source unit's words that must be in the deck for it to count
#: as present when its exact text is not.
WORDS_PRESENT = 0.85
#: Share of a slide sentence's word pairs that must occur in the source.
#: Pairs, not words: invented text is made of the source's own vocabulary
#: ("Konkrete Leitplanken für den Einsatz von KI-Tools" -- every word is in
#: the source), and only the order of the words gives it away.
PAIRS_SOURCED = 0.5
#: Slide texts shorter than this are labels, not statements.
MIN_SENTENCE_WORDS = 4

_QUOTES = re.compile(r"[„“”\"»«‚‘’']")
_DASHES = re.compile(r"[‐‑‒–—−]")
_WORD = re.compile(r"[0-9a-zäöüß]+(?:[.,][0-9]+)*", re.I)
_NUMBER = re.compile(r"\d+(?:[.,]\d+)*")
_SLIDE_HEADING = re.compile(
    r"^#{1,4}\s*(?:Folie\s+)?(\d+)\s*(?:[·:.\-–—)]\s*)?(.*)$", re.I)
_QUOTED = re.compile(r"„([^“”]{3,}?)[“”]|»([^«]{3,}?)«|\"([^\"]{3,}?)\"")
_INSTRUCTION = re.compile(r"^\*\*[^*]{1,40}:\*\*")
#: A list item that describes a shape -- a position in brackets, a hex
#: colour -- and not the words on it: "[x 13 %, y 27 %] Kreis, Füllung #FFFFFF".
_SHAPE_LINE = re.compile(r"^\[x\s*\d|#[0-9A-Fa-f]{6}\b|\(Kopfzeile")
_BOLD_LABEL = re.compile(r"^\*\*([^*]{1,60})\*\*\s*[–—-]\s*")
_STOP = {"und", "oder", "der", "die", "das", "den", "dem", "des", "ein",
         "eine", "einen", "einem", "einer", "mit", "von", "für", "auf", "aus",
         "ist", "sind", "wird", "werden", "nicht", "auch", "wie", "dass",
         "the", "and", "for", "with", "that", "this", "are", "from", "nur"}


def norm(text: str) -> str:
    """Lower case, one kind of quote and dash, single spaces, no bullets."""
    text = text.replace("­", "").replace("•", " ").replace("·", " ")
    text = _QUOTES.sub('"', _DASHES.sub("-", text))
    return " ".join(text.lower().split())


def words(text: str) -> list:
    return _WORD.findall(norm(text))


@dataclass
class Unit:
    """One piece of source content and where it came from."""
    text: str
    section: str          # the slide heading it sits under
    notes: bool = False   # compare with the speaker notes


def _strip_md(text: str) -> str:
    text = re.sub(r"\*\*|__|`", "", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)
    return text.strip()


def source_units(markdown: str) -> list:
    """The content units of a Markdown source, in order."""
    lines = markdown.splitlines()
    starts = [i for i, line in enumerate(lines) if _SLIDE_HEADING.match(line)]
    # A specification numbers its own chapters too ("## 1.1 Grundraster");
    # when slides are headed "Folie N" / "Slide N", only those are slides,
    # and everything before the first one is preamble.
    named = [i for i in starts
             if re.match(r"^#{1,4}\s*(?:Folie|Slide)\s+\d+", lines[i], re.I)]
    if named:
        starts = named
    if not starts:
        sections = [("", lines)]
    else:
        sections = []
        for k, start in enumerate(starts):
            end = starts[k + 1] if k + 1 < len(starts) else len(lines)
            sections.append((lines[start], lines[start + 1:end]))

    units = []
    for heading, body in sections:
        name = _strip_md(heading.lstrip("#").strip())
        match = _SLIDE_HEADING.match(heading) if heading else None
        if match and match.group(2).strip():
            units.append(Unit(_strip_md(match.group(2)), name))
        # Join wrapped list items and paragraphs: a continuation line is
        # indented and does not start a new item.
        items, current = [], None
        for line in body:
            stripped = line.strip()
            starts_item = bool(re.match(r"^(?:[-*+]|\d+\.)\s", stripped))
            if (not stripped or starts_item or stripped.startswith(("|", ">", "#"))
                    or current is None):
                if current is not None:
                    items.append(current)
                current = stripped if stripped else None
            else:
                current += " " + stripped
        if current is not None:
            items.append(current)

        # In a section that carries its slide text in quotes -- a
        # specification's "Textfeld – Text: „…“" lines -- an item without
        # quotes describes a shape ("Kreis/Oval, Füllung Weiß") and is not
        # slide text.
        # Quotes in a plain paragraph are the builder's layout notes
        # ("„Ebene n“ klein fett", "Icon „FiEdit3“") when the section's list
        # items already carry the slide text in quotes.
        quoted_lists = any(re.match(r"^(?:[-*+]|\d+\.)\s", i)
                           and _QUOTED.search(i) for i in items)
        header, skip_cols = None, set()
        for k, item in enumerate(items):
            if item.startswith(">"):
                text = _strip_md(item.lstrip("> ").strip())
                if len(words(text)) >= 3:
                    units.append(Unit(text, name, notes=True))
                continue
            if item.startswith("|"):
                cells = [c.strip() for c in item.strip("|").split("|")]
                if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    continue
                following = items[k + 1] if k + 1 < len(items) else ""
                if re.fullmatch(r"\|?(\s*:?-{2,}:?\s*\|)+\s*:?-*:?\s*\|?",
                                following.strip()):
                    # The header row names the columns of a description
                    # table; an icon column names pictures, not words.
                    skip_cols = {i for i, c in enumerate(cells)
                                 if norm(_strip_md(c)) in ("icon", "symbol",
                                                            "bild", "piktogramm")}
                    continue
                for i, cell in enumerate(cells):
                    if i in skip_cols or _SHAPE_LINE.search(cell.strip()):
                        continue
                    cell = _strip_md(cell)
                    if cell and (len(words(cell)) >= 2 or _NUMBER.search(cell)):
                        units.append(Unit(cell, name))
                continue
            skip_cols = set()
            listed = bool(re.match(r"^(?:[-*+]|\d+\.)\s", item))
            body_text = re.sub(r"^(?:[-*+]|\d+\.)\s+", "", item)
            # A quote right after "Icon" names a picture, not words on it.
            quotes = [next(g for g in m.groups() if g)
                      for m in _QUOTED.finditer(body_text)
                      if not re.search(r"\bIcon\s*$",
                                       body_text[:m.start()], re.I)]
            if quotes and not listed and quoted_lists:
                continue
            for quote in quotes:
                if len(words(quote)) >= 2 or _NUMBER.search(quote):
                    units.append(Unit(_strip_md(quote), name))
            if (quotes or _INSTRUCTION.match(body_text)
                    or body_text.rstrip().endswith(":")):
                # With quotes in it, an item is a description of the slide
                # text, and the quotes are the slide text.
                continue
            if _SHAPE_LINE.search(body_text):
                continue
            label = _BOLD_LABEL.match(body_text)
            if label:                            # "**Mai 2025** – GitHub MCP: …"
                units.append(Unit(_strip_md(label.group(1)), name))
                body_text = body_text[label.end():]
            if re.match(r"^(?:[-*+]|\d+\.)\s", item):
                text = _strip_md(re.sub(r"\s*\(Icon:[^)]*\)", "", body_text))
                if len(words(text)) >= 2:
                    units.append(Unit(text, name))
    # The same text twice is one requirement.
    seen, out = set(), []
    for unit in units:
        key = (norm(unit.text), unit.notes)
        if key not in seen:
            seen.add(key)
            out.append(unit)
    return out


def _present(unit_text: str, corpus: str, corpus_words: set) -> bool:
    n = norm(unit_text)
    if n in corpus:
        return True
    tokens = words(unit_text)
    if not tokens:
        return True
    numbers = _NUMBER.findall(unit_text)
    if any(num not in corpus for num in numbers):
        return False
    have = sum(1 for t in tokens if t in corpus_words)
    return have / len(tokens) >= WORDS_PRESENT


def missing_units(units, slide_texts, notes_texts):
    """``[Unit]`` the deck does not carry."""
    slides = norm(" \n ".join(slide_texts))
    notes = norm(" \n ".join(notes_texts))
    slide_words, note_words = set(words(slides)), set(words(notes))
    out = []
    for unit in units:
        corpus, vocab = (notes, note_words) if unit.notes else (slides, slide_words)
        if not _present(unit.text, corpus, vocab):
            out.append(unit)
    return out


def _sentences(text: str):
    for part in re.split(r"(?<=[.!?])\s+|\n+", text):
        part = part.strip(" •-–")
        if len(words(part)) >= MIN_SENTENCE_WORDS:
            yield part


def _pairs(tokens):
    return list(zip(tokens, tokens[1:]))


def invented(slides, source: str, baseline=frozenset()):
    """``[(slide, sentence)]`` on the slides that the source does not hold.

    ``slides`` is ``[(number, [text, ...])]``.  Template wording
    (``baseline``) is skipped: it is the template's, not the build's.
    """
    known = set(_pairs(words(source)))
    out = []
    for number, texts in slides:
        for text in texts:
            if text in baseline:
                continue
            for sentence in _sentences(text):
                pairs = _pairs(words(sentence))
                if not pairs:
                    continue
                share = sum(1 for p in pairs if p in known) / len(pairs)
                if share < PAIRS_SOURCED:
                    out.append((number, sentence))
    return out


def fidelity_findings(markdown: str, slides, notes, baseline=frozenset()):
    """``[(text, key)]`` QA findings for a deck against its source.

    ``slides`` is ``[(number, [text, ...])]``, ``notes`` ``[text, ...]``.
    """
    units = source_units(markdown)
    # Every text counts as present, the template's own included: a slide
    # titled "Agenda" carries "Agenda" whoever wrote it.
    flat = [t for _, texts in slides for t in texts]
    out = []
    for unit in missing_units(units, flat, notes):
        where = "the speaker notes" if unit.notes else "the slides"
        section = f" ({unit.section[:40]})" if unit.section else ""
        out.append((f"source{section}: {unit.text[:70]!r} is not in {where}, "
                    f"or not as written -- the source's wording, numbers "
                    f"included, goes on the slide unchanged",
                    ("content", 0, norm(unit.text)[:60])))
    for number, sentence in invented(slides, markdown, baseline):
        out.append((f"slide {number}: {sentence[:70]!r} is not in the source "
                    f"-- never add text the source does not give; a gap is "
                    f"\"[OFFEN: what is missing]\", not a sentence of your own",
                    ("content", number, norm(sentence)[:60])))
    return out
