#!/usr/bin/env python3
"""Candidate authoring notes: the words a template says to its builder.

    python3 scripts/find_notes.py brand.pptx

A pattern-collection template talks to whoever builds from it -- "EXAMPLE",
"Customer logo here", a paragraph of instructions the designer left on slide
four. They are ordinary shapes. Nothing flags them, `retext` has no reason to
touch them, and they ship: a finished customer deck with the word *Beispiel*
still on three slides.

`strip_annotations` deletes them, and it needs the list. This proposes one;
a person ticks it. **Only a person can do the ticking** -- "Example" is an
annotation on one template and a column header on another, and getting it wrong
either leaves instructions in a customer deck or silently deletes real content.

Three signals, and a string only has to trip one:

* **it repeats across designs.** A short string on five different slides of a
  pattern collection is furniture, not content.
* **it is highlighted.** Designers mark fill-in fields in yellow. That is the
  most reliable signal in the file and the one no model would think to look at.
* **it reads as an instruction** -- an imperative opening, or several sentences
  in a box too small to be body copy.
"""

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine  # noqa: E402

engine()

from office.impress import load_document, slides  # noqa: E402
from office.template import read_text, walk_shapes  # noqa: E402

#: Fills a designer marks a fill-in field with.
HIGHLIGHTS = {0xFFFF00, 0xF5EB05, 0xFFFF99, 0xFFF200}

#: A string on this many designs or more is the template talking to itself.
#: Two, not three: missing one ships it in a customer deck, and an extra row
#: costs a glance. Raising it hid the customer-logo marker on a real template.
REPEATS_ON = 2

#: Longer than this and it is content, not a label -- unless it opens like an
#: instruction, which is what `IMPERATIVE` is for.
LABEL_CHARS = 40

#: How a designer's note to the builder starts, in the two languages this
#: skill's checks cover. Deliberately short: a false positive here costs a
#: tick-box, a false negative costs a shipped deck.
IMPERATIVE = re.compile(
    r"^\s*(please|insert|replace|add|delete|use |note:|example|sample"
    r"|bitte|einfügen|ersetzen|hinzufügen|verwenden|hinweis:|beispiel"
    r"|vorgehen:|anleitung)",
    re.I,
)


#: A candidate has to contain a letter.  Step numbers "1".."5" repeat across
#: every process design in a template and are content on all of them; the same
#: rule keeps `qa.py` from reporting digits as leftover template wording.
HAS_LETTER = re.compile(r"[^\W\d_]")


def _norm(text):
    return " ".join(text.split())


def survey(path):
    """``{text: {"slides": set, "highlighted": bool, "chars": int}}``."""
    doc = load_document(path, read_only=True)
    found = defaultdict(lambda: {"slides": set(), "highlighted": False})
    try:
        for number, page in enumerate(slides(doc), start=1):
            for _address, shape in walk_shapes(page):
                fill = None
                if "SOLID" in str(getattr(shape, "FillStyle", "")):
                    fill = int(getattr(shape, "FillColor", -1)) & 0xFFFFFF
                try:
                    paragraphs = read_text(shape)
                except Exception:
                    continue
                for body, _cprops, _pprops in paragraphs:
                    text = _norm(body)
                    if not text:
                        continue
                    entry = found[text]
                    entry["slides"].add(number)
                    if fill in HIGHLIGHTS:
                        entry["highlighted"] = True
    finally:
        doc.close(False)
    return found


def rank(found):
    """``[(text, reasons, slides)]``, most likely an annotation first."""
    out = []
    for text, entry in found.items():
        reasons = []
        if entry["highlighted"]:
            reasons.append("highlighted")
        if len(entry["slides"]) >= REPEATS_ON and len(text) <= LABEL_CHARS:
            reasons.append(f"on {len(entry['slides'])} designs")
        if IMPERATIVE.match(text):
            reasons.append("reads as an instruction")
        if reasons and HAS_LETTER.search(text):
            out.append((text, reasons, sorted(entry["slides"])))
    out.sort(key=lambda row: (-len(row[1]), -len(row[2]), row[0]))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("template")
    args = ap.parse_args()

    ranked = rank(survey(args.template))
    if not ranked:
        print("No candidate authoring notes. Either the template is clean, or "
              "its notes look like ordinary content — read a render or two "
              "before concluding the first.")
        return 0

    print(f"# Candidate authoring notes — {args.template}\n")
    print("Tick the ones that are the template talking to its builder rather "
          "than content. **Read them; do not accept the list.** Deleting a "
          "column header is worse than leaving a label in.\n")
    print("| # | text | why it is here | designs |")
    print("|---:|---|---|---|")
    for index, (text, reasons, slides_on) in enumerate(ranked, start=1):
        shown = text if len(text) <= 60 else text[:57] + "…"
        print(f"| {index} | `{shown}` | {', '.join(reasons)} | "
              f"{', '.join(str(n) for n in slides_on[:8])}"
              f"{'…' if len(slides_on) > 8 else ''} |")

    print("\n**Whole-string matches go in `exact`; a long instruction block "
          "goes in `prefixes`** — match its first few words, because the rest "
          "of the paragraph is where a designer's typo lives.\n")
    print("Then: `make_pack.py --name <pack> --template " + args.template
          + " --note '<text>' --note-prefix '<first words>'`")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
