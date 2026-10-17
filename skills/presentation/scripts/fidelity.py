#!/usr/bin/env python3
"""Check a deck against the source it was built from: nothing dropped,
nothing changed, nothing invented.

    python3 scripts/fidelity.py deck.pptx content.md [--template acme]

Every quote, table cell, plain list item, slide heading and speaker-note
line of the Markdown source has to be in the deck as written; every
sentence of six words or more on a slide has to come from the source.
``make_deck.py --source content.md`` runs the same check inside the gate.
Exit 1 on any finding.  See ``office/qa/fidelity.py`` for what counts.
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))

from office.impress import load_document, notes_text, slides  # noqa: E402
from office.pack import active  # noqa: E402
from office.qa.fidelity import fidelity_findings  # noqa: E402
from office.qa.text import _page_texts  # noqa: E402


def deck_texts(path):
    """``([(number, [text, ...])], [notes, ...])`` for a deck."""
    doc = load_document(Path(path), read_only=True)
    try:
        out, notes = [], []
        for number, page in enumerate(slides(doc), start=1):
            out.append((number, [text for _, text in _page_texts(page)]))
            try:
                notes.append(notes_text(page))
            except Exception:
                notes.append("")
        return out, notes
    finally:
        doc.close(False)


def check_fidelity(deck, source, template=None):
    """``[(text, key)]`` -- the findings, as the gate reports them."""
    from office.deckbuild import find_template
    from office.qa.baseline import _template_baseline
    template = str(find_template(template)) if template else None
    with active(template):
        texts, notes = deck_texts(deck)
        baseline = _template_baseline(template).strings if template else frozenset()
    markdown = Path(source).read_text(encoding="utf-8")
    return fidelity_findings(markdown, texts, notes, baseline)


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("deck", type=Path)
    ap.add_argument("source", type=Path, help="the Markdown the deck was built from")
    ap.add_argument("--template", help="the pack, so its own wording is not "
                                       "reported as invented")
    args = ap.parse_args()
    findings = check_fidelity(args.deck, args.source, args.template)
    for text, _ in findings:
        print(f"- {text}")
    print(f"\n{len(findings)} finding(s).")
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
