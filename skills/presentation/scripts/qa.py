#!/usr/bin/env python3
"""
QA a finished deck without looking at it: overflow, leftovers, file integrity.

    python3 scripts/qa.py deck.pptx
    python3 scripts/qa.py deck.pptx --template Kite_brand_master.pptx
    python3 scripts/qa.py deck.odp --slides 3,7-9

The checks, in the order they catch real defects:

1. **Text overflow** -- every text shape is re-laid out at its own width and the
   height it wants is compared with the height it has.  This is the one thing
   the renders show that a text dump does not, so it is what makes the deck
   checkable when you cannot see the slides.
2. **Collision** -- two blocks of text printed over each other.  Not an
   overrun: both fit their own boxes, and the boxes were placed on top of one
   another to begin with.  This is what hand-placing labels around a diagram
   produces, and until it was measured every check here passed on decks whose
   renders showed two sentences through each other.
3. **Off the slide** -- a shape whose geometry hangs past the trim, which is
   what a diagram sized by eye does while every text check still passes.
4. **Leftover text** -- generic placeholders (lorem, TODO, xxx, [insert]), and,
   with ``--template``, any string still shared verbatim with the template you
   built from.  Retexting a designed slide by z-order index puts text in the
   wrong box and leaves the template's own words behind; this finds that.  It
   covers **the footer** too, which is a page property rather than a shape and
   which nothing that walks shapes can see: four decks shipped with the house
   template's German ``subject | title | place | date`` under English slides,
   and all four reported clean.
5. **Markdown on a slide** -- ``~~struck~~`` and ``**bold**`` render as literal
   tildes and asterisks.  They arrive by pasting drafted wording into
   ``set_text`` unchanged.  Alongside it, **emoji** (a colour-font picture the
   template has no say over) and **type under the smallest size the template
   itself sets**, which is legible on the monitor it was built on and not from
   a seat in the room.
6. **Contrast** -- each run's colour against the fill behind it, or the slide
   background.  A headline drawn in the default black onto a dark divider is
   the right words in the right box and invisible, and nothing else here sees
   it.
7. **Occlusion** -- an opaque shape covering the one beneath it, which is what
   a ``cut_*`` pair unpacked in the wrong order looks like on the slide.
8. **Stray template pictures** -- with ``--template``, a picture from the
   template still sitting on a slide.  ``retext`` replaces words and never
   pictures, so a design's stock photo survives a deck that reports clean.
9. **File integrity** -- runs ``office/validate.py`` on a ``.pptx``.  ODF needs
   no equivalent: LibreOffice writes it itself.

Exit codes: 0 clean, 1 issues found, 2 bad usage or unreadable file.

**This still does not replace looking at the slides.** It cannot see that a
diagram is off-centre, that a slide is half empty, or that a Venn's circles do
not overlap.  Render and look whenever you can -- see the QA section of
SKILL.md.
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.deckbuild import find_template  # noqa: E402
from office.impress import (PPTX_EXTS, PRESENTATION_EXTS,  # noqa: E402
                            SofficeError, parse_slide_spec)
from office.qa import QAResult, check, summary, validate_pptx  # noqa: E402,F401


def main():
    parser = argparse.ArgumentParser(
        description="Check a deck for overflow, leftover text and file defects"
    )
    parser.add_argument("input", help="Deck to check")
    parser.add_argument(
        "--template",
        help="The deck this was built from; flags text still shared with it",
    )
    parser.add_argument("--slides", help="Only these slides, e.g. 4,7-9")
    parser.add_argument(
        "--skip-validate", action="store_true",
        help="Skip the .pptx schema check",
    )
    args = parser.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"Error: {path} does not exist", file=sys.stderr)
        return 2
    if path.suffix.lower() not in PRESENTATION_EXTS:
        print(f"Error: {path} is not a presentation; expected one of: "
              f"{', '.join(sorted(PRESENTATION_EXTS))}", file=sys.stderr)
        return 2
    # Same resolver build_deck uses: the template path in the documentation is
    # written relative to the skill, and a build script runs next to the deck.
    if args.template:
        try:
            args.template = str(find_template(args.template))
        except FileNotFoundError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    wanted = None
    if args.slides:
        try:
            wanted = set(parse_slide_spec(args.slides))
        except ValueError as e:
            print(f"Error: --slides {e}", file=sys.stderr)
            return 2

    try:
        result = check(path, args.template, wanted)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    print(f"# QA {path.name}\n")
    print("\n".join(summary(result)))
    failed = bool(result.issues)

    if path.suffix.lower() in PPTX_EXTS and not args.skip_validate:
        code, output = validate_pptx(path, args.template)
        print("\n## File integrity\n")
        if code == 0:
            print("validate.py: passed")
        else:
            failed = True
            print(f"validate.py: FAILED (exit {code})")
            print(output)
            if not args.template:
                print("\nIf this deck came from a template, re-run with "
                      "`--template <template>` -- the template's own schema "
                      "defects otherwise read as yours.")

    print("\nThis check cannot see the slides: make_deck.py renders them, and "
          "vision_brief.py hands them to the vision-agent.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
