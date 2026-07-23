#!/usr/bin/env python3
"""
QA a finished deck without looking at it: overflow, leftovers, file integrity.

    python3 scripts/qa.py deck.pptx
    python3 scripts/qa.py deck.pptx --template Kite_brand_master.pptx
    python3 scripts/qa.py deck.odp --slides 3,7-9

Three checks, in the order they catch real defects:

1. **Text overflow** -- every text shape is re-laid out at its own width and the
   height it wants is compared with the height it has.  This is the one thing
   the renders show that a text dump does not, so it is what makes the deck
   checkable when you cannot see the slides.
2. **Leftover text** -- generic placeholders (lorem, TODO, xxx, [insert]), and,
   with ``--template``, any string still shared verbatim with the template you
   built from.  Retexting a designed slide by z-order index puts text in the
   wrong box and leaves the template's own words behind; this finds that.
3. **File integrity** -- runs ``office/validate.py`` on a ``.pptx``.  ODF needs
   no equivalent: LibreOffice writes it itself.

Exit codes: 0 clean, 1 issues found, 2 bad usage or unreadable file.

**This does not replace looking at the slides.** It cannot see alignment,
contrast, crowding or a title colliding with its underline.  Render and look
whenever you can -- see the QA section of SKILL.md.
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import (  # noqa: E402
    PRESENTATION_EXTS, SofficeError, load_document, shapes, slides,
)
from office.template import overflowing  # noqa: E402

PPTX_EXTS = {".pptx", ".potx"}

PLACEHOLDER = re.compile(
    r"\b(lorem|ipsum|todo|tbd|fixme|platzhalter|blindtext)\b"
    r"|\bx{3,}\b"
    r"|\[\s*(insert|insert [^\]]*|your [^\]]*|name|date|company)\s*\]",
    re.I,
)

# Only prose this long counts as leftover template wording.  A designed deck
# reuses its template's short headers on purpose -- "Agenda", "What we found" --
# and flagging those buries the real defect, which is always a body paragraph
# nobody replaced.
MIN_LEFTOVER = 25


def _parse_slides(spec):
    out = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, hi = part.split("-", 1)
            out.update(range(int(lo), int(hi) + 1))
        else:
            out.add(int(part))
    return out


def _norm(text):
    return " ".join(text.split())


def _texts(doc):
    """``{slide number: {shape index: text}}`` for every shape carrying text."""
    out = {}
    for number, page in enumerate(slides(doc), start=1):
        found = {}
        for i, shape in enumerate(shapes(page)):
            try:
                text = _norm(shape.getString())
            except Exception:
                continue
            if text:
                found[i] = text
        out[number] = found
    return out


def _template_strings(path):
    doc = load_document(path, read_only=True)
    try:
        return {
            text
            for found in _texts(doc).values()
            for text in found.values()
            if len(text) >= MIN_LEFTOVER
        }
    finally:
        doc.close(False)


def check(path, template=None, wanted=None):
    """``(issues, slides_checked)`` -- issues are ready-to-print strings."""
    issues = []
    tpl_strings = _template_strings(template) if template else set()

    doc = load_document(path, read_only=True)
    try:
        pages = slides(doc)
        checked = 0
        for number, page in enumerate(pages, start=1):
            if wanted and number not in wanted:
                continue
            checked += 1

            for index, reason in overflowing(doc, page):
                issues.append(f"slide {number} shape [{index}]: {reason}")

            for index, shape in enumerate(shapes(page)):
                try:
                    text = _norm(shape.getString())
                except Exception:
                    continue
                if not text:
                    continue
                hit = PLACEHOLDER.search(text)
                if hit:
                    issues.append(
                        f"slide {number} shape [{index}]: placeholder text "
                        f"{hit.group(0)!r} in {text[:60]!r}"
                    )
                elif text in tpl_strings:
                    issues.append(
                        f"slide {number} shape [{index}]: still the template's "
                        f"own wording -- {text[:60]!r}"
                    )
        return issues, checked
    finally:
        doc.close(False)


def validate_pptx(path):
    script = Path(__file__).resolve().parent / "office" / "validate.py"
    proc = subprocess.run(
        [sys.executable, str(script), str(path)],
        capture_output=True, text=True,
    )
    return proc.returncode, (proc.stdout + proc.stderr).strip()


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
        print(
            f"Error: {path} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        return 2
    if args.template and not Path(args.template).exists():
        print(f"Error: {args.template} does not exist", file=sys.stderr)
        return 2

    wanted = _parse_slides(args.slides) if args.slides else None

    try:
        issues, checked = check(path, args.template, wanted)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    print(f"# QA {path.name}")
    print(f"\n{checked} slide(s) checked.\n")

    if issues:
        print(f"## {len(issues)} issue(s)\n")
        for issue in issues:
            print(f"- {issue}")
    else:
        print("No overflow and no leftover text.")

    failed = bool(issues)

    if path.suffix.lower() in PPTX_EXTS and not args.skip_validate:
        code, output = validate_pptx(path)
        print("\n## File integrity\n")
        if code == 0:
            print("validate.py: passed")
        else:
            failed = True
            print(f"validate.py: FAILED (exit {code})")
            print(output)
            print(
                "\nIf this deck came from a template, re-run with "
                "`office/validate.py --original <template>` -- the template's "
                "own defects otherwise read as yours."
            )

    print(
        "\nThis check cannot see the slides. Render them "
        "(`python3 scripts/render.py`) and look, whenever you can."
    )
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
