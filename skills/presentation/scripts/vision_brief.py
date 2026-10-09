#!/usr/bin/env python3
"""Write the vision-agent's briefs, one per image, and collect its answers.

    python3 scripts/vision_brief.py slides deck.pptx [--language German]
    python3 scripts/vision_brief.py collect deck.pptx
    python3 scripts/vision_brief.py designs templates/acme/catalog

The build model cannot see images and must not write the brief itself: a brief
it writes asks about nine slides at once and gets "all ok" back.  This writes a
tested brief beside every image, so the delegation is one line --

    Read the brief at <dir>/slide-3.brief.md and follow it exactly.

-- and the build model saves the vision-agent's whole reply, unedited, beside
it as ``slide-3.answer.md``.  ``collect`` then reads every answer and fills
the report's verdict table, so nobody hand-edits a markdown table either.

``slides`` needs ``make_deck.py``'s report and renders.  ``designs`` writes one
brief per design into a catalog directory ``catalog_extract.py`` made; the
answers are read by ``presentation-setup``'s ``name_designs.py``.
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.vision import (design_brief, last_json,  # noqa: E402
                           slide_brief, verdict_line)

ROW = re.compile(r"^\|\s*(\d+)\s*\|([^|]*)\|([^|]*)\|$")


def _report(deck: Path) -> Path:
    return deck.resolve().with_name(deck.stem + "-report.md")


def _measured(report: Path) -> dict:
    """``{slide: measured cell}`` from the report's verdict table."""
    out = {}
    for line in report.read_text(encoding="utf-8").splitlines():
        match = ROW.match(line.strip())
        if match:
            out[int(match.group(1))] = match.group(2).strip()
    return out


def _renders(outdir: Path) -> dict:
    return {int(m.group(1)): png for png in outdir.glob("slide-*.png")
            if (m := re.fullmatch(r"slide-(\d+)\.png", png.name))}


def write_slide_briefs(deck: Path, outdir: Path, language: str) -> list:
    report = _report(deck)
    if not report.is_file():
        raise SystemExit(f"No report at {report}. Run make_deck.py on the "
                         f"deck first; it writes the report and the renders.")
    measured = _measured(report)
    renders = _renders(outdir)
    if not renders:
        raise SystemExit(f"No slide-N.png in {outdir}. Run make_deck.py "
                         f"without --no-render.")
    written = []
    for number, png in sorted(renders.items()):
        brief = outdir / f"slide-{number}.brief.md"
        brief.write_text(slide_brief(str(png.resolve()),
                                     measured.get(number, ""), language),
                         encoding="utf-8")
        # A stale answer belongs to an earlier build of this slide.
        (outdir / f"slide-{number}.answer.md").unlink(missing_ok=True)
        written.append(brief)
    return written


def collect(deck: Path, outdir: Path) -> tuple[int, list, list]:
    """Fill the report's verdict cells from ``slide-N.answer.md``.

    Returns ``(filled, unanswered, unparsed)``.  A verdict already written by
    hand is kept: a person overrides the vision-agent, never the reverse.
    """
    report = _report(deck)
    lines = report.read_text(encoding="utf-8").splitlines()
    filled, unanswered, unparsed = 0, [], []
    for i, line in enumerate(lines):
        match = ROW.match(line.strip())
        if not match:
            continue
        number, measured, verdict = (int(match.group(1)),
                                     match.group(2).strip(),
                                     match.group(3).strip())
        if verdict:
            continue
        answer = outdir / f"slide-{number}.answer.md"
        if not answer.is_file():
            unanswered.append(number)
            continue
        cell = verdict_line(last_json(answer.read_text(encoding="utf-8")))
        if not cell:
            unparsed.append(number)
            continue
        lines[i] = f"| {number} | {measured} | {cell} |"
        filled += 1
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return filled, unanswered, unparsed


def write_design_briefs(catalog: Path) -> list:
    manifest = catalog / "manifest.json"
    if not manifest.is_file():
        raise SystemExit(f"No manifest.json in {catalog}. Point this at the "
                         f"directory catalog_extract.py wrote.")
    written = []
    for row in json.loads(manifest.read_text())["slides"]:
        folder = catalog / f"slide-{row['number']:02d}"
        structure = json.loads((folder / "structure.json").read_text())
        brief = folder / "brief.md"
        brief.write_text(design_brief(str((folder / "render.png").resolve()),
                                      structure), encoding="utf-8")
        written.append(brief)
    return written


def _delegation(briefs: list, answer_name) -> str:
    first = briefs[0]
    return (f"\nNext: one vision-agent call PER brief -- never several images "
            f"in one call. The whole delegation is:\n\n"
            f"    Read the brief at {first} and follow it exactly.\n\n"
            f"Save the vision-agent's whole reply, unedited, as "
            f"{answer_name(first)}, and the same for every brief.")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    slides = sub.add_parser("slides", help="one brief per rendered slide")
    slides.add_argument("deck", type=Path)
    slides.add_argument("--outdir", type=Path,
                        help="the renders (default: render/ beside the deck)")
    slides.add_argument("--language", default="",
                        help="the deck's language, e.g. German (default: "
                             "whatever the slide titles are written in)")
    coll = sub.add_parser("collect", help="fill the verdicts from the answers")
    coll.add_argument("deck", type=Path)
    coll.add_argument("--outdir", type=Path)
    designs = sub.add_parser("designs", help="one brief per template design")
    designs.add_argument("catalog", type=Path)
    args = ap.parse_args()

    if args.cmd == "designs":
        briefs = write_design_briefs(args.catalog)
        print(f"Wrote {len(briefs)} design brief(s) into {args.catalog}/")
        print(_delegation(briefs, lambda b: b.with_name("answer.md")))
        print("\nThen run name_designs.py on the catalog: it reads every "
              "answer.md into the worksheet for review.")
        return 0

    outdir = args.outdir or args.deck.resolve().parent / "render"
    if args.cmd == "slides":
        briefs = write_slide_briefs(args.deck, outdir, args.language)
        print(f"Wrote {len(briefs)} slide brief(s) into {outdir}/")
        print(_delegation(briefs, lambda b: b.with_name(
            b.name.replace(".brief.md", ".answer.md"))))
        print(f"\nThen: python3 {sys.argv[0]} collect {args.deck}")
        return 0

    filled, unanswered, unparsed = collect(args.deck, outdir)
    print(f"Filled {filled} verdict(s) in {_report(args.deck).name}.")
    if unanswered:
        print("No answer yet for slide(s): "
              + ", ".join(map(str, unanswered)))
    if unparsed:
        print("No JSON object in the answer for slide(s): "
              + ", ".join(map(str, unparsed))
              + " -- ask the vision-agent again with the same brief.")
    print("\nEvery `fix` verdict is work: change the build script, re-run "
          "make_deck.py, and brief the new renders. Overrule a finding only "
          "with a measurement (references/qa.md).")
    return 1 if unanswered or unparsed else 0


if __name__ == "__main__":
    sys.exit(main())
