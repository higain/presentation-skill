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

The image a brief points at is the render with a review code printed under it
(``office.review``); an answer counts only when it gives that code back.  A
slide whose render has not changed keeps the answer it has, so after a fix
only the slides that changed are sent again.

``slides`` needs ``make_deck.py``'s report and renders.  ``designs`` writes one
brief per design into a catalog directory ``catalog_extract.py`` made; the
answers are read by ``presentation-setup``'s ``name_designs.py``.
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))

from office import review  # noqa: E402
from office.vision import design_brief, slide_brief, verdict_line  # noqa: E402

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


def write_slide_briefs(deck: Path, outdir: Path, language: str) -> tuple:
    """``(briefs, pending)``: a brief beside every render, and the slides that
    still need a vision-agent call.

    A slide whose render has not changed since its answer was saved keeps
    that answer: it is not in ``pending`` and needs no new call.
    """
    report = _report(deck)
    if not report.is_file():
        raise SystemExit(f"No report at {report}. Run make_deck.py on the "
                         f"deck first; it writes the report and the renders.")
    measured = _measured(report)
    if not review.renders(outdir):
        raise SystemExit(f"No slide-N.png in {outdir}. Run make_deck.py "
                         f"without --no-render.")
    written, pending = [], []
    images = review.prepare(outdir)
    ledger = review.load(outdir)
    for number, (image, _fresh) in images.items():
        brief = review.brief_path(outdir, number)
        brief.write_text(slide_brief(str(image.absolute()),
                                     measured.get(number, ""), language),
                         encoding="utf-8")
        written.append(brief)
        if review.verify(outdir, number, ledger)[0] != "ok":
            pending.append(brief)
    return written, pending


def collect(deck: Path, outdir: Path) -> tuple[int, dict]:
    """Fill the report's verdict cells from the verified answers.

    Returns ``(filled, problems)`` with ``problems`` as ``{slide: status}``
    for every slide whose answer does not count -- missing, unreadable, or
    without the review code printed on its image (``office.review``).  A
    verdict already in the table is replaced by the answer that verifies;
    the table is a view of the answers, never a place to decide one.
    """
    report = _report(deck)
    lines = report.read_text(encoding="utf-8").splitlines()
    ledger = review.load(outdir)
    filled, problems = 0, {}
    for i, line in enumerate(lines):
        match = ROW.match(line.strip())
        if not match:
            continue
        number, measured = int(match.group(1)), match.group(2).strip()
        status, parsed = review.verify(outdir, number, ledger)
        cell = verdict_line(parsed) if status == "ok" else ""
        if status == "ok" and not cell:
            status = "unparsed"
        if status != "ok":
            problems[number] = status
            lines[i] = f"| {number} | {measured} | |"
            continue
        lines[i] = f"| {number} | {measured} | {cell} |"
        filled += 1
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return filled, problems


def write_design_briefs(catalog: Path) -> list:
    """A brief in every design folder; returns those still needing an answer.

    A design whose render has not changed since its answer was saved keeps
    that answer, so re-running after a re-extraction sends only what changed.
    """
    manifest = catalog / "manifest.json"
    if not manifest.is_file():
        raise SystemExit(f"No manifest.json in {catalog}. Point this at the "
                         f"directory catalog_extract.py wrote.")
    numbers = [row["number"] for row in
               json.loads(manifest.read_text())["slides"]]
    images = review.prepare_designs(catalog, numbers)
    written = []
    for number, (image, _fresh) in images.items():
        folder = review.design_folder(catalog, number)
        structure = json.loads((folder / "structure.json").read_text())
        brief = folder / "brief.md"
        brief.write_text(design_brief(str(image.absolute()), structure),
                         encoding="utf-8")
        if review.verify_design(catalog, number)[0] != "ok":
            written.append(brief)
    return written


def _delegation(briefs: list, answer_name) -> str:
    first = briefs[0]
    return (f"\nNext: one vision-agent call PER brief -- never several images "
            f"in one call. The whole delegation is:\n\n"
            f"    Read the brief at {first} and follow it exactly.\n\n"
            f"Save the vision-agent's whole reply, unedited, as "
            f"{answer_name(first)}, and the same for every brief.\n"
            f"Send every brief, not a sample. If a reply is CANNOT SEE "
            f"IMAGE, send that same brief again -- up to three tries; the "
            f"failure is often transient. Only then save CANNOT SEE IMAGE as "
            f"its answer and tell the user which slides nothing has seen. "
            f"**Never write, repair or restore an answer yourself**: each "
            f"image carries a review code under the slide that only "
            f"something with eyes can read, and an answer without it is "
            f"thrown away.")


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
        if not briefs:
            print("Every design already has an answer for its current "
                  "render -- nothing to send.")
            return 0
        print(f"{len(briefs)} design(s) to send: " + ", ".join(
            str(int(b.parent.name.split("-")[1])) for b in briefs))
        print(_delegation(briefs, lambda b: b.with_name("answer.md")))
        print("\nThen run name_designs.py on the catalog: it reads every "
              "answer.md into the worksheet for review.")
        return 0

    outdir = args.outdir or args.deck.resolve().parent / "render"
    if args.cmd == "slides":
        briefs, pending = write_slide_briefs(args.deck, outdir, args.language)
        print(f"Wrote {len(briefs)} slide brief(s) into {outdir}/")
        done = len(briefs) - len(pending)
        if done:
            print(f"{done} slide(s) keep the answer they already have -- "
                  f"their renders did not change. Send only the "
                  f"{len(pending)} brief(s) below.")
        if pending:
            print("Slides to send: " + ", ".join(
                b.name.split(".")[0].split("-")[1] for b in pending))
            print(_delegation(pending, lambda b: b.with_name(
                b.name.replace(".brief.md", ".answer.md"))))
        print(f"\nThen: python3 {sys.argv[0]} collect {args.deck}")
        return 0

    filled, problems = collect(args.deck, outdir)
    print(f"Filled {filled} verdict(s) in {_report(args.deck).name}.")
    for status in ("no-answer", "cannot-see", "unparsed", "no-code",
                   "wrong-code", "unbriefed"):
        slides = [n for n, s in sorted(problems.items()) if s == status]
        if slides:
            why = review.STATUS.get(status, "no JSON verdict in the answer")
            print(f"Not counted -- {why}: slide(s) "
                  + ", ".join(map(str, slides)))
    if any(s in ("no-code", "wrong-code") for s in problems.values()):
        print("An answer without the right review code was not written by "
              "something that saw the image. Send that slide's brief to the "
              "vision-agent again and save its whole reply. Never type or "
              "repair an answer yourself.")
    print("\nEvery `fix` verdict is work: change the build script, re-run "
          "make_deck.py, and brief the new renders. Overrule a finding only "
          "with a measurement, and only with the user's approval "
          "(references/qa.md).")
    return 1 if problems else 0

if __name__ == "__main__":
    sys.exit(main())
