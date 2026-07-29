#!/usr/bin/env python3
"""
Build, check and render a deck in one command -- the last thing you run.

    python3 scripts/make_deck.py build.py deck.pptx --template company-template.pptx
    python3 scripts/make_deck.py deck.pptx --template company-template.pptx
    python3 scripts/make_deck.py deck.pptx --no-render

Give it your build script and the deck that script writes, and it will:

1. run the build script,
2. QA the result -- text overflow, leftover template wording, placeholders,
3. schema-check the ``.pptx``,
4. render every slide to a PNG and print the paths,
5. print a contact sheet of the whole deck.

Then **open the PNGs and look at them.**  Steps 2 and 3 cannot see a slide;
they catch text that runs off the page and words nobody replaced, and they are
blind to a title colliding with its underline, a diagram drawn off-centre, or a
slide that is simply empty.  Every deck that shipped broken had passed a check
like this one and had never been looked at.

Exit codes: 0 clean, 1 the build failed or the checks found something, 2 bad
usage.
"""

import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from office.impress import PRESENTATION_EXTS, SofficeError, shutdown  # noqa: E402
from qa import check, validate_pptx  # noqa: E402

PPTX_EXTS = {".pptx", ".potx"}


def _run_script(script: Path) -> int:
    """Run a build script in its own interpreter, streaming its output."""
    print(f"# Building — {script.name}\n", flush=True)
    proc = subprocess.run([sys.executable, str(script)], cwd=script.parent)
    print(flush=True)
    return proc.returncode


def _render(deck: Path, outdir: Path, dpi: int) -> list:
    proc = subprocess.run(
        [sys.executable, str(HERE / "render.py"), str(deck),
         "--outdir", str(outdir), "--dpi", str(dpi)],
        capture_output=True, text=True,
    )
    if proc.returncode != 0:
        print(f"render.py failed:\n{proc.stdout}{proc.stderr}", file=sys.stderr)
        return []
    return [line for line in proc.stdout.splitlines() if line.strip()]


def _contact_sheet(deck: Path, outdir: Path):
    proc = subprocess.run(
        [sys.executable, str(HERE / "thumbnail.py"), str(deck),
         str(outdir / "contact-sheet")],
        capture_output=True, text=True,
    )
    if proc.returncode != 0:
        return None
    sheets = sorted(outdir.glob("contact-sheet*.jpg"))
    return sheets[0] if sheets else None


def main():
    parser = argparse.ArgumentParser(
        description="Build a deck, check it, and render it for review",
    )
    parser.add_argument(
        "args", nargs="+", metavar="[build.py] deck.pptx",
        help="The deck to check, optionally preceded by the script that builds it",
    )
    parser.add_argument("--template", help="The deck this was built from; flags "
                                           "text still shared with it")
    parser.add_argument("--outdir", default="render", help="Where to write PNGs")
    parser.add_argument("--dpi", type=int, default=110)
    parser.add_argument("--slides", help="Only QA these slides, e.g. 4,7-9")
    parser.add_argument("--no-render", action="store_true",
                        help="Skip the renders (you will not be able to review it)")
    parser.add_argument("--skip-validate", action="store_true")
    args = parser.parse_args()

    script, deck = None, None
    for item in args.args:
        path = Path(item)
        if path.suffix.lower() == ".py":
            script = path
        elif path.suffix.lower() in PRESENTATION_EXTS:
            deck = path
        else:
            print(f"Error: {item} is neither a .py build script nor a deck",
                  file=sys.stderr)
            return 2
    if deck is None:
        print("Error: name the deck to check, e.g. make_deck.py deck.pptx",
              file=sys.stderr)
        return 2
    if script is not None and not script.exists():
        print(f"Error: {script} does not exist", file=sys.stderr)
        return 2
    if args.template and not Path(args.template).exists():
        print(f"Error: {args.template} does not exist", file=sys.stderr)
        return 2

    if script is not None:
        code = _run_script(script)
        if code != 0:
            print(f"Build failed (exit {code}) — nothing to check.", file=sys.stderr)
            return 1
    if not deck.exists():
        print(f"Error: {deck} was not written", file=sys.stderr)
        return 1

    failed = False

    print(f"# Checking — {deck.name}\n")
    wanted = None
    if args.slides:
        from qa import _parse_slides
        wanted = _parse_slides(args.slides)
    try:
        issues, checked = check(deck, args.template, wanted)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    print(f"{checked} slide(s) checked.\n")
    if issues:
        failed = True
        print(f"## {len(issues)} issue(s) — fix every one\n")
        for issue in issues:
            print(f"- {issue}")
    else:
        print("No overflow and no leftover template text.")

    if deck.suffix.lower() in PPTX_EXTS and not args.skip_validate:
        code, output = validate_pptx(deck, args.template)
        print("\n## File integrity\n")
        if code == 0:
            print("validate.py: passed")
        else:
            failed = True
            print(f"validate.py: FAILED (exit {code})")
            print(output)

    if not args.no_render:
        outdir = Path(args.outdir)
        print("\n## Renders\n")
        images = _render(deck, outdir, args.dpi)
        for image in images:
            print(image)
        sheet = _contact_sheet(deck, outdir)
        if sheet:
            print(f"\ncontact sheet: {sheet.resolve()}")
        if images:
            print(
                "\n**Open these and look at them.** The checks above cannot see "
                "a slide: they will not tell you that a diagram is off-centre, "
                "that two shapes collide, that a slide is half empty, or that "
                "one is still in the template's language."
            )

    shutdown()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
