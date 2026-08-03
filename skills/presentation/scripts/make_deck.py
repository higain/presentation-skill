#!/usr/bin/env python3
"""
Build, check and render a deck in one command -- the last thing you run.

    python3 scripts/make_deck.py build.py deck.pptx --template company-template.pptx
    python3 scripts/make_deck.py deck.pptx --template company-template.pptx
    python3 scripts/make_deck.py deck.pptx --no-render

Give it your build script and the deck that script writes, and it will:

1. run the build script,
2. QA the result -- overflow, off-slide shapes, leftover template wording and
   pictures, placeholders, markdown, contrast, occlusion,
3. schema-check the ``.pptx``,
4. write the findings to ``<deck>-report.md`` beside the deck,
5. render every slide to a PNG in ``render/`` beside the deck, and print paths,
6. print a contact sheet of the whole deck.

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


def _start_render(deck: Path, outdir: Path, dpi: int):
    """Kick the renders off and return the process, without waiting.

    Rendering and checking are independent -- both only read the built deck --
    and each is a LibreOffice round trip of roughly the same length.  Run in
    series they add up to more than the two minutes a shell typically allows,
    and a run that gets cut off before the renders exist is a run that ships
    without anyone looking.  Overlapped, they cost the longer of the two.
    """
    return subprocess.Popen(
        [sys.executable, str(HERE / "render.py"), str(deck),
         "--outdir", str(outdir), "--dpi", str(dpi)],
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
    )


def _finish_render(proc) -> list:
    out, err = proc.communicate()
    if proc.returncode != 0:
        print(f"render.py failed:\n{out}{err}", file=sys.stderr)
        return []
    return [line for line in out.splitlines() if line.strip()]


def _write_report(deck: Path, issues: list, checked: int, outdir=None) -> Path:
    """Leave the findings on disk, next to the deck, as ``<deck>-report.md``.

    A build that ends between "here is what is wrong" and "fixed it" otherwise
    leaves a deck that looks finished and a set of findings that exist only in
    a transcript nobody will re-read.  On disk, the next thing to do is
    obvious, and a clean run overwrites it with a clean report.

    The renders are named here too, because this whole command takes a couple
    of minutes and a shell that gives up at two will have thrown away the paths
    it printed.  One ``cat`` of this file gets them back.
    """
    report = deck.resolve().with_name(deck.stem + "-report.md")
    lines = [
        f"# {deck.name}", "",
        f"{checked} slide(s) checked.", "",
    ]
    if issues:
        lines += [f"## {len(issues)} unfixed issue(s)", "",
                  "**This deck is not finished.** Every line below is a defect "
                  "in the file as it stands now.", ""]
        lines += [f"- [ ] {issue}" for issue in issues]
    else:
        lines += ["## No automated findings", "",
                  "Overflow, collisions, off-slide geometry, leftover template "
                  "text and pictures, markdown, contrast and occlusion all "
                  "clean.", "",
                  "This says nothing about whether the slides look right. "
                  "Open the renders."]
    if outdir is not None:
        lines += ["", "## Renders", "",
                  f"- [ ] Look at every slide in `{outdir}`",
                  f"- [ ] Start with the contact sheet, "
                  f"`{outdir / 'contact-sheet-1.jpg'}`", "",
                  "Nothing above was seen by anything with eyes."]
    try:
        report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        return None
    return report


def _contact_sheet(deck: Path, outdir: Path):
    # Built from the PNGs already on disk: converting the deck again here was
    # a second full LibreOffice round trip for a second copy of the same
    # pixels, and it was pure tail latency on the command's slowest step.
    proc = subprocess.run(
        [sys.executable, str(HERE / "thumbnail.py"), str(deck),
         str(outdir / "contact-sheet"), "--from-images", str(outdir)],
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
    # Deliberately not "render": a relative default resolves against the
    # working directory, so running this from the skill folder wrote the PNGs
    # into the skill instead of beside the deck.  A run that cannot reach its
    # own renders cannot look at them, and ships whatever it built.
    parser.add_argument("--outdir", default=None,
                        help="Where to write PNGs (default: render/ beside the deck)")
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

    outdir = (Path(args.outdir) if args.outdir
              else deck.resolve().parent / "render")

    # Said before any of the work, because all of it together runs to a couple
    # of minutes and a shell that times out at one loses everything printed
    # after this point.  Whatever happens, the two paths worth knowing are up
    # here, at second zero.
    if not args.no_render:
        print(f"# {deck.name} — findings will be in "
              f"{deck.resolve().with_name(deck.stem + '-report.md')},\n"
              f"# renders in {outdir}/\n"
              f"# Both are written to disk. This takes a minute or two, so "
              f"give it a five-minute timeout;\n"
              f"# if your shell cut it off anyway, read those two paths "
              f"rather than building again.\n", flush=True)

    if script is not None:
        code = _run_script(script)
        if code != 0:
            print(f"Build failed (exit {code}) — nothing to check.", file=sys.stderr)
            return 1
    if not deck.exists():
        print(f"Error: {deck} was not written", file=sys.stderr)
        return 1

    failed = False

    # Started before the checks, and collected after them: the deck is built,
    # so rendering it and checking it are independent, and overlapping them is
    # the difference between this command finishing inside a shell's timeout
    # and being cut off just before the images it exists to produce.
    render_proc = (None if args.no_render
                   else _start_render(deck, outdir, args.dpi))

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
        print("No overflow, collisions, leftovers, markdown, contrast or "
              "occlusion defects.")

    _write_report(deck, issues, checked, None if args.no_render else outdir)

    if deck.suffix.lower() in PPTX_EXTS and not args.skip_validate:
        code, output = validate_pptx(deck, args.template)
        print("\n## File integrity\n")
        if code == 0:
            print("validate.py: passed")
        else:
            failed = True
            print(f"validate.py: FAILED (exit {code})")
            print(output)

    if render_proc is not None:
        print("\n## Renders\n")
        images = _finish_render(render_proc)
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
        else:
            print(
                "\nNo renders were produced, so nothing here has been looked "
                "at. Fix that before calling the deck done -- run "
                "`python3 scripts/render.py <deck> --outdir <dir>` with a "
                "directory you can read."
            )

    report = deck.resolve().with_name(deck.stem + "-report.md")
    if report.exists():
        print(f"\nfindings: {report}")

    shutdown()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
