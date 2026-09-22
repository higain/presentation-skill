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
6. print a contact sheet of the whole deck,
7. leave a per-slide verdict table in the report, with the verdicts blank.

Then **open the PNGs and look at them.**  Steps 2 and 3 cannot see a slide;
they catch text that runs off the page and words nobody replaced, and they are
blind to a title colliding with its underline, a diagram drawn off-centre, or a
slide that is simply empty.  Every deck that shipped broken had passed a check
like this one and had never been looked at.

Exit codes: 0 clean, 1 the build failed or the checks found something, 2 bad
usage.
"""

import argparse
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from office.deckbuild import SKIP_QA_ENV, find_template  # noqa: E402
from office.impress import PRESENTATION_EXTS, SofficeError, shutdown  # noqa: E402
from qa import check, validate_pptx  # noqa: E402
from review_renders import MODEL_ENV as REVIEW_MODEL_ENV  # noqa: E402

PPTX_EXTS = {".pptx", ".potx"}


def _build_env() -> dict:
    """The environment a build script runs in.

    Two things it cannot arrange for itself:

    ``PYTHONPATH`` -- this skill's ``scripts/`` directory, so ``from
    office.draw import ...`` works whether or not the script opens with the
    ``sys.path.insert`` incantation.  We know where the skill is; making every
    build script rediscover it produced a steady supply of
    ``ModuleNotFoundError: No module named 'office'``.

    ``PRESENTATION_SKIP_BUILD_QA`` -- ``build_deck`` QAs its own output, and
    this command runs the identical check on the identical file a moment later.
    On a fifteen-slide deck that second pass is about a minute of pure
    duplication, on the one command the docs tell you to run.
    """
    env = dict(os.environ)
    env["PYTHONPATH"] = (str(HERE) + os.pathsep + env["PYTHONPATH"]
                         if env.get("PYTHONPATH") else str(HERE))
    env[SKIP_QA_ENV] = "1"
    return env


def _run_script(script: Path) -> int:
    """Run a build script in its own interpreter, streaming its output."""
    print(f"# Building — {script.name}\n", flush=True)
    proc = subprocess.run([sys.executable, str(script)], cwd=script.parent,
                          env=_build_env())
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


def _verdict_table(checked: int, sparse=()) -> list:
    """The per-slide verdict table: one row per slide, the last column blank.

    The automated half of this command writes findings, exits 1, and has to
    come back clean.  The visual half wrote nothing and gated nothing, and it
    is the half that failed: of four runs measured against the same plan, one
    opened every render and *listed* three real defects in its closing message
    rather than fixing them, one opened forty-five renders and called a deck
    with five identical filler slides finished, and one dismissed a correct
    finding it could not see.  All three reports said clean, because a report
    that asks "did you look?" is answered by whoever is already sure they did.

    An empty column is a different kind of question.  It is visible in the
    file, it survives the transcript, and filling it in is an act.
    """
    measured = dict(sparse)
    rows = ["| Slide | Measured | Verdict — `ok`, or what you will fix |",
            "|---|---|---|"]
    for number in range(1, checked + 1):
        rows.append(f"| {number} | {measured.get(number, '')} | |")
    return rows


def _wireframes(outdir: Path, numbers) -> list:
    """Text ink-maps for ``numbers``, or ``[]`` if none can be drawn.

    Included for the slides already flagged as worth a look, because those are
    the ones where "is this empty or is it a statement slide" is the question,
    and that is a question the grid answers on its own.  For a reviewer with no
    eyes it is the whole of the visual evidence; for one with eyes it is a
    second opinion that costs a tenth of a second.
    """
    if not numbers:
        return []
    try:
        from wireframe import block, find_renders, slide_number
    except Exception:                         # no PIL: the grid is optional
        return []
    blocks = []
    for png in find_renders(outdir, set(numbers)):
        try:
            blocks.append(block(png, number=slide_number(png)))
        except Exception:
            continue
    if not blocks:
        return []
    return ["", "## The slides to look at, as text", "",
            "Each cell is the mean ink of that patch of the slide, ` ` to `@`. "
            "It shows emptiness, banding and things that are not centred — not "
            "whether a slide is any good. `scripts/wireframe.py <png>` draws "
            "one for any render.", ""] + blocks


def _write_report(deck: Path, issues: list, checked: int, outdir=None,
                  sheet=None, sparse=(), review=None, looked=False,
                  notes=(), prose=None) -> Path:
    """Leave the findings on disk, next to the deck, as ``<deck>-report.md``.

    A build that ends between "here is what is wrong" and "fixed it" otherwise
    leaves a deck that looks finished and a set of findings that exist only in
    a transcript nobody will re-read.  On disk, the next thing to do is
    obvious, and a clean run overwrites it with a clean report.

    The renders are named here too, because this whole command takes a couple
    of minutes and a shell that gives up at two will have thrown away the paths
    it printed.  One ``cat`` of this file gets them back.

    Written twice: once before the renders finish, so a run killed on a short
    timeout still leaves its findings, and again once the contact sheet exists
    and can be named exactly.  ``sheet`` used to be guessed as
    ``contact-sheet-1.jpg``, which is what ``thumbnail.py`` writes only when a
    deck needs more than one page; a ten-slide deck gets ``contact-sheet.jpg``,
    and the first instruction in its own report was a path that did not exist.
    """
    report = deck.resolve().with_name(deck.stem + "-report.md")
    # Line two, not line thirty: a reader who stops after the heading should
    # already know whether anything has looked at this deck.  It used to be the
    # last line of the file, under the render paths, and it was read as a
    # footnote by every run that shipped a deck nobody had opened.
    eyes = ("a vision model (advisory)" if looked
            else "nothing — no render has been looked at yet")
    lines = [
        f"# {deck.name}", "",
        f"**Eyes:** {eyes}", "",
        f"{checked} slide(s) checked.", "",
    ]
    for note in notes:
        lines += [note, ""]
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
    if sparse:
        lines += ["", f"## {len(sparse)} slide(s) to look at", "",
                  "Measured, but not defects: a statement slide is sparse on "
                  "purpose and a forgotten one measures the same. Open these "
                  "first and decide.", ""]
        lines += [f"- [ ] slide {number}: {reason}" for number, reason in sparse]
    # Before the renders and before the verdicts, because it is read fastest
    # and it is the one part of this report that can still be acted on cheaply:
    # a storyline fixed here costs one edit to plan.md, and after the slides
    # are built it costs the deck.
    prose_issues, titles = prose if prose else ((), ())
    if titles:
        lines += ["", "## The ghost deck", "",
                  "Read these, in order, and nothing else. They should tell "
                  "the whole argument on their own — each a claim that follows "
                  "from the last. If they do not, the storyline is what needs "
                  "fixing, not the slides (`references/storyline.md`).", ""]
        lines += [f"{number}. {title if title else '*(no title found)*'}"
                  for number, title in titles]
    if prose_issues:
        lines += ["", f"## {len(prose_issues)} line(s) to look at", "",
                  "How the deck reads, not what is broken in it: a banned word "
                  "can have an honest use and a short title is right on a "
                  "divider. See `references/language.md`.", ""]
        lines += [f"- [ ] {finding}" for finding in prose_issues]
    if review:
        lines += ["", review, "",
                  "**A finding here is dismissed only by a measurement you "
                  "show.** Reasoning about why the reviewer must be wrong is "
                  "how a German footer shipped on every slide of an English "
                  "deck: the run searched the shapes, found nothing, and wrote "
                  "the finding off as a hallucination — the footer is a page "
                  "property, and the search could never have reached it. "
                  "`probe.py`, `qa.py` and `wireframe.py` settle most of these; "
                  "if none of them can, the finding stands and the fix is "
                  "cheaper than the argument."]
    if outdir is not None:
        lines += ["", "## Renders", "",
                  f"- [ ] Look at every slide in `{outdir}`"]
        if sheet is not None:
            lines += [f"- [ ] Start with the contact sheet, `{sheet}`"]
        lines += ["",
                  "A vision model looked at these. You still have not."
                  if looked else
                  "Nothing above was seen by anything with eyes.", "",
                  "## Verdicts", "",
                  "One row per slide. **The deck is not finished while a "
                  "verdict is blank**, and a slide nobody has opened cannot "
                  "have one. Write them in this file — that is the record that "
                  "the deck was looked at.", ""]
        lines += _verdict_table(checked, sparse)
        lines += _wireframes(outdir, [number for number, _ in sparse])
    try:
        report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        return None
    return report


def _review(outdir: Path, timeout=600):
    """Ask a vision model to look at the renders.

    Returns ``(markdown, looked)``.  ``looked`` is the honest bit: a reviewer
    that timed out or refused still produces text to print, and reporting that
    as "a vision model has looked at these" is the same lie as a clean report
    on a deck nobody opened.  ``(None, False)`` when none is configured.

    Runs whenever ``$PRESENTATION_REVIEW_MODEL`` names one, so looking at the
    slides is a property of the environment rather than a step the build agent
    has to remember.  That is the whole point: the run this was written for was
    told to open the PNGs, was handed their paths, and never opened one.

    Never raises and never fails the build -- see ``review_renders.py``.  A
    reviewer that is down, slow or talking nonsense must not cost you a deck
    that is otherwise finished.
    """
    if not os.environ.get(REVIEW_MODEL_ENV):
        return None, False
    try:
        proc = subprocess.run(
            [sys.executable, str(HERE / "review_renders.py"), str(outdir)],
            capture_output=True, text=True, timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        return (f"## Visual review timed out after {timeout}s\n\n"
                f"Nobody looked at these slides."), False
    except Exception as e:                       # missing script, bad interpreter
        return f"## Visual review could not run\n\n{type(e).__name__}: {e}", False
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout).strip()[:400] or "no output"
        return (f"## Visual review could not run\n\n{detail}\n\n"
                f"Nobody looked at these slides."), False
    text = proc.stdout.strip()
    return (text, True) if text else (None, False)


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
    # Through the same resolver ``build_deck`` uses, so the one path string the
    # documentation prints for both -- "assets/company-template.pptx" -- means
    # the same file here as it does in the build script.
    if args.template:
        try:
            args.template = str(find_template(args.template))
        except FileNotFoundError as e:
            print(f"Error: {e}", file=sys.stderr)
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
        try:
            wanted = qa._parse_slides(args.slides)
        except ValueError as e:
            print(f"Error: --slides {e}", file=sys.stderr)
            if render_proc is not None:
                render_proc.kill()
            return 2
    try:
        issues, checked, sparse, notes, prose = check(
            deck, args.template, wanted)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        if render_proc is not None:
            # Started before the checks; nobody will read its pipes now.
            render_proc.kill()
            render_proc.communicate()
        return 2

    print(f"{checked} slide(s) checked.\n")
    for note in notes:
        print(f"{note}\n")
    if issues:
        failed = True
        print(f"## {len(issues)} issue(s) — fix every one\n")
        for issue in issues:
            print(f"- {issue}")
    else:
        print("No overflow, collisions, leftovers, markdown, contrast or "
              "occlusion defects.")

    # Deliberately not `failed`: see qa.sparse_slides.  A statement slide is
    # sparse on purpose, and a check that cannot be satisfied is a check that
    # gets ignored -- along with the ones above it that can.
    if sparse:
        print(f"\n## {len(sparse)} slide(s) to look at\n")
        for number, reason in sparse:
            print(f"- slide {number}: {reason}")

    prose_issues, titles = prose
    if titles:
        print("\n## The ghost deck\n")
        for number, title in titles:
            print(f"{number:>3}. {title if title else '(no title found)'}")
        print("\nRead these in order. They should tell the whole argument on "
              "their own.")
    if prose_issues:
        print(f"\n## {len(prose_issues)} line(s) to look at\n")
        for finding in prose_issues:
            print(f"- {finding}")

    _write_report(deck, issues, checked, None if args.no_render else outdir,
                  sparse=sparse, notes=notes, prose=prose)

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

        review, looked = _review(outdir) if images else (None, False)
        if review:
            print(f"\n{review}")
        elif images and not os.environ.get(REVIEW_MODEL_ENV):
            print(f"\n(No visual review: set ${REVIEW_MODEL_ENV} to a model "
                  f"that accepts images — a local one served by ollama or "
                  f"vLLM, or a hosted one if the deck may leave the machine — "
                  f"and every run gets one. Nothing here has been seen.)")

        # Rewritten now that the sheet has a name -- the early write above had
        # to guess at one, and guessed wrong on every deck short enough to fit
        # on a single sheet.
        _write_report(deck, issues, checked, outdir,
                      sheet.resolve() if sheet else None, sparse, review,
                      looked, notes, prose)
        if images:
            print(
                "\n**Open these and look at them.** "
                + ("A vision model has, and said what it saw above — that is a "
                   "second opinion, not a substitute for yours."
                   if looked else
                   "The checks above cannot see a slide: they will not tell "
                   "you that a diagram is off-centre, that two shapes collide, "
                   "or that one is still in the template's language.")
                + (" Start with slide%s %s — measured above as possibly empty."
                   % ("" if len(sparse) == 1 else "s",
                      ", ".join(str(n) for n, _ in sparse)) if sparse else "")
                + " Then write a verdict per slide in the report's Verdicts "
                  "table; a blank one is a slide nobody has looked at."
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
