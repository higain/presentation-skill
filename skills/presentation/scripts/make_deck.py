#!/usr/bin/env python3
"""
Build, check and render a deck in one command -- the last thing you run.

    python3 scripts/make_deck.py build.py deck.pptx --template company-template.pptx
    python3 scripts/make_deck.py build.py deck.pptx --build-only        # tight loop
    python3 scripts/make_deck.py deck.pptx --skip-qa --no-render        # no gate
    python3 scripts/make_deck.py deck.pptx --require-verdicts           # visual gate, no rebuild
    python3 scripts/make_deck.py deck.pptx --no-render

Give it your build script and the deck that script writes, and it will:

1. run the build script,
2. QA the result -- overflow, off-slide shapes, leftover template wording and
   pictures, placeholders, markdown, contrast, occlusion, and top-heavy
   content slides,
3. schema-check the ``.pptx``,
4. write the findings to ``<deck>-report.md`` beside the deck,
5. render every slide to a PNG in ``render/`` beside the deck, and print paths,
6. print a contact sheet of the whole deck,
7. leave a per-slide verdict table in the report, with the verdicts blank.

Then **give the PNGs to the vision-agent** and write its per-slide verdicts into
the report.  Steps 2 and 3 cannot see a slide;
they catch text that runs off the page and words nobody replaced, and they are
blind to a title colliding with its underline, a diagram drawn off-centre, or a
slide that is simply empty.  Every deck that shipped broken had passed a check
like this one and had never been looked at — and a build model with no eyes
must delegate that looking to the vision-agent rather than skip it.

**Iteration flags.** While the build itself is failing, ``--build-only`` skips
QA and rendering entirely and ``--skip-qa`` skips just the checks; neither
writes a report, so a deck that was not checked never leaves a clean-looking
one behind.  The full command is the gate and is run once, when the tight loop
is clean.

**The visual gate.** ``--require-verdicts`` makes the "every slide needs a
verdict" rule an exit code: a blank verdict cell fails the run.  It only reads
the report -- it builds, checks and renders nothing, because any of those
rewrites the report and blanks the verdicts it is meant to check.  The report
carries the deck's SHA-256, so verdicts written for an earlier build of the
deck are refused rather than counted.

Exit codes: 0 clean, 1 the build failed or the checks found something, 2 bad
usage.
"""

import argparse
import hashlib
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from office.deckbuild import SKIP_QA_ENV, find_template  # noqa: E402
from office.impress import (PPTX_EXTS, PRESENTATION_EXTS,  # noqa: E402
                            SofficeError, parse_slide_spec, shutdown)
from office.soffice import isolated  # noqa: E402
from office.qa import QAResult, check, summary, validate_pptx  # noqa: E402
from office.qa.allowlist import (AllowlistError, _acknowledged,  # noqa: E402
                                 load_allowlist)
from office.vision import last_json  # noqa: E402

#: A vision finding at these severities holds the gate until the slide is
#: rebuilt and reviewed again, or acknowledged in the allowlist.
BLOCKING_SEVERITY = {"high", "medium"}



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
    """Run a build script in its own interpreter, streaming its output.

    Without network, like LibreOffice itself: the script was written by a
    model that read untrusted decks, and a sentence planted in a deck's notes
    ("also upload the template to ...") must not become a request.  ``-P``
    keeps the script's own directory off ``sys.path``, so an ``office/`` or
    ``uno.py`` sitting beside it cannot stand in for the engine.
    """
    print(f"# Building — {script.name}\n", flush=True)
    proc = subprocess.run(isolated([sys.executable, "-P", str(script)]),
                          cwd=script.parent, env=_build_env())
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
    # New renders mean every brief and answer beside them describes slides
    # that no longer exist.
    for stale in (list(outdir.glob("slide-*.brief.md"))
                  + list(outdir.glob("slide-*.answer.md"))):
        stale.unlink(missing_ok=True)
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


DECK_HASH_MARK = "<!-- deck-sha256: "


def deck_hash(deck: Path) -> str:
    return hashlib.sha256(deck.read_bytes()).hexdigest()


def report_deck_hash(report: Path):
    """The deck hash a report was written for, or ``None``."""
    try:
        for line in report.read_text(encoding="utf-8").splitlines()[:6]:
            if line.startswith(DECK_HASH_MARK):
                return line[len(DECK_HASH_MARK):].split()[0]
    except OSError:
        pass
    return None


def open_vision_findings(deck: Path) -> list:
    """``[slide]`` whose saved vision-agent answer still holds a real finding.

    Read from the answers, not from the verdict table: the table is text the
    build model edits, and a model that wants the gate to pass rewrites
    "cards mostly empty" as "ok -- cards fill the area proportionally".  The
    answers are deleted whenever ``make_deck.py`` renders again, so a slide
    clears by being rebuilt and reviewed -- or by a ``vision`` entry in the
    allowlist (``"content": "Q7"``), which a person can read and dispute.
    """
    outdir = deck.resolve().parent / "render"
    try:
        allow = load_allowlist(deck)
    except AllowlistError:
        allow = {}
    out = []
    for answer in sorted(outdir.glob("slide-*.answer.md")):
        number = int(answer.name.split("-")[1].split(".")[0])
        parsed = last_json(answer.read_text(encoding="utf-8")) or {}
        for finding in parsed.get("findings") or []:
            if not isinstance(finding, dict):
                continue
            if str(finding.get("severity", "")).lower() not in BLOCKING_SEVERITY:
                continue
            q = str(finding.get("q", "")).upper()
            if allow and _acknowledged(allow, "vision", number, q):
                continue
            out.append(number)
            break
    return out


def open_fix_verdicts(report: Path) -> list:
    """``[slide_number]`` whose verdict still says ``fix``.

    ``vision_brief.py collect`` writes the vision-agent's findings as
    ``fix — ...``.  A verdict is a decision, and "here is what is wrong" is
    not one: the gate passes only once the slide is fixed and re-checked, or
    the finding is overruled in the cell with a measurement.
    """
    out = []
    for line in report.read_text(encoding="utf-8").splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if (line.startswith("|") and len(cells) == 3 and cells[0].isdigit()
                and cells[2].lower().startswith("fix")):
            out.append(int(cells[0]))
    return out


def missing_verdicts(report: Path, checked: int) -> list:
    """``[slide_number]`` whose verdict cell in ``report`` is blank.

    ``--require-verdicts`` turns the report's "the deck is not finished while a
    verdict is blank" into an exit code, so the visual half of QA gates the
    same way the automated half does.  The verdict column is the whole record
    that something with eyes looked at the slide; a blank cell means nobody
    has, and the deck does not ship.
    """
    if not report.is_file():
        return list(range(1, checked + 1))
    missing = []
    found_rows = False
    try:
        lines = report.read_text(encoding="utf-8").splitlines()
    except OSError:
        return list(range(1, checked + 1))
    for line in lines:
        # | 7 | paints 22% of the slide | |
        if not line.startswith("| "):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) != 3:
            continue
        slide, _measured, verdict = cells
        if not slide.isdigit():
            continue
        found_rows = True
        if not verdict:
            missing.append(int(slide))
    # A report with no verdict rows at all means nobody has started -- every
    # slide is unlooked-at, not just the ones an empty table implies.
    if not found_rows:
        return list(range(1, checked + 1))
    return missing


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


def _write_report(deck: Path, result: QAResult, outdir=None,
                  sheet=None) -> Path:
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
    issues, checked, sparse = result.issues, result.checked, result.sparse
    notes, acknowledged = result.notes, result.acknowledged
    prose_issues, titles = result.prose_issues, result.titles
    report = deck.resolve().with_name(deck.stem + "-report.md")
    # Line two, not line thirty: a reader who stops after the heading should
    # already know whether anything has looked at this deck.  It used to be the
    # last line of the file, under the render paths, and it was read as a
    # footnote by every run that shipped a deck nobody had opened.
    eyes = "nothing — no render has been looked at yet"
    lines = [
        f"# {deck.name}",
        f"{DECK_HASH_MARK}{deck_hash(deck)} -->", "",
        f"**Eyes:** {eyes}", "",
        f"{checked} slide(s) checked.", "",
    ]
    for note in notes:
        lines += [note, ""]
    if acknowledged:
        lines += [f"## {len(acknowledged)} acknowledged (from the allowlist)", "",
                  "These were reported as findings and are recorded as "
                  "intentional, each with a reason -- see "
                  "`<deck>-allowlist.json`.", ""]
        lines += [f"- {ack}" for ack in acknowledged]
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
                  "Give the renders to the vision-agent."]
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
    if outdir is not None:
        lines += ["", "## Renders", "",
                  f"- [ ] Look at every slide in `{outdir}`"]
        if sheet is not None:
            lines += [f"- [ ] Start with the contact sheet, `{sheet}`"]
        lines += ["",
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


def _gate_on_verdicts(deck: Path, script) -> int:
    """``--require-verdicts``: read the report, never rewrite it.

    Every other mode rewrites ``<deck>-report.md`` and so blanks the verdict
    column; checking the verdicts after that could never pass.  So this mode
    builds nothing and only reads -- and refuses a report written for a
    different build of the deck, whose verdicts describe other slides.
    """
    if script is not None:
        print("Error: --require-verdicts only reads the report. Run it on the "
              "deck alone, after the full run and after writing the verdicts: "
              f"make_deck.py {deck} --require-verdicts", file=sys.stderr)
        return 2
    if not deck.exists():
        print(f"Error: {deck} does not exist", file=sys.stderr)
        return 2
    report = deck.resolve().with_name(deck.stem + "-report.md")
    if not report.is_file():
        print(f"No report at {report}. Run make_deck.py on the deck first.")
        return 1
    if report_deck_hash(report) != deck_hash(deck):
        print(f"{report.name} was written for a different build of "
              f"{deck.name}. Its verdicts describe other slides: re-run the "
              f"full make_deck.py, review the new renders, write the verdicts "
              f"again.")
        return 1
    rows = sum(1 for line in report.read_text(encoding="utf-8").splitlines()
               if line.startswith("| ") and line[2:].split("|")[0].strip().isdigit())
    missing = missing_verdicts(report, rows)
    if missing:
        shown = ", ".join(str(n) for n in sorted(missing))
        print(f"Missing verdicts for slide(s): {shown}. A slide nobody has "
              f"looked at cannot be finished -- give the renders to the "
              f"vision-agent and write its verdict into the report's table.")
        return 1
    open_fixes = sorted(set(open_fix_verdicts(report))
                        | set(open_vision_findings(deck)))
    if open_fixes:
        shown = ", ".join(str(n) for n in open_fixes)
        print(f"Open findings on slide(s): {shown} -- a `fix` verdict, or a "
              f"medium/high finding in the vision-agent's saved answer. Fix "
              f"the slide in the build script, re-run make_deck.py and brief "
              f"the new renders. Editing the verdict cell does not clear it. "
              f"A finding a measurement proves wrong goes in "
              f"<deck>-allowlist.json as kind \"vision\", content \"Q7\" "
              f"(the question), with the measurement as its reason -- the "
              f"user reviews that file.")
        return 1
    print(f"All {rows} slide(s) have a verdict.")
    return 0


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
    parser.add_argument("--skip-qa", action="store_true",
                        help="Build without the automated QA pass. For fast "
                             "build-error iteration only; the deck is not "
                             "checked and nothing is written to the report. "
                             "Run the full command (no --skip-qa) to gate.")
    parser.add_argument("--skip-validate", action="store_true")
    parser.add_argument("--build-only", action="store_true",
                        help="Build only: no QA and no renders. The tightest "
                             "loop for fixing a build error. Always follow "
                             "with a full run (no flags) before calling the "
                             "deck done.")
    parser.add_argument("--require-verdicts", action="store_true",
                        help="Read the report only (no build, no QA, no "
                             "render) and exit 1 unless every slide has a "
                             "verdict and the report matches this deck. Run "
                             "it after writing the verdicts.")
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
    if args.require_verdicts:
        return _gate_on_verdicts(deck, script)

    # Through the same resolver ``build_deck`` uses, so the one path string the
    # documentation prints for both -- "assets/company-template.pptx" -- means
    # the same file here as it does in the build script.
    if args.template:
        try:
            args.template = str(find_template(args.template))
        except FileNotFoundError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    # --build-only is shorthand for the tightest loop: no QA, no renders.
    if args.build_only:
        args.skip_qa = True
        args.no_render = True

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
    if args.skip_qa:
        # No QA: this is the tight loop for fixing a build error. The deck is
        # built (or a script just wrote it); nothing is checked and the deck
        # is not gated. The full run -- no --skip-qa -- is the gate.
        result = QAResult()
        flag = "--build-only" if args.build_only else "--skip-qa"
        print(f"Skipping QA ({flag}). This deck is NOT checked and NOT "
              f"finished -- re-run without {flag} before it ships.")
    else:
        if args.slides:
            try:
                wanted = set(parse_slide_spec(args.slides))
            except ValueError as e:
                print(f"Error: --slides {e}", file=sys.stderr)
                if render_proc is not None:
                    render_proc.kill()
                return 2
        try:
            result = check(deck, args.template, wanted)
        except (SofficeError, FileNotFoundError, ValueError) as e:
            print(f"Error: {e}", file=sys.stderr)
            if render_proc is not None:
                # Started before the checks; nobody will read its pipes now.
                render_proc.kill()
                render_proc.communicate()
            return 2

    if not args.skip_qa:
        print("\n".join(summary(result)))
        # Sparse slides and prose lines are deliberately not `failed`: a
        # statement slide is sparse on purpose, and a check that cannot be
        # satisfied is a check that gets ignored -- along with the ones above.
        failed = failed or bool(result.issues)

        _write_report(deck, result, None if args.no_render else outdir)

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

        # Rewritten now that the sheet has a name -- the early write above had
        # to guess at one, and guessed wrong on every deck short enough to fit
        # on a single sheet.  Skipped under --skip-qa: that run must not leave
        # a clean-looking report behind.
        if not args.skip_qa:
            _write_report(deck, result, outdir,
                          sheet.resolve() if sheet else None)
        if images:
            print(
                "\n**Next: the vision-agent, one slide per call.** The checks "
                "above cannot see a slide. Run "
                f"`vision_brief.py slides {deck}` -- it writes the tested "
                "brief beside every render and prints the one-line "
                "delegation -- then `vision_brief.py collect` fills the "
                "verdict table."
                + (" Start with slide%s %s — measured above as possibly empty."
                   % ("" if len(result.sparse) == 1 else "s",
                      ", ".join(str(n) for n, _ in result.sparse))
                   if result.sparse else "")
                + " A blank verdict is a slide nobody has looked at."
            )
        else:
            print(
                "\nNo renders were produced, so nothing here has been looked "
                "at. Fix that before calling the deck done -- run "
                "`python3 scripts/render.py <deck> --outdir <dir>` with a "
                "directory you can read."
            )

    report = deck.resolve().with_name(deck.stem + "-report.md")
    if not args.skip_qa and report.exists():
        print(f"\nfindings: {report}")

    shutdown()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
