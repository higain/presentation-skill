# Big decks: one file per slide, built by several agents

**Use this when** the deck has more than about 20 slides, or its source — a
slide-by-slide specification, a long document — is over about 50 KB
(`wc -c source.md`). Below that, the single `build.py` of `SKILL.md` is
simpler.

Why: a 58-slide specification was 288 KB, half of a 200k-token context
before a line of code. The agent that read it whole lost it again to context
pruning and re-read pieces at random. It then parsed the specification with
regular expressions, and one pattern silently left a field empty on all 57
content slides. Its single build of 58 slides stopped at the first broken
slide, once per run. The fix is to give every agent only the part it works
on:

- the **source** is split into one file per slide;
- the **deck** is one Python file per slide, `slides/s07.py`, written by
  builder subagents that each read only their own slides;
- **one** `build_parts` call assembles them in one LibreOffice session.

Never build slides as separate decks and merge the files: a slide cannot be
moved between documents over UNO, only rebuilt shape by shape, and that loses
the master, the placeholders and the notes.

## You, the orchestrator

**1. Split the source. Do not read it whole.**

```bash
scripts/split_source.py source.md --out spec/
```

It writes `spec/slide-NN.md` (each slide's section, verbatim),
`spec/_shared.md` (everything that belongs to no single slide: the brief,
design rules, an overview table) and `spec/index.md` (number, title, size).
Read `index.md` and `_shared.md`; open a slide file only when you need it.
If it picks the wrong headings, pass `--heading '^### Folie (\d+)'`.

**2. Decide everything that must be the same on every slide, once.** Read the
pack's `catalog.md`, then decide which design each kind of slide uses, how
**every** colour the source uses maps to (its tints and dark text colours
too, not only the main roles), which icons stand in for the source's, how the
headline and kicker are written, the pack's smallest type size, and where
elements go that the source puts in the template's title band or footer (a
tag beside the kicker, a page badge). A builder who meets an undecided case
decides it alone, differently on every batch. Write the decisions into **`slides/_common.py`** as
named constants and small helpers. Start with a **slide frame** that every
content slide calls: kicker, title, phase tracker, the canvas area,
`frame(page, kicker, title, phase=2) -> area`. Then the furniture is
identical on every slide by construction. Add the other decisions beside it
(`tag(doc, page, box, kind)`, `ROLE = {"risk": D.color("alert"), ...}`). Every
part imports it as `import deck_common as C`. Builders reuse a helper; they
do not re-decide it.

**A specification that gives positions and sizes** (`[x 10 %, y 23 %, B 58 %]`,
"14 pt") describes the original's *arrangement*, not coordinates for this
template. Decide here which helper draws each recurring arrangement: a row of
cards is `D.cards`, a flow is `D.steps`, numbers are `D.kpis`, a grid of rows
is `D.table`. Put that mapping in `_common.py`. **Never put page-percentage
helpers (`pct_x(sw, 28)`) in it**: a field run did, its builders placed 570
shapes by percentage, and 18 slides printed text over text. QA now fails
hand-placed shapes that almost line up (`references/qa.md` §Hand-placed
slides). Type
sizes come from the pack's roles (`role="body"`, `role="caption"`), never
from the specification's point sizes: a spec's 10 pt kicker is under the
pack's floor and fails the gate.

**3. Write `build.py` and the first one or two slides yourself.**

```python
from office.deckbuild import build_parts
build_parts("acme", "deck.pptx", "slides", footer="Acme · Workshop · 2026")
```

```python
# slides/s02.py
from office import draw as D
from office.template import set_title
from office.impress import mm, set_notes
import deck_common as C

DESIGN = 14                     # the template slide this one starts from

def fill(doc, page):
    area = C.canvas(page)
    set_title(page, "Three levers secure the March launch")
    D.text(doc, page, area.top(mm(20)),
           D.para("Agenda", weight=D.BOLD), "Risks, then guardrails",
           size=16)
    set_notes(page, "Speaker notes, verbatim from the source.")
```

`page.Visible = False` hides a slide. Check your examples:

```bash
scripts/make_deck.py build.py deck.pptx --parts 1-2 --template acme
```

`--parts` builds only those slides into `deck-parts-1_2.pptx` beside the
deck and checks them. The deck itself is never touched. Always pass
`--template`: without it the check cannot tell the template's own shapes
from yours, and reports white text on a blue band as white on white. Fix the helpers until
these slides are clean: from here on every builder copies them.

Keep this step short: two or three examples, not one of every kind. A finding
you think the check gets wrong for this template is not yours to debug: do
not read the engine's source or test its measurements. Note the finding and
why you doubt it, and move on; the user decides at the end. A field run spent
forty minutes of an hour probing the QA code from here.

**4. Brief the builders.** Use a general-purpose subagent (opencode:
`general`; Claude Code: `general-purpose`), not the vision-agent. Give each
one 8 to 10 consecutive slides, and **run at most three at a time**, fewer
if the user says so. Start the next builder when one reports, not before.
Every builder runs its own LibreOffice and holds a connection to the model
server: a field run started ten at once on a 4-core, 8 GB machine, and the
machine went down with it. The skill also lets only a few LibreOffice runs
go at once on a machine (`PRESENTATION_MAX_OFFICE`, by default half the
cores, one per 3 GB of RAM); a build over that limit prints `[office]
waiting` and waits its turn. More builders than that do not go faster.
For a shape-heavy deck, prefer short sessions with `--parts` over one long
continuous run: the engine's intermittent crashes grow likelier the longer a
single session runs (SKILL.md step 5).

Send the brief below with the placeholders filled in. Keep their reports,
not their code: you do not need to read the slide files.

**5. Assemble and gate.** Run `scripts/make_deck.py build.py deck.pptx
--template acme --source <the source document>`, the full gate of `SKILL.md`
step 5. `--source` checks that every text and speaker note of the source is
on the slides as written, and nothing is invented (`references/qa.md`
§Fidelity). If a part raises, the
build reports **every** failing slide with its file and traceback, not just
the first. Then do the vision review of step 6, one slide per call, as
always. For each finding, send the builder of that slide (or a new one) the
finding, the slide's source file and its part file, or fix it yourself if
it is one line.

If the assembly dies with `Binary URP bridge disposed` or `soffice exited
abnormally`, that is soffice crashing, not a slide: read the
`soffice-stderr-<profile>.log` the run leaves and re-run once (probabilistic);
die again at a different slide, build in halves with `--parts`; after three
different-slide failures, tell the user before retrying (SKILL.md step 5).

## The builder brief

```text
You write slides {FIRST}–{LAST} of a presentation, one Python file per slide.
Skill directory: {SKILL}. Deck directory: {DECK_DIR}.

Read in full before writing:
- {SKILL}/references/api.md — every function you may call, with signatures
- {SKILL}/references/drawing.md — how to lay out with Box and the helpers
- {DECK_DIR}/slides/_common.py — the shared helpers and decisions; use them
- {DECK_DIR}/slides/{EXAMPLE}.py — a finished slide in this deck's style
- {DECK_DIR}/spec/_shared.md — the rules that hold for every slide
- for each of your slides N: {DECK_DIR}/spec/slide-N.md
- the design files your slides use: {PACK}/designs/design-{D}.md

For each slide N write {DECK_DIR}/slides/sNN.py with `DESIGN = <number>`
and `def fill(doc, page):`. Transcribe the source by reading it: every text
verbatim (quotes, dashes, line breaks), every element, the speaker notes with
office.impress.set_notes. Never write code that parses the source files.

Lay every slide out from the Box that D.canvas returns, with the composition
helpers (D.cards, D.steps, D.kpis, D.table, D.bullets, D.statement) and the
helpers in _common.py. The source's percentages and point sizes describe the
arrangement, not coordinates: never compute a position from a page percentage
and never copy a point size; take sizes from role=. A diagram no helper draws
is laid out from one Box (grid, cut_*): the check fails shapes that should
line up and miss by a millimetre or two.

Decisions already made (do not re-decide): {DECISIONS}

Where the source has no content (a placeholder, "TBD", an empty field), put
"[OFFEN: what is missing]" on the slide. Never invent content.
Do not edit _common.py, build.py or other slides' files. If you need a helper
that is missing, write it in your own file and name it in your report.

Check your slides until the command exits 0, with a timeout of at least
five minutes. If it dies with `Binary URP bridge disposed` or `soffice exited
abnormally`, that is soffice crashing, not a slide: read the
`soffice-stderr-<profile>.log` it leaves and run the same command again
(SKILL.md step 5):
  python3 {SKILL}/scripts/make_deck.py {DECK_DIR}/build.py {DECK_DIR}/{DECK} --parts {FIRST}-{LAST} --template {PACK_NAME}
Every finding it prints for your slides is yours to fix. You cannot see
images: do not describe how a slide looks.

Report in at most 10 lines: the files written, any decision you had to make
yourself, every [OFFEN] and why, and helpers you wish were in _common.py.
Do not paste code.
```

## Keeping your own context small

- A 58-slide build prints a lot. Send it to a file and keep the exit code,
  `... > build.log 2>&1; echo "exit $?"`, then read the end of `build.log`.
  Never pipe it into `grep`, `head` or `tail`: the pipe reports exit 0 for a
  build that failed, and a filter drops the error.
- Look up a function in `references/api.md` or with `scripts/api_ref.py draw`
  instead of reading the skill's source with `sed`.
- Never wrap a slide's code in `try/except: pass` to make a build go through.
  The error is the information. `build_parts` already reports all of them in
  one run.
