# Storyline: deciding what the deck says, before it is built

Loaded from SKILL.md whenever a deck is **created from a brief** — a sentence or a
paragraph of intent — rather than filled from text the user already wrote.

**Skip this file** when the user supplies an outline, a document to turn into slides, or
a deck to edit. They have already done this work and it is not yours to redo.

**The rule the file exists to enforce:** no slide is built until the titles alone tell
the argument. A deck with a weak storyline and beautiful slides fails; the reverse
rarely does.

Five stages, each writing one file beside the deck. Do not merge them and do not skip
ahead — the files are what let you work on one bounded thing at a time, and they are
what the user reviews.

```
brief ─▶ 1 INTAKE ─▶ 2 SPINE ─▶ 3 GHOST DECK ─▶ 4 LOOK ─▶ 5 BUILD
          brief.md    spine.md   plan.md         look.md   deck.pptx
                                 ▲ ask the user  ▲ ask the user
```

**Write in the language of the brief.** A German brief gets German slides, German
titles, and the template's German footer left in place. Do not translate the user's
subject matter into English on the way past.

---

## Stage 1 — Intake

Ask these six questions. **Exactly these, once, and then stop asking.**

1. Who is in the room, and what is their job?
2. What do you want them to **do** when it ends?
3. How long have you got?
4. What is the one number or fact you would keep if you could keep only one?
5. What will someone in the room push back on?
6. Is anything here confidential, or not yet verified?

Write `brief.md` with the six answers **verbatim**. Do not embellish them, do not
expand an answer into prose, and do not answer one on the user's behalf.

Question 2 is the load-bearing one: it picks the scaffold in stage 2. Question 5 exists
because the objection slide is what separates a deck from a summary and nobody
volunteers it. Question 4 buys evidence discipline without a research phase.

**Gate.** Questions 1, 2 and 3 must have answers, and each has a defined fallback, so a
user who stops replying never leaves the run improvising:

| | If it comes back empty |
|---|---|
| **Q1** audience | Ask once more. Then assume **the user's own team**, write it down as an assumption, and carry on. Nothing downstream branches on it except tone. |
| **Q2** the ask | Ask once more in plainer words — *"what happens differently on Monday?"* Then treat the deck as **informative**, say so, and use the **Report-out** scaffold. |
| **Q3** the slot | Ask once more. Then assume **20 minutes**, the commonest internal slot, and say which figure you used — the whole slide budget hangs off it. |

Answers to 4, 5 and 6 may be `unknown`; an unknown becomes a visible `[TBD]` on a slide,
never a guess. **Record every fallback you took in `brief.md`.** A deck built on three
assumed answers is fine; a deck built on three assumed answers nobody was told about is
how the wrong deck reaches the room.

**Never invent a number to fill question 4.** A slide reading `[TBD: direct vs.
transitive dependency count]` is useful and gets filled in before the meeting. A slide
carrying a plausible invented figure is the one that gets quoted.

**One more thing to settle here, because nothing later asks for it:** the footer line.
A corporate master usually renders a `subject | title | place | date` line along the
bottom of every content slide, in the template's own language. It is a page property no
`retext` reaches, and four runs shipped it untouched under a deck written in another
language. Take `subject · occasion · date` from the answers to 1 and 3 and record
it at the foot of `brief.md`, so stage 5 has it: `footer="Dependency risk · Board
briefing · 2026"`.

---

## Stage 2 — Spine

### Pick one scaffold

Match the answer to question 2 against this table, top to bottom, **first match wins.**
This is a classification, not a composition — pick, do not blend two.

| The ask is… | Scaffold | Beats |
|---|---|---|
| approve, fund, or authorise something | **Case** | situation · tension · options · recommendation · evidence · risk · cost of inaction · ask |
| adopt a practice, control, or standard | **Rollout** | where we stand · how it actually works · it has already happened · why today's answers don't cover it · tiered plan · owners and dates |
| understand what happened | **Report-out** | headline · scorecard · what worked · what didn't · the insight · next period · ask |
| think about something differently | **Keynote** | opening scene · conventional wisdom · the crack · the reframe · three evidence beats · implication · the charge |
| be able to do something | **Workshop** | the gap · the framework · the components · worked example · common traps · the exercise |
| oversee or govern | **Board** | state of play · what changed · versus plan · where we are exposed · deliberation item · ask · forward look |
| commit, sign, or buy | **Close** | recap in their words · proof stack · differentiator · the objection · terms · one next step |

If two rows match, take the one whose **ask** is more specific. If none matches, use
**Case** — it is the safest default and the commonest internal need.

**Rollout is the commonest internal engineering deck** — persuading peers to adopt a
measure — and the one most often built wrong, as a tour of features. Its distinguishing
beat is *"why today's answers don't cover it"*: a technical room has already thought of
the obvious objections, and a deck that skips them loses the room before the plan.

### Write the spine

`spine.md`, four fields, nothing more:

```
governing question: the single question this deck answers
governing thought:  the one-sentence answer — this becomes the summary slide's title
branches:           2-4 arguments that together prove it
ask:                the specific thing the room does next, with an owner if known
```

**Branch rules.** Two to four, never five. Mutually exclusive — no fact belongs under
two branches. Collectively sufficient — if an obvious counter-question ("what does it
cost?", "who does the work?") has no home under any branch, add one or fold it in. Each
branch is one body slide in a short deck, one section in a long one.

**Name branches short.** A branch that becomes a section divider is a divider title, and
a divider sets its title large, in a gap the master's decoration leaves — so the budget
is often **under 25 characters on one line**. Over it the title wraps into the
decoration and the words come out with a stroke through them, nothing collides, and
`qa.py` reports clean. The budget has to be respected here, in the spine, where it is
free. **The pack's `catalog.md` states the exact figure per design** under *Title
budget*; take it from there rather than from this paragraph.

**Slide budget: roughly one slide per *presented* minute**, minus the title and closing
slides — and the slot from question 3 is not all presented minutes.

| The slot is… | Presented | Budget |
|---|---|---|
| a talk, questions at the end | all of it | 20 min ≈ 15 content slides |
| a talk with discussion built in | about two thirds | 45 min ≈ 28 |
| a **working session** — the room does an exercise | about half, and the exercise needs 2–3 slides of its own | 60 min ≈ 18 presented + 2 exercise + the ask |
| a board or steering item | about half; the rest is deliberation | 30 min ≈ 12, plus an appendix |

Take the multiplier from how question 3 was answered, not from the number in it. Sixty
minutes read as sixty slides is the arithmetic that produces a deck nobody can get
through, and it happens because the slot sounds generous.

If the branches do not fill the budget, the deck is shorter. **Never pad** — a symmetric
deck with a bullet invented to balance a column is the failure this whole file is built
to prevent.

---

## Stage 3 — Ghost deck

`plan.md`. One block per slide, three fields, **nothing else**:

```
## 4
title:  Transitive dependencies outnumber what we review 34 to 1.
proof:  42 direct → 1,447 installed  [source: SBOM, build 2291]
visual: area-compare — small block left, large block right, one red node deep right
```

- `title` — an **action title**: a complete sentence stating the conclusion, ending in a
  period. `references/language.md` §1 is the specification and `qa.py` enforces the
  mechanical half. If you write a label here, this stage has failed and no amount of
  layout work later recovers it.
- `proof` — **one** fact, number, quote or comparison, with its source in brackets. Not
  three. If there is none, write `[TBD]` and leave it visible.
- `visual` — the **name of a pattern** from `references/accents.md`, not a description of
  a picture and not a design number.

No speaker notes, no design numbers, no colours at this stage. They are decided in
stages 4 and 5, and putting them here is how the argument gets lost inside the
production detail.

### Prepend and append automatically

Any deck of four or more content slides opens with:

- **A title slide.** Check the catalog's *Designs by role* table for a **cover**
  design. Many corporate templates have none — their dividers are one-line slides and a
  real title with a subtitle and a date does not fit that budget — and then the opener
  is built on the **canvas** design, with the governing thought as the subtitle. Do not
  leave the deck without one.
- **A summary slide, second.** Governing thought as the title, the branches as one line
  each, in the order the deck proves them.

and ends with:

- **the ask** — specific, with an owner and a date — for anything persuasive, or
- **the strongest single takeaway**, for anything informative.

**Never end on "Thank you" or "Questions?".** The last slide stays on screen through the
whole discussion; make it carry the ask. `qa.py` reports a deck that ends this way.

### The ghost deck test — this is the gate

Read **only the titles, in order.** They must:

- **flow** — each follows from the last and builds toward the ask
- **tell the whole story** — someone who sees nothing but the titles gets the argument
- **each make a point** — every one a claim, never a topic label

`qa.py` prints the titles as a block for exactly this reason: a machine cannot judge
whether they flow, but it can put them in front of you in the ten seconds it takes to
tell. If they do not flow, the fix is here in `plan.md`, not in the slides.

**Then show `plan.md` to the user and wait for approval.** Do not build slides before
they have it. Restructuring here costs one edit; restructuring after the deck is built
costs the deck, and this is the cheapest place in the whole process to be wrong.

---

## Stage 4 — Look

`references/accents.md`. In short: the scaffold suggests an accent pack, you render the
deck's own strongest title in three packs, show the user the three images, and ask which
one feels right — not which colours they like. Write the chosen pack to `look.md`.

**Also a user gate.** It costs three thumbnails and it is the only aesthetic decision in
the pipeline that a person makes rather than a model.

---

## Stage 5 — Build

The recipe in SKILL.md, unchanged. `plan.md` drives it:

| Field in `plan.md` | Where it lands |
|---|---|
| `title` | **verbatim** into the design's title key, via `retext` — see below, this is where builds fail |
| `proof` | the body, or the caption under the figure |
| `visual` | picks the design: the pattern's row in `accents.md` names one, or the canvas |
| the branches | section dividers **only** — never the title of a content slide |
| `brief.md`'s footer line | `build_deck(..., footer="…")` |

### The titles are what the user approved. Carry them; do not retype them.

**Every title in the built deck is a string copied out of `plan.md` unchanged.**
Not shortened to fit, not turned back into a topic, and never replaced by the name
of the branch the slide sits under.

This is the one step of the whole file that has actually been observed to fail, and
it fails silently. A run took an approved plan through stage 4 and then built this:

```python
# plan.md block 1 says:
#   title: Our current dependency visibility is a major blindspot.
retext(page, {
    TITLE_KEY: "Current Blindspots",     # ← the BRANCH name, not the title
    LABEL_KEY: "Visibility",             # ← a topic label
})
```

The approved sentence is nowhere in the deck. Eleven of sixteen slides carried a
branch name as their title, four of them the same one, and the ghost deck read as
four labels repeated instead of sixteen claims. Every other check passed — the
build was clean, `validate.py` passed, the renders looked fine. **A deck that has
lost its titles is indistinguishable from a good one by every measure except the
titles themselves.**

**The pattern that prevents it:** lift the titles into one dict at the top of the
build script, straight out of `plan.md`, and never write a title literal anywhere
below it.

```python
# plan.md, copied in once. Nothing below this retypes a title.
TITLES = {
    1: "Our current dependency visibility is a major blindspot.",
    2: "Transitive dependencies introduce unmanaged risk.",
    3: "We cannot apply the same security gate to every package.",
}

def fill(doc, page, slot, design):
    strip_annotations(page)
    if slot in TITLES:
        retext(page, {TITLE_KEY[design]: TITLES[slot]})   # the design's title key
    ...
```

One place to check against `plan.md`, one place to fix, and the block numbers stay
visible next to the sentences. A title typed a second time is a title being
rewritten, whatever the intention.

**Branches are dividers.** A branch is a *section*, so its name goes on a section
divider slide of its own — the divider designs the catalog's *Designs by role* table
names, within their title budget. The
slides under it keep their own action titles. A branch name in a content slide's
title slot is exactly the defect above, and one divider in sixteen slides is the
symptom.

**`qa.py` now measures the transfer.** It reads `plan.md` from beside the deck and
reports every approved title that reaches no slide as a **finding** — these fail a
build, unlike everything else about the wording — plus, as prose, a title that
arrived reworded and any headline used on more than one slide. Write the deck into
the same directory as `plan.md` or the check cannot run.

Pick a design per slide from the decision table in the pack's `catalog.md`,
first match wins, and use the canvas design stripped back for anything it has no row
for.
One `fill` per page, one `build_deck` call, then
`scripts/make_deck.py build.py out.pptx --template templates/acme/template.pptx` — and
open every PNG it prints.

**A clean QA report does not mean the deck argues.** It means nothing was mechanically
wrong. The argument was settled at stage 3 and the only check for it is reading the
titles.

### Two things `qa.py` will report on a deck built this way

Both are correct behaviour on its side and expected on yours. Neither is fixed by editing
the deck.

- **Every `[TBD]` comes back as a placeholder finding.** A deck carrying six unknowns
  reports sixteen findings and exits 1. That is the point: the findings *are* the
  pre-meeting checklist. Say so in the report, and keep the two states apart — a deck can
  be **finished as an approved draft** and **not yet finished to present**. Never delete a
  `[TBD]` to get a clean run.
- **Every section divider reports `no title found`** and shows as `(no title found)` in
  the ghost-deck listing. A divider design carries its display title in a placeholder
  `qa.py` does not read as a title. The title is there — check the render — and the
  ghost-deck test skips dividers anyway.

---

## Density, once the slides exist

- **One idea per slide.** More than one assertion means split it.
- **A visual relationship beats a list.** Whenever you are about to write bullets, ask
  whether the content is a comparison, a sequence, a matrix or a single number. It
  usually is one of the four, and `references/accents.md` has the pattern for it.
- **~25 words per slide** outside tables. The written companion is the leave-behind; the
  deck is not a document.
- **Every slide stands alone.** Decks get forwarded to people who were not in the room:
  sources on the slide, units labelled, no claim that needs narration to make sense.
- Detail goes to the **appendix**, never into a smaller font.

---

## Attribution

The scaffold set, the pyramid/MECE framing and the headline-flow test are adapted from
`executive-slide-skill` (Apache-2.0) and `academic-pptx-skill` (MIT); the staged,
user-gated workflow follows `Olito-Labs/slides` (MIT); the three-preview picker in stage
4 follows `high-quality-slides` (MIT). The Rollout scaffold, the intake set and the
small-model staging are ours. See `THIRD-PARTY.md`.
