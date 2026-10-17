# QA — what the checks mean, and what they cannot see

**A deck is not finished until the checks have run and every slide has a verdict
written in the report's table.** Every deck that shipped broken had skipped one
of the two — and of four runs measured on the same plan, three had looked at
the renders and shipped the defects anyway. The verdict column is that half of
the check made into something on disk: one row per slide, blank until you fill
it in, and a slide nobody has opened cannot have one.

```bash
scripts/make_deck.py build.py out.pptx --template brand.pptx   # do this
scripts/qa.py out.pptx --template brand.pptx                   # checks only
```

`make_deck.py` writes both its renders (`render/`) and its findings
(`<deck>-report.md`) **next to the deck**, not next to wherever you ran it. If a
build ends between "here is what is wrong" and "fixed it", that report is what
says so; a clean run overwrites it with a clean one.

**Only changed slides are measured again.** Laying text out to measure it is
the slow part — a 58-slide deck of 2,400 shapes took nine minutes per run, and
an agent fixing one slide at a time spent an hour re-checking 57 untouched
ones. So the per-slide findings go into `<deck>-qa-cache.json` under a
fingerprint of the slide (its XML without the ids a save renumbers, plus its
layout and media, not its notes), salted with the QA code and the template.
A slide that did not change reads its findings back; the deck-wide checks run
every time. Same deck: nine minutes cold, under twenty seconds warm, identical
report. Delete the file to force a full re-measure.

## What the automated checks report

They report what a reader would see as broken and nothing else:

- text overrunning its box *into* other text or off the slide, and a lone word
  broken across lines because its box is too narrow;
- **two blocks of text printed over each other** — a caption placed where a
  diagram's labels already are, a Venn label on the note beneath it, a title
  long enough to reach the list beside it. Measured on the ink rather than on
  the boxes, so a big text box holding two words does not report. A pair of
  *glyphs* is ignored (a tick on the tile it marks, a step number on its card),
  but a glyph and a block of words is not: a source line drawn under a table ran
  straight through a row whose cells held `?` and `✓`, and because both sides
  were tested against the four-character floor separately, the check written for
  exactly that defect never looked at the pair;
- **a shape drawn past the edge of the slide**;
- **a title the user approved that reached no slide** — only when a `plan.md` sits beside the deck. See below;
- wording still shared with the template, and leftover placeholders — down to
  six characters, one line per slide it is on. The
  threshold used to be twenty-five, on the theory that only a body paragraph
  could be a real leftover; two runs against a corporate template shipped a
  15-character label off its canvas design on 13 of 15 and 9 of 10
  slides, in German, on finished English decks, and both reports said clean.
  A design's short labels are retext keys like any other — the catalog lists
  them. **Only the slide's own design counts**: a word you drew yourself is
  content even when another design of the template happens to carry it. A
  phase bar drawn on 28 canvas slides ("Design", "Testen", "Bereitstellen" —
  the sample labels of an unrelated design) once produced 112 findings, every
  one allowlisted; it now produces none. A deck edited rather than built has
  no recorded designs and is compared against the whole template;
- **a picture still on the slide that came from the template** — `retext`
  replaces words and never pictures, so a design's stock photo and its process
  diagram survive a deck that reports clean on text, still in the template's
  own language;
- **the template's own footer**, which is a page property (`page.FooterText`)
  rather than a shape and which nothing that walks shapes can see. Four decks
  shipped the template's own `subject | title | place | date` under English
  slides — every slide of every deck — and all four reports said clean. Fix it
  with `build_deck(..., footer="…")` or `set_footer(page, …)`; `None` hides it;
- **markdown** (`~~x~~`, `**x**`) that will render as literal characters;
- **emoji** — ✅ ⚠️ ❌ and the rest are colour-font pictures, sized and coloured
  by whatever emoji font is installed and drawn differently by LibreOffice and
  PowerPoint. One run put them in a status column of a board deck. The
  monochrome ✓ and ✗ are ordinary glyphs in the deck's own font and are not
  flagged;
- **type under the smallest size the template itself sets** — 9pt on one
  corporate template we measured. Below that it is legible on the monitor it was built on and not
  from a seat. `D.text(..., role="footnote")` takes the size from the scale
  rather than picking one per slide;
- **contrast** — each run's colour against the fill or slide background behind
  it. Under 1.8:1 the text is not visible at all and is always a bug; between
  that and 3.0:1 it is a brand colour used at a size it was not meant for, and
  that one is your call. A slide that sets no background of its own counts as
  **white**, which is what it renders as: treating it as "unknown" meant text
  on an ordinary body slide was never checked at all, and a label in the
  brand's own lime — 2.15:1 on white — reported clean. Pairings the template
  itself ships are baselined out, so what fires is the colour *this build*
  chose;
- **occlusion** — an opaque shape covering the one beneath it, which is what a
  `cut_*` pair unpacked in the wrong order looks like on the slide;

then `validate.py` on a `.pptx`. Exit code 1 means findings. Pass `--template`
for any template-derived deck; without it neither the leftover-wording nor the
stray-picture check can run, and a bleed the template does on purpose cannot be
told from one you introduced. For a house-default deck the template to pass is
`templates/acme/template.pptx`.

**The off-slide check is the one that catches bad arithmetic.** A table whose
column widths add up to more than the slide, or a diagram given a radius too
large for its slot, is not a text-overflow defect — every box holds its own text
perfectly — so nothing else notices that the last column is over the trim. It
fires per object rather than per cell, and a `--template` run baselines the
decorations the template itself hangs past the edge. Lay out from the `Box`
`D.canvas` returns and it stays quiet.

**The template baseline is cached, and you want it to be.** Measuring what the
template already does — its wording, its bleeds, its brand colour pairings, the
shapes it stacks — costs about 40 s on a 21-design template, against 29 s
to check a ten-slide deck: well over half of a `--template` run, and it used to
be paid again on every build iteration. It is a pure function of the template,
so it is now written to `$XDG_CACHE_HOME/presentation-skill/` and reused. The
key covers the template bytes **and** the checker sources, so editing a rule in
`qa.py` or `office/template.py` recomputes every baseline by itself — a cache
that could serve an answer the current code would not produce would hide the
finding someone just taught the checker to make. `PRESENTATION_NO_CACHE=1`
disables it; `PRESENTATION_CACHE_DIR` moves it.

Calibrated against six decks built from one brand template: silent on the
untouched template, and on the rest it finds precisely the defects their authors
missed and shipped — a leftover template paragraph, pill labels breaking
mid-word, a headline in black on a navy divider, a black fill swallowing the
panel it was meant to be a band of, four markdown strikethroughs, and two German
template graphics still sitting on a slide whose text had been fully replaced.

**Text taller than its box is not itself a defect**, which is why `qa.py` does
not report it. Designed templates size a body box to one line and let real
content flow into the whitespace below; nothing is clipped and the slide is
fine. Do not "fix" that.

## Slides to look at

Reported under their own heading, separately from the findings, and **they do
not fail the run**. A slide qualifies when it paints under a quarter of its own
area, or leaves at least a third of its height in a single band nothing
touches.

Both numbers come from measuring two agent decks whose slides had already been
judged by eye: the ones that read as unfinished painted 9–24 % of the slide,
the ones that read as designed painted 32–89 %, and nothing landed in between.
The band rule catches the other shape of the same defect — everything packed
under the title with the bottom half bare, which scored 30–51 % where a
finished slide scored 8 %, the margin below the footer.

They are not findings because **no measurement separates a deliberate
one-sentence statement slide from a slide somebody forgot to finish** — they
score identically, and both decks contained one of each. A check that cannot be
satisfied is a check that gets ignored, and it takes the checks around it down
with it. So this one names the slides and stops: four of fifteen is a different
instruction from "look at the renders", which is what the runs that shipped
these decks were given. One of them looked, saw the empty slides, called them
minor and shipped; the other never opened an image.

If a listed slide is sparse on purpose, it is fine. Say so and move on — in the
verdict column, where saying so is recorded.

### Top-heavy slides are a finding

The bare band above has a specific and fixable shape, and it is pulled out as a
**hard finding** on its own: a content slide that paints real ink against its
title and then leaves at least a third of its height as one empty band starting
in the lower half. That is not a statement slide (a statement paints little
overall); it is content crammed against the title by hardcoded offsets, and it
is how a deck ships looking "not really pretty" while passing every other check
— nothing overflows, nothing collides, nothing is off the slide.

`balance_findings` measures it and `make_deck.py` fails the build on it. The
fix is to lay out from the `Box` `D.canvas` returns so the content fills the
slide (`references/drawing.md` §Composition). The only legitimate way past it
is the allowlist, with a reason — see below.

### Hand-placed slides: looked at first, and checked for alignment

Every shape `office.draw` creates is named after what made it — `D:cards`,
`D:steps`, `D:table` when a composition helper laid it out, `D:rectangle`,
`D:text`, `D:ellipse` when the build placed it by hand — and the name
survives the save. The template's own shapes are never counted. Two things
read those names:

- **A slide to look at, not a finding.** A slide with more than twelve
  hand-placed shapes and none from a helper is listed under *slides to look
  at*, and its row in the verdict table tells the vision-agent so. A custom
  diagram no helper draws is legitimately built that way, so the count alone
  never fails a build.
- **Near-misses are a finding.** Among the hand-placed shapes, *peers* — same
  kind, sizes within 10 %, like the cells of one row or the cards of one
  column — must share their edges exactly. Two peers whose closest edges miss
  by 0.6 to 3 mm, or a run of three or more whose gaps differ by 0.6 to 3 mm,
  fail the gate. That is the defect layouts worked out by page percentage
  have: a DeepSeek run placed 570 rectangles and text boxes that way and
  drifted on slide after slide. A label inside a box is an inset, not a miss;
  offsets over 3 mm are a choice.

The fix is to lay the shapes out from one `Box` (`grid`, `cut_*`), so peers
share edges and gaps by construction (`references/drawing.md`). An offset that
is meant — a staggered row — is acknowledged as kind `alignment`, content
`near-miss`, with a reason.

### Defects a glance shows, measured

A field run drew ten slides whose report said "no findings". Seven of them
had defects a glance shows. These checks now catch them:

- **Text on a translucent panel:**
  - a panel at 85 % transparence is still a panel;
  - its colour is blended with what lies under it before contrast is
    measured;
  - white on a pale tint is reported as the near-invisible text it is.
- **A word wider than its box:** "Wirkung →" in a box 6 mm wide breaks
  letter by letter. Each paragraph's longest word is measured in that
  paragraph's own format.
- **A box far shorter than its text:** a label given 0.01 mm of height
  paints a full line, centred on the box, into the heading above it. It is
  measured as that line.
- **Text across panels** (kind `straddle`): a line lying across two or more
  drawn panels, such as a source line in a box the height of the content
  area. Text belongs inside one panel or outside all of them.
- **A shape over text:** an opaque, textless shape stacked above a block of
  text and crossing its letters, such as a bar laid through three columns.
- **A rival headline:** text above the title, or in the footer band, set as
  large as the title. Typically a kicker written into an empty placeholder,
  which takes the master's body style. Use `D.kicker`.

### Clean is not done

`make_deck.py` exits **3** when every automated check is clean but slides
have not been seen by the vision-agent, and prints `NOT DONE: …`. A run
that stopped at a clean gate and reported a finished deck had shown its
slides to nothing. On exit 3 it writes the briefs itself and prints the
exact delegation for the first unseen slide: the next action, not a
description of it. Exit 0 means clean *and* seen. `--parts` and `--slides`
runs are partial and keep exit 0.

### Chart placeholders and unsourced figures

Two more lines can appear under *slides to look at*. Neither fails a build:
- **chart placeholder**: the slide reserves a place for the author's Excel
  chart (`D.chart_placeholder`). Expected, and the verdict table tells the
  vision-agent so.
- **figures and no source line**: three or more percentages, amounts,
  decimals or large numbers, and no "Quelle", "Source", "Note" or
  "Beispieldaten" anywhere on the slide. Add `D.source`, or retext the
  design's own source line.

### The allowlist: how correct content gets a clean report

A finding is a defect this build introduced. **Correct content that merely
matches the template is not a defect** — an SDLC phase name that is the point
of the slide, a divider's full-bleed background graphic, a brand's own footer
URL. Before the allowlist, a deck carrying such content faced a choice between
shipping a permanently red report and corrupting the content, and runs burned
whole transcripts re-litigating it.

`<deck>-allowlist.json`, beside the deck, records such findings as intentional:

```json
{"acknowledgements": [
  {"slide": 16, "kind": "leftover-wording", "content": "Testen",
   "reason": "SDLC phase label on the lifecycle circle"},
  {"slide": 5,  "kind": "stray-image", "content": "image5.png",
   "reason": "divider background graphic"}
]}
```

- `kind` is one of `leftover-wording`, `stray-image`, `footer`, `balance`,
  `no-title`, `variety`, `alignment`, `straddle` or `vision`. A `vision` entry overrules the
  vision-agent and counts only with the user's `"approved_by"` (see *Visual
  QA* below); it is matched by the gate, not by `qa.py`, so it is never
  reported as stale.
- Matching is **exact** on kind + slide + content. There is no wildcard, so an
  entry cannot cover a finding the author had not seen — which is what an entry
  *should* cover, slide by slide, content by content.
- An entry with no `reason` is an error, not an acknowledgement: ignoring a
  check with no justification is a way to ignore the check.
- A malformed allowlist is an **error** (exit 2), never a silent pass or a
  silent gate.
- An acknowledgement that matches **nothing** in the deck — the finding was
  fixed, or never existed — is itself reported as a stale entry, so the file
  cannot quietly grow into a blanket "pass everything" list. A clean deck with a
  stale allowlist is not clean.

An acknowledged finding is dropped from the findings and listed separately as
*acknowledged*; a deck whose only remaining findings are acknowledged exits
clean. This is the escape hatch that keeps the gate honest: you can get a clean
report, but only by naming each finding you are keeping and saying why.

### `--require-verdicts`: gate on having been looked at

The report's verdict column is the record that something with eyes saw each
slide, and "the deck is not finished while a verdict is blank" was prose. 
`make_deck.py --require-verdicts` makes it an exit code: if any slide's verdict
cell is blank the run exits 1. Use it for a deliverable that must actually be
reviewed before it ships — it turns the visual half of QA into the same kind of
gate the automated half already is.

## The ghost deck, and the lines to look at

Every report prints **the titles of every slide, in order, and nothing else.**
Read them as a block. They should tell the whole argument on their own, each one
a claim following from the last. If they read as a table of contents — "The
Trifecta", "Why it matters", "Next steps" — the deck lists rather than argues,
and the fix is in the storyline, not in the slides.

This is the one check nothing else here replaces. Two decks built from the same
research passed every measurement above; one read as an argument end to end and
the other as an agenda, and only the title block told them apart.

Under it, **the lines to look at**: titles that name a topic instead of making a
point, titles that ask a question, a deck ending on "Vielen Dank", and the words
from the banned list in `references/language.md`. These are prose, not defects,
and they never fail a build — *robust* has an honest use in statistics and a
four-word title is right on a section divider. Overriding one is a decision you
are allowed to make; not seeing it is not.

**Every listed slide also gets a wireframe in the report**, and any render can
have one on demand:

```bash
scripts/wireframe.py render/slide-01.png     # one slide
scripts/wireframe.py render/ --slides 1,7-9  # several
```

It draws the slide as a coarse grid of ink — ` ` to `@`, 64 × 24 — with the
centroid, the widest blank band and the four margins. It says nothing about
whether a slide is any good, and it is exact about the three things that go
wrong most: a slide that is empty, a slide banded under its title, and a block
that is not where the rest of the slide says it should be. It needs no model,
which is the point: it is the visual evidence for a reviewer with no eyes, and a
second opinion for one with them.

```
 0|    .                                                           |
 1|    .:.--.:-:.-:::::=:::-::::--:-.-:--                          |
   … sixteen blank rows …
18|                  ----.=--:==--.--:==-:-=---=-                  |
ink 6% · centroid x=0.45 y=0.47 · largest blank band 67% of the height
```

That is a title slide nobody finished, and you can tell without seeing it.

## Titles the user approved

Only when `plan.md` — stage 3 of `references/storyline.md` — sits **in the same
directory as the deck**. It is read back and every approved title is looked for in
the built file:

| | |
|---|---|
| the sentence is on a slide | nothing to say |
| most of it is | prose: *reworded* — a tightened headline is fine, a moved claim is not |
| it is nowhere | a **finding**, and this one fails the build |

The last row is the only wording check here that is not advisory, and it is not a
judgement: the string is in the file or it is not. It exists because a run carried
an approved plan through stage 4 and then wrote the branch name into the title slot
on eleven slides of sixteen. Nothing else noticed. The build was clean, the renders
were fine, and the deck argued nothing.

**Placement is deliberately not checked** — only presence. The display title of
a divider design is not readable as a title at all, and on the title slide the
governing thought belongs in the subtitle, so a check on *which slot* the sentence
landed in would report three correct slides as broken on every deck this template
builds.

Alongside it, needing no `plan.md`: **one headline used on more than one slide**,
reported as prose. That is the visible shape of the same defect — a section name
sitting where sixteen different claims should be.

## Visual QA through the vision-agent

**You cannot see the renders yourself** — a build model has no image processing —
so the eyes of this skill are the **vision-agent** (the `vision` subagent). A
build model must never describe a slide it has not seen, and must never write
the vision-agent's brief itself:

```bash
scripts/vision_brief.py slides deck.pptx --language German   # after make_deck.py
# for each slide it lists, one call:  Read the brief at <render>/slide-N.brief.md and follow it exactly.
# save each reply, unedited:          <render>/slide-N.answer.md
scripts/vision_brief.py collect deck.pptx                   # fills the verdict table
```

**Every answer carries proof that eyes saw the slide.** The brief points at
`render/review/slide-N.png`: the render with a strip under it that prints a
random review code, which the answer has to return in its JSON, and the
slide's four quarters enlarged beneath it (A–D). With the full slide alone,
gemma-4-31b called a bar struck through a paragraph "ok"; with the quarters
it reported it. The ledger
beside it (`render/review/ledger.json`) keeps only the code's SHA-256, so
nothing that reads files can produce it. An answer without the right code is
"not seen", however plausible it reads. That closes the failure this exists
for: a run whose answers had been wiped by a rebuild retyped twenty-nine of
them from memory, the gate passed, and three dividers shipped with invisible
titles the real reviewer had flagged.

**Answers survive rebuilds when the slide does.** The ledger also records
which render each code was issued for. `make_deck.py` keeps the answer of
every slide that renders to the same bytes — following it if the slide moved —
and drops only the changed ones; `vision_brief.py slides` then lists just
those. LibreOffice renders an unchanged slide byte-identically.

**Why the brief is generated, not written** — measured on gemma-4-31b against
slides with planted defects, on a German and an English template:

| | a brief written by the build model | `vision_brief.py` |
|---|---|---|
| images per call | all of them | one |
| leftover template label on every slide | found on 2 of 18 runs | 18 of 18 |
| off-centre diagram | missed | found (English deck) |
| tiny, crammed text | missed | found |
| hollow cards, a diagram tiny for its space | passed as `ok` | found |
| invented defects ("cards at different heights") | on four slides, every run | none |
| a real deepseek deck a designer would reject | "all 9 ok" | 6 of 9 sent back, with the reason |

The brief describes the slide on a 3×3 grid before judging it, asks ten closed
questions (headline, focal point, empty band, off-centre, overlap, peers,
hollow boxes, broken word, language, unfinished) with *unsure* allowed and a
location required for every *yes*, and ends in a JSON verdict with a concrete
fix per finding. It carries what `qa.py` measured for that slide, and tells the
vision-agent what is already measured — spelling, leftover template wording,
overflow, contrast, sizes — so it does not re-report or guess them. **Never ask
the vision-agent about colours**: one named a brand's cyan "yellow-orange", and
the run that believed it picked the wrong accent.

**A `fix` verdict is work, not a record.** Change the build script, re-run
`make_deck.py`, brief the changed slides. `make_deck.py deck.pptx
--require-verdicts` reads the report and the saved answers without rebuilding,
and exits 1 while any slide lacks an answer with its review code, any answer
still holds a medium or high finding, a `vision` overrule waits for the
user's approval, or the report belongs to an earlier build. Editing a verdict
cell clears nothing — the table is a view of the answers. A run that rewrote
seven "cards mostly empty" verdicts as "cards fill the area proportionally"
is why.

**An overrule is the user's decision.** A finding a measurement disproves is
a `vision` entry in `<deck>-allowlist.json` — `content` the question, the
measurement as `reason` — and it holds the gate until the user adds
`"approved_by"`. The build model never writes that field. Every approved
overrule is printed on every gate run, pass or fail, so it reaches whoever
reads the result. Without a vision-agent, a slide the user looked at is an
entry with `"content": "seen"` and the user's `approved_by`. The run that
made this rule overruled four findings itself, one of them "no headline" on
dividers whose titles were in the file and under a picture: present is not
visible, and the reviewer was right.

**Treat what the vision-agent reports as a reason to look, never as a fact.**
A vision model's false-positive rate on slide layout is unknown and its findings
move between runs — where one invented a spelling defect, another invented a
misaligned Venn that was measurably false. A finding here is dismissed only by a
measurement you show, never by reasoning about why the reviewer must be wrong:

| It says | Settle it with |
|---|---|
| text in the wrong language | `qa.py --template`, and `footer_text(page)` for the footer |
| a diagram off-centre, a block not where it belongs | `wireframe.py` — the centroid and the margins |
| a slide that reads as empty | the ink share and the blank band, same place |
| cards at different heights, peers styled differently | `probe.py` — the geometry and the colours, per shape |
| two things printed over each other | `qa.py` — the collision check measures the ink |

The failure this table exists for: a run with no eyes was told, five times over,
that the footer was in the wrong language. It searched the slide's shapes for
the string, found nothing — the footer is a page property — and recorded the
finding as a hallucination in its own report. The search was wrong, the reviewer
was right, and the deck shipped in two languages. **If nothing can settle it,
the finding stands.**

**Hidden slides are not rendered.** A deck with hidden slides produces fewer
pages than it has slides, so PNG numbers stop matching slide numbers after the
first hidden one. `dump.py` flags which slides are hidden, and
`office.rasterize.rasterize(..., hidden="show")` renders an all-visible copy
when the numbering has to line up.

`thumbnail.py` works on your finished deck too, not just on templates: one
contact sheet for the vision-agent to scan, then per-slide renders only for the
slides that look wrong. That is a great deal cheaper than reading fifteen PNGs
one at a time.

**If the vision-agent is unavailable**, run `wireframe.py` over the render
directory, say that is what you worked from, and name the slides nothing with
eyes has seen — rather than claiming the deck is checked. The wireframe cannot
invent anything, but it shows only emptiness, banding and placement — never
whether a slide is any good.

## Content QA

`qa.py` covers this. To grep a deck for something specific of your own:

```bash
scripts/dump.py deck.odp | grep -niE "lorem|ipsum|\bTODO|\[insert|platzhalter"
```

## File QA (.pptx only)

```bash
scripts/office/validate.py deck.pptx
scripts/office/validate.py deck.pptx --original template.pptx
```

Checks XML well-formedness, relationships, content types, unique IDs, layout
references, master/theme sharing, chart constraints, and the slide-XML defects
PowerPoint refuses to open a file over. Each failure names its fix. Exit codes:
`0` pass, `1` validation failure, `2` bad usage or unreadable file.

- **A deck built through this skill should pass, exit 0.** It did not always:
  LibreOffice's `.pptx` export writes `Target="../customXml/item1.xml"` into the
  package root (a reference that climbs out of the package — reported CRITICAL)
  and drops the `<Default>` content type for `.wmf` media, which together fired
  22 findings on every deck this skill produced. `impress.save_as` now repairs
  both, so **a failure here is yours and worth reading.** `repair_pptx(path)`
  runs it by hand if you ever save through raw UNO.
- **Pass `--original` for any template-derived deck** — `qa.py --template` does
  it for you. A template may itself contain parts the XSD rejects; `--original`
  baselines against it so the template's own defects do not read as yours.
- `--auto-repair` adds missing `xml:space="preserve"` to `<a:t>` runs with
  leading or trailing whitespace, in place.
- The slide check is a deliberate denylist of defects PowerPoint actually
  rejects, not full schema conformance — a schema-invalid deck that PowerPoint
  opens fine is reported as passing, by design.
- **There is no equivalent for `.odp`, and none is needed.** LibreOffice writes
  ODF itself. QA an `.odp` by rendering it and reading it back.

## What a round trip preserves

Measured on this stack with `tests/roundtrip_report.py`, not assumed. A fixture
deck carrying 11 masters, real layouts, speaker notes, a slide transition, a
hidden slide, a shape animation, a table, an embedded image, an OLE chart and
document properties was round-tripped three ways:

| | `.odp` → `.odp` | `.odp` → `.pptx` | `.pptx` → `.pptx` |
|---|---|---|---|
| Slides, hidden flags | kept | kept | kept |
| Masters & layouts, applied design | kept | kept | kept |
| Speaker notes | kept | kept | kept |
| Slide transitions (effect, timing) | kept | kept | kept |
| Shape animation effects | kept | kept | kept |
| Tables and their content | kept | kept | kept |
| Images, OLE charts | kept | kept | kept |
| Text and character formatting | kept | kept | kept |
| Document properties | kept | kept | kept |
| Master **count** | kept | **grows** | **grows** |
| Theme **colour scheme** | kept | **lost** on invented masters | **lost** on invented masters |

Two findings worth acting on:

- **`.odp` → `.odp` is a fixed point.** Five successive open-and-save cycles
  produced identical property reports. Editing ODF costs nothing.
- **Every `.pptx` save appends one more master.** Measured 17 → 18 → 19 → 20 →
  21 → 22 across five successive saves, with no sign of a ceiling. The deck
  still renders correctly, but **do not loop a deck through repeated `.pptx`
  saves.** Build in one pass and save once; if the work needs several editing
  rounds, keep the working copy as `.odp` and export to `.pptx` at the end.
- **The masters LibreOffice invents carry its own theme colours, not the
  template's.** Measured on a 3-master template whose every theme set `dk2`
  to the brand navy: the saved deck had six masters, two of them new and
  carrying `dk2=cccccc`. A divider layout that paints its background as
  `<a:schemeClr val="dk2"/>` rendered light grey under white text, about
  1.6:1, and every shape-level check passed because the defect lives in the
  theme behind the shapes. `build_deck` repairs this after the save
  (`deckbuild.repair_theme_colours`): each exported master is mapped back to
  the template master it came from by layout-name overlap and its colour
  scheme is restored. It reports every theme it rewrote; a theme shared by
  masters that disagree is left alone.

One-time normalisation to expect in both formats: a table shape is resized to
the natural height of its rows on first load. It is stable afterwards.

## Fast tests

```bash
python3 tests/test_geometry.py       # no LibreOffice, well under a second
python3 tests/test_theme_repair.py   # same: four XML strings stand in for a .pptx
python3 tests/roundtrip_report.py check fixture.odp   # after a LibreOffice upgrade
```

`test_geometry.py` covers the arithmetic that decides whether a slide is right —
`Box` slicing, table widths, text fitting, contrast, Venn bounds. Every case in
it is a defect somebody shipped. Run it after touching `office.draw`.
`test_theme_repair.py` holds down the theme-colour repair above: the
sysClr-versus-srgbClr comparison, the fallback for masters LibreOffice
invented, and the refusal to paint one brand in another's colours. Run it
after touching `office.deckbuild`.
