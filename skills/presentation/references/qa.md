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
  six characters, and collapsed to one line naming every slide it is on. The
  threshold used to be twenty-five, on the theory that only a body paragraph
  could be a real leftover; two runs against a corporate template shipped a
  15-character label off its canvas design on 13 of 15 and 9 of 10
  slides, in German, on finished English decks, and both reports said clean.
  A design's short labels are retext keys like any other — the catalog lists
  them;
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
- **type under the smallest size the template itself sets** — 9pt on the house
  template. Below that it is legible on the monitor it was built on and not
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
shapes it stacks — costs 37–43 s on the 21-design house template, against 29 s
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

## The vision reviewer

```bash
export PRESENTATION_REVIEW_MODEL=google/gemma-4-31b-it
export OPENROUTER_API_KEY=...          # or point at a local server
```

With that set, `make_deck.py` sends every render to that model and prints what
it saw, in the report and on stdout. Standalone:

```bash
scripts/review_renders.py render/ --openrouter --model google/gemma-4-31b-it
scripts/review_renders.py render/ --model gemma3:27b     # local ollama
scripts/review_renders.py render/ --dry-run              # see the request
```

**The model must accept images.** Point this at a text-only model and it will
return a confident review of a slide it never saw, which is worse than no
review — check the model's input modalities first. `deepseek-v4-flash` is
text-only; `gemma-4-31b-it` and `qwen3.6-27b` are not.

It is asked for what measurement cannot reach: figures off-centre, peers styled
inconsistently, cards at different heights, a Venn whose circles do not
overlap, a region that reads as unfinished, text left in another language. It
is told that overflow, off-slide geometry, contrast and collisions are already
measured, so it does not spend its findings re-reporting them.

**It is forbidden to report spelling, typos, grammar or punctuation.** The
deck's exact text is read from the file by `qa.py`; a vision model guessing
letters out of pixels is strictly worse, and on the first real run it produced
two findings of the form *"misspelling of 'EXFILTRATION' as 'EXFILTRATION'"* —
self-contradicting, and 2 of 10 findings wasted. Language and leftover
placeholder wording are still fair game; those are not letter-level claims.

**Findings are clustered across slides before they are printed.** Every slide
is its own call, so a model has no idea it already reported the footer four
slides ago — it just reports it again in different words. Findings whose
wording *and location* overlap by 40% of their content words are folded into
one line naming every slide. That run went from ten lines to six, with the four
footer findings becoming one; a second run described the same footer as
"placeholder text in a different language" twice and "placeholder text in
footer" once, which share one content word out of four and need the shared
"bottom left footer" to come together.

**Expect the findings themselves to move between runs.** Those two runs over
the same deck agreed on the footer and on the emptiest slide and on nothing
else — where the first invented two spelling defects, the second invented a
misaligned Venn and an off-centre table, both measurably false. Banning a
category of junk redirects it rather than removing it; the model has an
appetite for finding something on every slide. Treat a finding as a reason to
open the slide, never as a fact about it.

**Advisory, always.** `review_renders.py` exits 0 whether it finds ten things
or none, and `make_deck.py` never fails a build on it. A vision model's
false-positive rate on slide layout is unknown, and a check that cannot be
satisfied is a check that gets ignored — taking the measured ones with it. Read
the finding, look at the slide, decide. A reviewer that is down, slow, or
talking nonsense costs you nothing: the run says nobody looked and carries on.

**But advisory is not the same as dismissible on argument.** A finding here is
dismissed only by a measurement you show:

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

This does not replace your own eyes where you have them. Where you have none it
is one of two things you do have, the other being `wireframe.py` — and unlike
the reviewer, the wireframe cannot invent anything.

## Visual QA

Slides are visual. Read the renders; do not reason about coordinates.

```bash
scripts/render.py deck.odp          # prints absolute PNG paths
```

Pass those paths to the view tool and look at every slide. After staring at the
generating code you see what you expected rather than what rendered, so look
fresh. What to check, most common first:

- **Text overflow, or text cut off at a box or slide edge.**
- Overlapping elements (text through shapes, lines through words)
- Footers or citations colliding with the content above
- Elements closer than 0.3", or margins under 0.5"
- Uneven gaps — cramped in one place, empty in another. **Start with the
  slides the run listed under "slides to look at"**, which are the ones already
  measured as painting almost nothing
- Columns or repeated cards not aligned consistently
- Low-contrast text or icons (light on light, dark on dark)
- Template decoration left behind after text replacement — e.g. a title
  underline positioned for one line when the new title wraps to two
- Leftover placeholder text

**Hidden slides are not rendered.** A deck with hidden slides produces fewer
pages than it has slides, so PNG numbers stop matching slide numbers after the
first hidden one. `dump.py` flags which slides are hidden, and
`office.rasterize.rasterize(..., hidden="show")` renders an all-visible copy
when the numbering has to line up.

`thumbnail.py` works on your finished deck too, not just on templates: one
contact sheet to scan, then per-slide renders only for the slides that look
wrong. That is a great deal cheaper than reading fifteen PNGs one at a time.

**If you cannot see images**, run `wireframe.py` over the render directory, say
that is what you worked from, and name the slides you could not verify — rather
than claiming the deck is checked.
Set `$PRESENTATION_REVIEW_MODEL` (see *The vision reviewer* above) so something
looks — then report what it found *as a model's opinion*, run `qa.py`, and
still ask the user to eyeball the slides, naming them. A build model without
vision must never describe a slide it has not seen.

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

Two findings worth acting on:

- **`.odp` → `.odp` is a fixed point.** Five successive open-and-save cycles
  produced identical property reports. Editing ODF costs nothing.
- **Every `.pptx` save appends one more master.** Measured 17 → 18 → 19 → 20 →
  21 → 22 across five successive saves, with no sign of a ceiling. The deck
  still renders correctly, but **do not loop a deck through repeated `.pptx`
  saves.** Build in one pass and save once; if the work needs several editing
  rounds, keep the working copy as `.odp` and export to `.pptx` at the end.

One-time normalisation to expect in both formats: a table shape is resized to
the natural height of its rows on first load. It is stable afterwards.

## Fast tests

```bash
python3 tests/test_geometry.py       # no LibreOffice, well under a second
python3 tests/roundtrip_report.py check fixture.odp   # after a LibreOffice upgrade
```

`test_geometry.py` covers the arithmetic that decides whether a slide is right —
`Box` slicing, table widths, text fitting, contrast, Venn bounds. Every case in
it is a defect somebody shipped. Run it after touching `office.draw`.
