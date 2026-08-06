# QA — what the checks mean, and what they cannot see

**A deck is not finished until the checks have run and you have looked at the
renders.** Every deck that shipped broken had skipped one of the two.

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
  the boxes, so a big text box holding two words does not report;
- **a shape drawn past the edge of the slide**;
- wording still shared with the template, and leftover placeholders — down to
  six characters, and collapsed to one line naming every slide it is on. The
  threshold used to be twenty-five, on the theory that only a body paragraph
  could be a real leftover; two runs against the delta template shipped the
  15-character design-8 label `"Navigator (1/2)"` on 13 of 15 and 9 of 10
  slides, in German, on finished English decks, and both reports said clean.
  A design's short labels are retext keys like any other — the catalog lists
  them;
- **a picture still on the slide that came from the template** — `retext`
  replaces words and never pictures, so a design's stock photo and its German
  process diagram survive a deck that reports clean on text;
- **markdown** (`~~x~~`, `**x**`) that will render as literal characters;
- **contrast** — each run's colour against the fill or slide background behind
  it. Under 1.8:1 the text is not visible at all and is always a bug; between
  that and 3.0:1 it is a brand colour used at a size it was not meant for, and
  that one is your call;
- **occlusion** — an opaque shape covering the one beneath it, which is what a
  `cut_*` pair unpacked in the wrong order looks like on the slide;

then `validate.py` on a `.pptx`. Exit code 1 means findings. Pass `--template`
for any template-derived deck; without it neither the leftover-wording nor the
stray-picture check can run, and a bleed the template does on purpose cannot be
told from one you introduced. For a house-default deck the template to pass is
`assets/company-template.pptx`.

**The off-slide check is the one that catches bad arithmetic.** A table whose
column widths add up to more than the slide, or a diagram given a radius too
large for its slot, is not a text-overflow defect — every box holds its own text
perfectly — so nothing else notices that the last column is over the trim. It
fires per object rather than per cell, and a `--template` run baselines the
decorations the template itself hangs past the edge. Lay out from the `Box`
`D.canvas` returns and it stays quiet.

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

If a listed slide is sparse on purpose, it is fine. Say so and move on.

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

This does not replace your own eyes where you have them. It replaces *nothing*,
which is what a text-only build model has.

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

**If you cannot see images**, say so rather than claiming the deck is checked.
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
