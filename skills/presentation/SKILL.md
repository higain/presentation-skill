---
name: maxi-presentation
description: "Create, edit, read and QA presentations — .odp and .pptx alike, plus .otp/.potx templates and legacy .ppt. Use whenever a slide deck is involved: building a deck or pitch, filling in a template, editing or restyling an existing one, extracting text or speaker notes, converting between formats, or rendering slides to images for visual review. Triggers on \"PowerPoint\", \"Impress\", \"deck\", \"slides\", \"presentation\", \"Präsentation\", \"Folien\", or any .odp/.otp/.pptx/.potx/.ppt filename. The user's word for a file says nothing about its extension — a file someone calls \"the PowerPoint\" is often an .odp — so check the actual file, and save a deck back in the format it arrived in."
compatibility: "Needs LibreOffice Impress 25.2+ with python3-uno, plus python3-lxml, python3-defusedxml, python3-pil and poppler-utils. All dependencies come from apt — never run pip or npm. Works offline; no network access required."
license: Proprietary. LICENSE.txt has complete terms
---

# Presentations: .odp and .pptx

**The engine is LibreOffice Impress, driven over UNO.** It opens the real
document and saves it in its own format, so an `.odp` stays ODF and a `.pptx`
stays OOXML. Nothing here rewrites a file format by hand.

| Task | How |
|---|---|
| **Read** a deck | `python3 scripts/dump.py deck.odp` |
| **Survey** a template's layouts | `python3 scripts/thumbnail.py template.odp` |
| **Inspect** shapes to edit them | `python3 scripts/probe.py template.odp` |
| **Create** a new deck | `office.impress` — see [Building a deck](#building-a-deck) |
| **Build in the house style** | no template supplied — see [Company template](#company-template-default) |
| **Adapt** a template deck | `office.template` — see [Adapting a template deck](#adapting-a-template-deck) |
| **Edit** an existing deck | `impress_document(path)` — saves in place, same format |
| **Convert** formats | `python3 scripts/convert.py in.odp out.pptx` |
| **Render** for visual QA | `python3 scripts/render.py deck.odp` |
| **Check** a finished deck | `python3 scripts/qa.py deck.pptx --template brand.pptx` |
| **Schema-check** a `.pptx` | `python3 scripts/office/validate.py deck.pptx` |

**Two rules that override anything below:**

1. **Keep the format the deck arrived in.** Editing an `.odp` produces an
   `.odp`. Convert only when the user asks for the other format.
2. **A new deck defaults to `.pptx`** unless the user asks for `.odp` or names
   an output path that says otherwise.

**And a default for content:** when a deck is requested with no template of the
user's own, build it from our house template — see [Company
template](#company-template-default) — rather than inventing a look.

## Scripts

Paths are relative to this skill's directory.

| Script | What it does |
|---|---|
| `scripts/dump.py deck.odp [-o out.md]` | Deck to markdown: per slide the layout, master, text, tables, speaker notes, empty placeholder slots and a hidden flag. Works on every supported format |
| `scripts/probe.py deck.odp [--slides 4,7-9]` | Every shape's **index**, position, size, fill, line and character formatting, plus a per-slide reading-order block grouping shapes by column. `dump.py` says what a slide *says*; this says what it *is*. Run it before editing any template slide |
| `scripts/thumbnail.py deck.odp [prefix]` | One labelled contact sheet of all slides, for picking template layouts. Output is named after the deck, so two decks never overwrite each other |
| `scripts/render.py deck.odp [-o dir] [--dpi N] [--slides 3,7-9]` | One PNG per slide; prints their absolute paths. **This is visual QA** |
| `scripts/qa.py deck.pptx [--template brand.pptx] [--slides 3,7-9]` | Broken text the renders would show: overflow, mid-word breaks, leftover template wording, placeholders — plus `validate.py` on a `.pptx`. Exit 1 on any finding |
| `scripts/convert.py in.odp out.pptx` | Convert between `.odp`/`.pptx`/`.otp`/`.potx`/`.ppt`/`.pdf` |
| `scripts/office/validate.py deck.pptx [--original src.pptx] [--auto-repair]` | Schema, relationship, content-type, chart and slide checks. **`.pptx` only** — see [File QA](#file-qa-pptx-only) |
| `scripts/catalog_extract.py template.pptx -o catalog/` | **Build-time only, not per deck:** per-layout dossiers (render + structure) for authoring the house catalog. See [Maintaining the default](#maintaining-the-default) |
| `tests/roundtrip_report.py check fixture.odp` | Re-proves the fidelity claims below after a LibreOffice upgrade |

Three references, read only when the job calls for them:
`references/design.md` (inventing a look — palette, layout, what makes a deck
look generated), `references/company-template.md` (our house catalog — which
template slide for which content; see [Company
template](#company-template-default)) and
`references/pptx-xml.md` (hand-editing OOXML).
`probe.py --json` emits the same survey as data when a script needs to pick
shapes itself.

`markitdown.sh` is **not** useful here: it rejects `.odp` outright, and on
`.pptx` it reports text without layouts, masters or notes. Use `dump.py`.

**Importing `office.*` from your own build script:** put the skill's `scripts`
directory on the path — the one that *contains* `office/`, never `office/`
itself, which is the mistake that produces `ModuleNotFoundError: No module
named 'office'`. A relative `"scripts"` only works while the shell sits in this
skill's directory; from anywhere else write the absolute path this SKILL.md was
read from:

```python
import sys; sys.path.insert(0, "/absolute/path/to/presentation/scripts")
from office.impress import impress_document, slides, shapes
```

## Reading a deck

```bash
python3 scripts/dump.py deck.odp
```

```
## Slide 1
*layout 0, master «Candy»*

**title:** Quartalsbericht Q3
*empty slots: subtitle*

**Notes:**
> Mit der Umsatzzahl aufmachen.
```

`empty slots` is what makes a template readable — it names the placeholders a
layout offers to fill. Text is listed in z-order (the order shapes were added),
not top-to-bottom reading order.

## Building a deck

```python
import sys; sys.path.insert(0, "scripts")
from office.impress import (
    new_document, add_slide, add_textbox, add_image, placeholder,
    set_notes, save_as, slides, mm, pt,
)

doc = new_document()
page = slides(doc)[0]                      # a new document ALREADY has one slide
page.Layout = 0                            # 0 = title + subtitle
placeholder(page, "Title").setString("Quartalsbericht Q3")
placeholder(page, "Subtitle").setString("Vertrieb & Marketing")
set_notes(page, "Mit der Umsatzzahl aufmachen.")

page = add_slide(doc, layout=20)           # 20 = blank, place everything yourself
box = add_textbox(doc, page, mm(20), mm(20), mm(240), mm(25), "Ergebnisse")
box.CharHeight = 40.0
box.CharWeight = 150.0                     # bold
box.CharColor = 0x1E2761
add_image(doc, page, "chart.png", mm(20), mm(60), mm(120), mm(80))

save_as(doc, "deck.pptx")                  # the extension picks the filter
doc.close(False)
```

**Geometry is in 1/100 mm, not points or EMU.** A 16:9 slide is 28000 × 15750
(280 × 157.5 mm). Use `mm()` and `pt()` rather than raw numbers — a value that
looks like points is 25× too small and lands off-slide with no error at all.

Autolayouts worth knowing: `0` title+subtitle, `1` title+content, `19` title
only, `20` blank. Anything you position yourself wants `20`.

**`new_document()` already contains one slide.** Use `slides(doc)[0]` for the
title slide, as above — calling `add_slide()` for it leaves a stray blank slide
at the front of every deck. A document must also always keep at least one
slide: deleting the last one wedges LibreOffice, so `delete_slide()` refuses.

**Fill placeholders rather than adding text boxes** wherever a layout offers
one. Text in a placeholder inherits the master's fonts, sizes and position; a
free text box inherits nothing, and that is how decks end up looking unstyled.

### Editing an existing deck

```python
from office.impress import impress_document, slides, placeholder, set_notes

with impress_document("deck.odp") as doc:      # saved in place, still .odp
    page = slides(doc)[0]
    placeholder(page, "Title").setString("Neuer Titel")
    set_notes(page, "Angepasste Notiz.")
```

- `impress_document(path, save=False)` opens read-only — use it for inspection.
- An exception inside the block leaves the file on disk untouched.
- **`store()` is a no-op unless the document is modified.** `impress_document`
  handles that; if you call `doc.store()` yourself, call `doc.setModified(True)`
  first or nothing is written.
- `save_as(doc, path)` writes elsewhere and leaves the input alone, so a
  load → edit → `save_as` chain never overwrites the source.

### Duplicating a slide

```python
from office.impress import duplicate_slide, slides
new_page = duplicate_slide(doc, slides(doc)[1])   # inserted right after
```

It clones everything the slide references, including OLE charts and embedded
media. Do all structural work — duplicate, delete — **before** editing content,
or you will clone a slide you already edited.

**The copy always lands directly after its source, and that cannot be undone:
Impress has no way to move a slide.** `page.Number` is read-only, `XDrawPages`
offers only insert and remove, and the `.uno:MovePageFirst/Up/Down/Last`
dispatches return a dispatch object and then silently do nothing (measured, not
assumed). So copies of a slide can only form one contiguous run beside it. Any
deck whose layouts interleave — divider, chart, cards, divider, columns — is
**not buildable with `duplicate_slide` alone.**

Two ways round it. For a `.pptx` built from a template, **`build_deck` is the
better one**: it duplicates at full fidelity and then fixes the order in the
saved file, where slide order is just a list — see [Building from a template at
full fidelity](#building-from-a-template-at-full-fidelity). For everything else
— `.odp`, or copying between two open documents — use `clone_slide` below, which
rebuilds each shape and accepts the losses that come with that.

## Building from a template at full fidelity

**`.pptx` from a template: reach for this first.** `build_deck` duplicates the
designs you name — a real Impress duplicate, so nothing is rebuilt — drops the
template originals, and then puts the slides in your order by permuting
`<p:sldIdLst>` in the saved file. Slide order in a `.pptx` is only that list;
the slide parts do not care.

```python
from office.deckbuild import build_deck
from office.template import retext

def fill(page, slot, design):
    if design == 7:
        retext(page, {"Formelemente, Navigator": "Kapitel 1: Vorgehen"})

build_deck("assets/company-template.pptx", "out.pptx",
           [7, 8, 13, 20, 1],          # catalog slide numbers, in final order
           fill=fill)                   # called per page, before the save
```

`designs` may repeat a slide. `fill(page, slot, design)` is where retexting
belongs — it runs after the originals are gone, so `slides(doc)` is already the
deck you asked for.

Why it beats `clone_slide` on a template: measured against the delta asset over
designs 7, 8, 13, 20 and 1, diffing each result against its source render, mean
RMS **7.8 versus 22.8**, better on every slide. Bullet markers, per-run colour
inside a paragraph and think-cell's OLE data all survive here and do not survive
a rebuild — 6 think-cell references kept versus none. Same output size, same
validator findings.

Limits: `.pptx` only, since the reorder is an OOXML step, and it builds a *new*
deck from a template rather than editing one in place. For those, use
`clone_slide`.

## Adapting a template deck

The common job: a brand deck arrives with one slide per layout, and you have to
produce fifteen slides of your own content in that design.

```bash
python3 scripts/thumbnail.py brand.pptx        # which layout is which slide
python3 scripts/probe.py brand.pptx            # shape indices + geometry + fonts
```

`probe.py` is the step that saves the afternoon. To retext a template slide you
need the *index* of each shape, not its text, and probe prints exactly that:

```
  idx shape        type                    x      y      w      h  fill      line
  [2] Shape 2      ooxml-roundRect      2910   6625  22039   9413  #FFFFFF   #ECDFD6/46
  [3] Shape 3      ooxml-roundRect      4120   7730   2838   1043  #FF5A3C   -
  [4] Text 4       ooxml-rect           4544   7941   2203    725  -         -
      'SPEED'
        Arial 15  ·  bold  ·  #FFFFFF  ·  kern 40  ·  left
```

`[i]` is exactly `shapes(page)[i]`. **The table is z-order, not reading
order** — so each slide is followed by its shapes grouped into columns:

```
  reading order — 3 columns
  column 1    [3] 'Background'  [4] 'Set the context first. ...'
  column 2    [5] 'What we found'  [6] 'Present the...'  [7] 'Two shorter...'
  column 3    [8] 'What it means'  [9] 'Close with the implicat...'
  ** columns hold 2, 3, 2 shapes — this is NOT an even grid. **
```

Read that block before indexing anything. The slide above looks like three
columns of two in z-order and is not one: column 2 carries a second body box,
which shifts `[8]` and `[9]`. Assuming the even grid is what puts column 3's
heading into column 2 and leaves the template's own text on the slide.

Then build every slide in one pass and save once:

```python
from office.impress import impress_document, slides, shapes, save_as, delete_slide, set_notes
from office.template import clone_slide, retext, fit_width, move, resize, set_char_color

with impress_document("brand.pptx", save=False) as doc:
    tpl = slides(doc)                              # the template block
    COVER, DIVIDER, CARDS = tpl[0], tpl[3], tpl[6]

    page = clone_slide(doc, CARDS)                 # always appended, in call order
    left = retext(page, {                          # key = the TEMPLATE's wording
        "Tagged cards for quick points": "Three of four are past their end date",
        "SPEED":                         "Backend API",
    })
    assert not left, left                          # anything you forgot to replace
    set_notes(page, "Read the dates out loud.")

    for old in tpl:                                # drop the templates LAST
        delete_slide(doc, old)
    save_as(doc, "out.pptx")                       # one .pptx save — see below
```

- **`retext(page, {old: new})` matches shapes by the words they already carry,
  so you never count `s[i]` by hand.** This is the difference between a deck
  that builds first try and one that takes three attempts. Matching nothing, or
  two shapes at once, raises before anything is written and prints the slide's
  own indexed text, so a wrong excerpt corrects itself. A value of `None`
  deletes the shape; `""` empties it but keeps the box. The return value lists
  every shape you left untouched that still holds text.
- **Do not assume a column is one header plus one body box.** A designed column
  often splits its body across two text boxes, which shifts every later index —
  the single commonest way a template build goes wrong, and exactly what
  `retext` sidesteps. Where you do index by hand, re-read `probe.py` and note
  that its `[i]` is precisely `shapes(page)[i]`.
- `find_shapes(page, "pattern")` and `shape_by_text(page, "prefix")` locate a
  shape when you need the object rather than a replacement.
- **`clone_slide(doc, src, index=None)`** appends a blank page and rebuilds every
  shape on it, so slides can be produced in any order. Measured pixel-identical
  to its source on a fixture with text, an image, a table and a group; for a
  slide carrying an OLE chart or embedded media prefer `duplicate_slide` when
  the order allows it.
  **Clone in final slide order and leave `index` alone.** Passing `index` while
  the originals are still in the deck shifts them too, and what ends up where
  stops being predictable; clone order *is* deck order once the templates go.
  The source's master is carried over, so a slide on a second master — a dark
  divider in a light template — keeps its background. A placeholder whose run
  colour overrides its presentation style is rebuilt as a plain text shape,
  because the `.pptx` export drops direct run formatting on placeholders and
  re-inherits the master's text style; without that the white title on a dark
  divider exports black. Those shapes stop being placeholders, and per-run
  formatting within a paragraph collapses to the paragraph's first run.
- **`set_text(shape, *lines)` — never bare `setString()` on a template shape.**
  `setString` resets font, size, weight, colour and kerning to the shape's
  defaults, which on a designed slide means throwing the design away. One
  argument per paragraph. `retext` calls it for you.
- **Reuse the template's own colours** rather than inventing new ones: probe
  prints them, and a card flipped to the dark fill needs `set_char_color(shape,
  0xFFFFFF)` on its text to stay legible.
- `move(shape, x=, y=)` and `resize(shape, width=, height=)` exist because
  `Position` and `Size` are structs — `shape.Position.X = 5` silently does
  nothing.
### Retexting a chart

A template chart is just rectangles — there is no chart object to talk to.
`scale_bars` does the arithmetic, reading the baseline and the tallest bar off
the template so the result keeps the proportions the designer drew:

```python
bars = [shapes(page)[i] for i in (3, 5, 7, 9)]
scale_bars(bars, [40, 48, 29, 34])
retext(page, {"Q1": "Backend", "Q2": "Frontend"})
```

Bars grow upward, so a bar's `y` is `baseline - height`; `set_bar(shape,
baseline, height)` is the single-bar version. **Axis labels are the catch** —
they are sized for "Q1", and a word too wide for its box breaks mid-word
("BAC / KEND"). Keep them short, or widen the box.

### Sizing a badge to new text

**A shape's `Size` never re-measures after you retext it** — not even with
`TextAutoGrowWidth`, which keeps whatever width the shape already had. A pill or
badge whose label got longer will wrap inside its old width instead of growing.
`fit_pill` relays the label and its backdrop, taking the padding from what the
template already uses so no constant has to be guessed:

```python
retext(page, {"SPEED": "INFRASTRUCTURE"})
fit_pill(doc, page, pill, label)           # pill and label both grow
```

`fit_width(doc, page, shape)` is the primitive under it, and `text_width` lays
a string out on a throwaway shape when you want the number without resizing.

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
  still renders correctly, but
  **do not loop a deck through repeated `.pptx` saves.** Build in one pass and
  save once; if the work needs several editing rounds, keep the working copy as
  `.odp` and export to `.pptx` at the end.

One-time normalisation to expect in both formats: a table shape is resized to
the natural height of its rows on first load. It is stable afterwards.

## QA (required)

### Automated QA

Run this on every deck you produce, before you look at anything:

```bash
python3 scripts/qa.py out.pptx --template brand.pptx
```

It reports what a reader would see as broken and nothing else: text overrunning
its box *into* other text or off the slide, a lone word broken across lines
because its box is too narrow, wording still shared with the template, and
leftover placeholders — then runs `validate.py` on a `.pptx`. Exit code 1 means
findings. Pass `--template` for any template-derived deck; without it the
leftover-wording check cannot run. For a deck built from the house default, the
template to pass is `assets/company-template.pptx`.

Calibrated against four decks built from one brand template: silent on the
untouched template and on a carefully built deck, and it finds the defects their
authors missed — a leftover template paragraph, and pill labels breaking
mid-word.

**Text taller than its box is not itself a defect**, which is why `qa.py` does
not report it. Designed templates size a body box to one line and let real
content flow into the whitespace below; nothing is clipped and the slide is
fine. Do not "fix" that.

### Visual QA

Slides are visual. Read the renders; do not reason about coordinates.

```bash
python3 scripts/render.py deck.odp          # prints absolute PNG paths
```

Pass those paths to the view tool and look at every slide. After staring at the
generating code you see what you expected rather than what rendered, so look
fresh. What to check, most common first:

- **Text overflow, or text cut off at a box or slide edge.**
- Overlapping elements (text through shapes, lines through words)
- Footers or citations colliding with the content above
- Elements closer than 0.3", or margins under 0.5"
- Uneven gaps — cramped in one place, empty in another
- Columns or repeated cards not aligned consistently
- Low-contrast text or icons (light on light, dark on dark)
- Template decoration left behind after text replacement — e.g. a title
  underline positioned for one line when the new title wraps to two
- Leftover placeholder text

**Hidden slides are not rendered.** A deck with hidden slides produces fewer
pages than it has slides, so PNG numbers stop matching slide numbers after the
first hidden one. `dump.py` flags which slides are hidden.

`thumbnail.py` works on your finished deck too, not just on templates: one
contact sheet to scan, then per-slide renders only for the slides that look
wrong. That is a great deal cheaper than reading fifteen PNGs one at a time.

**If you cannot see images**, say so rather than claiming the deck is checked.
`qa.py` is the fallback and it covers the commonest defect — broken text — but
it cannot judge alignment, contrast, crowding or whether a chart's bars are
proportional. Run it, report what it found, and ask the user to eyeball the
slides you could not verify, naming them.

### Content QA

`qa.py` above covers this. To grep a deck for something specific of your own:

```bash
python3 scripts/dump.py deck.odp | grep -niE "lorem|ipsum|\bTODO|\[insert|platzhalter"
```

### File QA (.pptx only)

```bash
python3 scripts/office/validate.py deck.pptx
python3 scripts/office/validate.py deck.pptx --original template.pptx
```

Checks XML well-formedness, relationships, content types, unique IDs, layout
references, master/theme sharing, chart constraints, and the slide-XML defects
PowerPoint refuses to open a file over. Each failure names its fix. Exit codes:
`0` pass, `1` validation failure, `2` bad usage or unreadable file.

- **Pass `--original` for any template-derived deck.** A template may itself
  contain parts the XSD rejects; `--original` baselines against it so the
  template's own defects do not read as yours. For a house-default deck that
  baseline is `assets/company-template.pptx`.
- `--auto-repair` adds missing `xml:space="preserve"` to `<a:t>` runs with
  leading or trailing whitespace, in place.
- The slide check is a deliberate denylist of defects PowerPoint actually
  rejects, not full schema conformance — a schema-invalid deck that PowerPoint
  opens fine is reported as passing, by design.
- **There is no equivalent for `.odp`, and none is needed.** LibreOffice writes
  ODF itself. QA an `.odp` by rendering it and reading it back.

## Typography

**Font names you write are rendered by the recipient's Office, not by this
machine.** QA renders through LibreOffice, which substitutes fonts it lacks. If
the substitute has different metrics, your overflow check is wrong.

Verified on this machine:

| You write | Renders here as | Overflow QA |
|---|---|---|
| Arial | Liberation Sans | trustworthy (metric-compatible) |
| Times New Roman | Liberation Serif | trustworthy |
| Courier New | Liberation Mono | trustworthy |
| Calibri | Carlito | trustworthy |
| Cambria | Caladea | trustworthy |
| Georgia, Trebuchet MS, Impact, Arial Black, Garamond, Consolas, Palatino | DejaVu family | **not trustworthy** — different widths |
| Aptos | DejaVu Sans | **avoid** — no metric-compatible substitute anywhere, and missing from older Office installs |

Calibri and Cambria are trustworthy only because `fonts-crosextra-carlito` and
`fonts-crosextra-caladea` are installed. If `fc-match Calibri` returns anything
but Carlito, treat them as untrustworthy again.

For fonts outside the trustworthy list: use them where asked, size those
containers with ~10% slack, and do not trust the preview's text fit.

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

## Editing .pptx XML directly

Only for OOXML details UNO cannot express — a UNO round trip is lossless for
everything listed above. When you do need it, **`references/pptx-xml.md`** has
the unpack/rezip recipe and the rules that keep a deck openable.

## Company template (default)

This skill ships our house template as the default. **When a deck is requested
and the user names no template of their own, build from
`assets/company-template.pptx`** — do not start from a blank document, and do
not invent a look with `references/design.md`. The user's own template always
wins: if they supply one, use theirs and ignore the default.

1. **Choose designs from the catalog.** `references/company-template.md` maps
   content to **slide numbers in the asset** — a decision table applied top to
   bottom, first match wins — and describes what each design holds. Read it
   first, pick a slide per section, and vary them per its variety rule. Do not
   eyeball the thumbnail and guess. Slide `N` in the table is the number you put
   in `build_deck`'s design list; the autolayout number `probe.py` prints is
   *not* an identifier here — every slide of the template shares it.
2. **Respect the catalog's capacities.** Each entry states how many bullets a
   body holds, how many columns, how many images. When content exceeds that,
   split the slide rather than overflow it — the numbers are there so you do not
   have to discover the limit in the render.
3. **Build with `build_deck`** against the asset — see [Building from a template
   at full fidelity](#building-from-a-template-at-full-fidelity). Pass the slide
   numbers you picked in final order and retext from its `fill` callback, by the
   template's own wording. It duplicates rather than rebuilds, so this template's
   bullets, inline colour and think-cell objects survive; `clone_slide` loses all
   three here.
4. **QA against the asset:**
   `python3 scripts/qa.py out.pptx --template assets/company-template.pptx`, and
   `validate.py --original assets/company-template.pptx` for a `.pptx`.

### Maintaining the default

The catalog and the asset are generated from the real template **once, on
internal infra** — never send company slides to a hosted model. The pipeline:

```bash
python3 scripts/catalog_extract.py company.pptx -o catalog/   # structure + renders, no model
```

That writes one `catalog/slide-NN/` per layout, each with a `render.png` and a
`structure.json` (slots, geometry, fonts). Then, on the internal vision model:

1. Feed each `catalog/slide-NN/` to the model with
   `references/gemma-catalog-prompt.md` — a per-layout pass, then an
   aggregation pass that produces the decision table.
2. Fold the output into `references/company-template.md` (skeleton:
   `references/company-template.example.md`) and **review every row against the
   renders** — the model overstates capacity, and a human catches it here.
3. Copy the template to `assets/company-template.pptx`.
4. Calibrate: build a few real briefs, QA them against the asset, and correct
   wrong slide picks in the decision table.

The asset, the catalog and any renders are brand-internal — commit them to the
internal fork only, never to a remote outside the company.

## Templates

For our house template, see [Company template](#company-template-default); the
steps here are for any *other* template a user brings.

```bash
python3 scripts/thumbnail.py template.odp     # contact sheet, labelled by master
python3 scripts/dump.py template.odp          # which slots each layout offers
python3 scripts/probe.py template.odp         # shape indices, geometry, fonts
```

Map each content section onto a template slide and **vary the layouts** — not
every section belongs on title-and-bullets. Then build the deck with
`clone_slide` and `retext` — see [Adapting a template
deck](#adapting-a-template-deck) for the whole recipe.

- **Template slots ≠ your content.** If the template shows 4 team members and
  you have 3, delete the 4th member's whole group — image and text boxes — not
  just its text, then check the render for orphaned visuals. `page.remove(shape)`
  deletes; do it before you index into `shapes(page)` again.
- To reuse a template's icon or image, clone or duplicate a slide that has it.
- `.otp`/`.potx` templates open and save like any deck; keep the extension.
- Legacy `.ppt`: `python3 scripts/convert.py old.ppt new.pptx` first.
- A template installed system-wide is usually read-only. Copy it before editing;
  `load_document` says so rather than failing obscurely.

## Design

Building a deck from scratch, or restyling one? Read
**`references/design.md`** before choosing anything: palette, layout variety,
spacing, and the list of tells that make a deck look generated (accent lines
under titles, decorative colour bars, text-only slides).

Filling a brand template does not need it — the template already made those
decisions, and your job is to follow them. That includes the house default: when
you build from `assets/company-template.pptx`, follow the catalog, not
`design.md`.

## Notes on this environment

- **apt only.** Never run `pip` or `npm`, in a script or in a shell. Everything
  needed is installed; if something is missing, say so and stop.
- **`from com.sun.star... import ...` only works after `import uno`.** The
  `com` package is synthesised by the UNO bridge, so put those imports *below*
  your `office.*` imports or you get `ModuleNotFoundError: No module named 'com'`.
  Same for `createInstance`: services like `drawing.ShapeCollection` come from
  the remote `ServiceManager`, not from `doc`.
- The locale is German (`de_DE.UTF-8`), so LibreOffice renders numbers and dates
  German-style (`1.234,50`, `55,2%`). German content is normal.
- Macros are never used here; documents open with macro execution disabled.
- Each script starts its own headless LibreOffice with a private profile and
  port, so parallel runs — and a LibreOffice the user has open — never collide.
