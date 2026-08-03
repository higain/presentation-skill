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
| **Read** a deck | `scripts/dump.py deck.odp` |
| **Survey** a template's layouts | `scripts/thumbnail.py template.odp` |
| **Inspect** shapes to edit them | `scripts/probe.py template.odp` |
| **Create** a new deck | `office.impress` — see [Building a deck](#building-a-deck) |
| **Build in the house style** | no template supplied — see [Company template](#company-template-default) |
| **Adapt** a template deck | `office.template` — see [Adapting a template deck](#adapting-a-template-deck) |
| **Edit** an existing deck | `impress_document(path)` — saves in place, same format |
| **Convert** formats | `scripts/convert.py in.odp out.pptx` |
| **Render** for visual QA | `scripts/render.py deck.odp` |
| **Check** a finished deck | `scripts/qa.py deck.pptx --template brand.pptx` |
| **Schema-check** a `.pptx` | `scripts/office/validate.py deck.pptx` |

**Two rules that override anything below:**

1. **Keep the format the deck arrived in.** Editing an `.odp` produces an
   `.odp`. Convert only when the user asks for the other format.
2. **A new deck defaults to `.pptx`** unless the user asks for `.odp` or names
   an output path that says otherwise.

**And a default for content:** when a deck is requested with no template of the
user's own, build it from our house template — see [Company
template](#company-template-default) — rather than inventing a look.

## The recipe

**This is the commonest job — a `.pptx` on the house template — and this is the
whole of it.** Follow these six steps in order. Everything further down is
detail on one of them.

```bash
# 1. See what the template offers.
scripts/thumbnail.py assets/company-template.pptx
scripts/probe.py assets/company-template.pptx --slides 8,13
```

**2. Pick a design per section** from `references/company-template.md` — a
decision table, top to bottom, first match wins. Its numbers are the slide
numbers you pass to `build_deck`. Content the table has no design for — a
diagram, a chart, a comparison you want to draw — goes on **design 8 stripped
to a canvas**; see [Drawing on a slide](#drawing-on-a-slide).

**3. Write one build script.** One `fill` per page, retexting *and* drawing.
Never reopen the finished file to add more: that is a second `.pptx` save, it
grows the masters, and it is how shapes end up layered over template text
nobody replaced.

```python
import sys; sys.path.insert(0, "/absolute/path/to/presentation/scripts")
from office.deckbuild import build_deck
from office.template import retext
from office.impress import mm
from office import draw as D

def fill(doc, page, slot, design):
    if slot == 0:
        retext(page, {"Potentielle Navigatoren": "How we got here"})
    if slot == 1:                                   # a drawn slide
        area = D.canvas(page, keep=(0, 1, 2, 3))    # title, topic, source
        body, foot = area.cut_bottom(mm(14))        # a PAIR -- unpack it
        D.venn(doc, page, body,                     # sizes itself to the box
               [0x009EE0, 0xE8730C, 0x6B3FA0],
               labels=["Private data", "Untrusted content", "External comms"],
               centre=0xC00000, centre_text="Exfiltration")
        D.text(doc, page, foot, "Any two is safe by design",
               size=12, align="center")             # colour picked for the bg

build_deck("assets/company-template.pptx", "out.pptx", [7, 8, 13], fill=fill)
```

```bash
# 4-6. Build, check, render — one command, and it prints the PNG paths.
scripts/make_deck.py build.py out.pptx --template assets/company-template.pptx
```

**Give this command a timeout of at least five minutes.** It opens LibreOffice
several times and a fifteen-slide deck runs to a minute or two. A shell that
gives up at the default kills it *after* the build and *before* it prints where
the renders are — which is how a deck gets called finished having never been
looked at. It writes `<deck>-report.md` and `render/` beside the deck either
way, and names both in its first line, so a run that did get cut off is
recovered by reading those rather than by building again.

**5. Fix every finding it reports, then run it again.** Exit 1 means findings,
and they are also left on disk in `<deck>-report.md` beside the deck. Leftover
template wording, a stray template picture, text running off the slide and
black-on-navy text are defects, not noise. **Editing the build script is not
fixing them** — nothing has changed until the build has re-run and the report
comes back clean.

**6. Open the PNGs and look at them.** This is not optional and nothing
substitutes for it. `make_deck.py` cannot see a slide: it will not tell you that
a diagram is off-centre, that two shapes collide, that a slide is half empty, or
that one is still in the template's language. If you cannot see images, say so
plainly and name the slides you could not verify — do not report the deck as
checked.

**Every slide needs something to look at.** A deck of twelve bullet lists is a
failed deck even when every check passes; see [Composition](#composition).

## Scripts

**Every path on this page is relative to this skill's own directory, not to
where you are working.** Find it once — the directory holding the `SKILL.md`
you are reading — and prefix it, or `cd` there. `assets/` and `references/` are
inside it. (`build_deck` resolves a template against it for you, so
`build_deck("assets/company-template.pptx", ...)` works from anywhere; nothing
else does.)

**Run each script directly** — every one carries a shebang and the executable
bit, so `scripts/dump.py deck.odp` is the whole invocation. If a copy of this
skill ever lands without the executable bit (an archive, a checkout with
`core.fileMode=false`), `python3 scripts/dump.py deck.odp` still works and does
the same thing.

| Script | What it does |
|---|---|
| `scripts/dump.py deck.odp [-o out.md]` | Deck to markdown: per slide the layout, master, text, tables, speaker notes, empty placeholder slots and a hidden flag. Works on every supported format |
| `scripts/probe.py deck.odp [--slides 4,7-9]` | Every shape's **index**, position, size, fill, line and character formatting, plus a per-slide reading-order block grouping shapes by column. `dump.py` says what a slide *says*; this says what it *is*. Run it before editing any template slide |
| `scripts/thumbnail.py deck.odp [prefix]` | One labelled contact sheet of all slides, for picking template layouts. Output is named after the deck, so two decks never overwrite each other |
| `scripts/render.py deck.odp [-o dir] [--dpi N] [--slides 3,7-9]` | One PNG per slide; prints their absolute paths. **This is visual QA** |
| `scripts/make_deck.py [build.py] deck.pptx --template brand.pptx` | **Run this last.** Runs the build script, QAs, schema-checks, renders every slide, and writes `render/` and `<deck>-report.md` beside the deck. Exit 1 on any finding |
| `scripts/qa.py deck.pptx [--template brand.pptx] [--slides 3,7-9]` | The checks alone, without the renders: overflow, mid-word breaks, off-slide shapes, leftover template wording and pictures, placeholders, markdown, contrast, occlusion — plus `validate.py` on a `.pptx`. Exit 1 on any finding |
| `scripts/convert.py in.odp out.pptx` | Convert between `.odp`/`.pptx`/`.otp`/`.potx`/`.ppt`/`.pdf` |
| `scripts/office/validate.py deck.pptx [--original src.pptx] [--auto-repair]` | Schema, relationship, content-type, chart and slide checks. **`.pptx` only** — see [File QA](#file-qa-pptx-only) |
| `scripts/catalog_extract.py template.pptx -o catalog/` | **Build-time only, not per deck:** per-layout dossiers (render + structure) for authoring the house catalog. See [Maintaining the default](#maintaining-the-default) |
| `tests/roundtrip_report.py check fixture.odp` | Re-proves the fidelity claims below after a LibreOffice upgrade |

The `office` modules your build script imports:

| Module | What it gives you |
|---|---|
| `office.impress` | open, save, slides, shapes, placeholders, notes, `mm()`/`pt()` |
| `office.template` | `retext`, `set_text`, `find_shapes`, `clone_slide`, `fit_pill`, `shrink_to_fit`, `scale_bars`, `overflowing`, `colliding` — editing what a template slide already has |
| `office.draw` | `rect`, `ellipse`, `line`, `arrow`, `polygon`, `chevron`, `text`, `chip`, `bullets`, `table`, `venn`, `canvas`, `Box`, `restyle` — **putting new things on a slide**, and `Box` to lay it out; see [Drawing on a slide](#drawing-on-a-slide) |
| `office.deckbuild` | `build_deck` — a whole deck from a template, in any order |

Three references, read only when the job calls for them:
`references/design.md` (inventing a look — palette and motif, for a deck with no
template), `references/company-template.md` (our house catalog — which
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

### Where temporary files go

Every script here needs scratch space: a disposable LibreOffice profile per run,
a PDF between a deck and its PNGs, an unzipped `.pptx` for the validator.
Importing anything from `office` picks a writable directory once and points both
Python and LibreOffice at it — you do not normally have to think about it.

**If `/tmp` is not writable** — a locked-down or containerised machine — it falls
back to `.presentation-tmp/` in the working directory and removes it on exit.
Without that, LibreOffice aborts with *"User installation could not be
completed"* and converts nothing, so every render and every QA run fails at once
with no obvious cause. To put the scratch space somewhere specific:

```bash
PRESENTATION_TMPDIR=./build/tmp scripts/make_deck.py build.py out.pptx
```

That path must be writable; if it is not, the run stops immediately and says so,
rather than failing later inside LibreOffice.

## Reading a deck

```bash
scripts/dump.py deck.odp
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

page = add_slide(doc)                      # blank; place everything yourself
box = add_textbox(doc, page, mm(20), mm(20), mm(240), mm(25), "Ergebnisse",
                  size=40, weight=150, color=0x1E2761)
add_image(doc, page, "chart.png", mm(20), mm(60), mm(120), mm(80))

save_as(doc, "deck.pptx")                  # the extension picks the filter
doc.close(False)
```

**Geometry is in 1/100 mm, not points or EMU.** Use `mm()` and `pt()` rather
than raw numbers — a value that looks like points is 25× too small and lands
off-slide with no error at all.

**Read the slide size; never hardcode it.** `slide_size(doc)` is the only answer
that is right for the deck in front of you. A LibreOffice default page is
28000 × 15750, but a brand template is usually its own size — the house
template is **32001 × 18000**, 14% wider, enough to shift every offset you
computed against the default.

**Style after the text, never before.** `setString()` on a shape re-inherits the
drawing default (Arial 18, black) and silently discards character formatting set
before it — this is why font sizes "don't stick". Passing `size=`/`color=`/
`weight=` to `add_textbox` gets the order right for you; if you call
`setString()` yourself, follow it with `office.draw.restyle(shape, ...)`.

`add_slide(doc)` gives you a blank slide. It takes an optional `layout=` —
Impress *autolayout* numbers, `0` title+subtitle, `1` title+content, `19` title
only, `20` blank (the default) — which have **nothing to do with the design
numbers in a layout catalog.** Catalog "design 20" is a specific template slide;
`layout=20` is "no placeholders". Confusing the two puts a densely filled
template slide where you wanted an empty one.

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
from office import draw as D

def fill(doc, page, slot, design):
    if design == 7:
        retext(page, {"Formelemente, Navigator": "Kapitel 1: Vorgehen"})
    if slot == 3:
        D.rect(doc, page, mm(20), mm(60), mm(80), mm(40), fill=0x009EE0)

build_deck("assets/company-template.pptx", "out.pptx",
           [7, 8, 13, 20, 1],          # catalog design numbers, in final order
           fill=fill)                   # called per page, before the save
```

`designs` may repeat a design. **`fill(doc, page, slot, design)` is where all
the content goes** — retexting *and* drawing. It runs after the originals are
gone, so `slides(doc)` is already the deck you asked for, and it hands you `doc`
because every shape constructor needs it.

**Do not reopen the saved file to add shapes.** That is a second `.pptx` save
(the masters grow — see [What a round trip
preserves](#what-a-round-trip-preserves)), and because the second pass sees a
deck you have already stopped thinking about, it is how drawn shapes end up
layered on top of template text nobody replaced. Everything happens in `fill`.

`build_deck` QAs its own output and prints what it finds; pass `qa=False` to
silence that.

Why it beats `clone_slide` on a template: measured against the delta asset over
designs 7, 8, 13, 20 and 1, diffing each result against its source render, mean
RMS **7.8 versus 22.8**, better on every slide. Bullet markers, per-run colour
inside a paragraph and think-cell's OLE data all survive here and do not survive
a rebuild — 6 think-cell references kept versus none. Same output size, same
validator findings.

Limits: `.pptx` only, since the reorder is an OOXML step, and it builds a *new*
deck from a template rather than editing one in place. For those, use
`clone_slide`.

## Drawing on a slide

A template gives you designed slides. It does not give you a Venn diagram, an
attack flow, a comparison matrix or a bar chart of your numbers — and the
difference between a deck that lands and a deck of bullet lists is almost always
whether those got drawn. `office.draw` is the toolbox:

```python
from office.impress import mm
from office import draw as D

D.rect(doc, page, mm(20), mm(40), mm(60), mm(30), fill=0x009EE0, radius=200)
D.ellipse(doc, page, mm(50), mm(55), mm(15), fill=0xE8730C, transparence=60)
D.text(doc, page, mm(20), mm(40), mm(60), mm(30), "Phase I", "eight weeks",
       size=18, color=D.WHITE, weight=D.BOLD, align="center", valign="center")
D.arrow(doc, page, mm(85), mm(50), mm(8), mm(10))
D.chevron(doc, page, mm(20), mm(80), mm(50), mm(14), fill=0x009EE0)
D.table(doc, page, mm(20), mm(100), [mm(60), mm(40), mm(40)],
        [["Control", "Cost", "Stops it?"], ["Prompt filter", "low", "no"]],
        row_h=mm(10), head_fill=0x009EE0)
D.venn(doc, page, mm(160), mm(90), mm(40), [0x009EE0, 0xE8730C, 0x6B3FA0],
       labels=["Private data", "Untrusted content", "External comms"],
       centre=0xC00000, centre_text="Exfiltration")
```

Full list: `rect`, `ellipse`, `line`/`hline`/`vline`, `polygon`, `arrow`,
`chevron`, `text`, `chip`, `bullets`, `table`, `venn` (+ `venn_geometry`,
`pair_spots`), `canvas`, `Box`/`page_box`, `restyle`, `walk`/`walk_page`,
`est_lines`/`fit_size`/`fit_lines`, plus `move`, `resize` and `set_text`
re-exported from `office.template`.

- **Ellipses take a centre and radii**, not a bounding box.
- **A Venn is `D.venn`, never three `D.ellipse` calls.** It places the circles so
  they actually overlap, sizes the figure to the `Box` you pass as `box=`, and
  puts each label outside its own lobe. Hand-placed circles come out tangent, or
  with their labels on top of each other and off the bottom of the slide — every
  hand-rolled Venn across six test builds was wrong, and every `D.venn` one was
  right. `qa.py` cannot see the difference; it is a rendering defect and it needs
  eyes. `D.venn_fit(box)` gives you the geometry alone.
- **`h` is optional on `D.text`** — omit it and the box is sized to the text at
  the width you gave. `D.text(doc, page, box.x, box.y, box.w, "Heading", size=20)`.
- **`color` on `D.text` defaults to whatever reads on the background**: white on
  a dark divider, black on a white body slide. Pass it explicitly to override.
  Do not assume black — a headline drawn in black onto navy is the right words
  in the right box, invisible, and it renders clean through every check that
  cannot see a slide.
- **No markdown in slide text.** `~~struck~~` and `**bold**` render as literal
  tildes and asterisks. Style the run instead: `weight=D.BOLD`, `italic=True`,
  `strike=True`. `qa.py` flags these.
- **Colours are ints** (`0x009EE0`). Take them from the template — `probe.py`
  prints every fill and line colour — rather than inventing a palette. Check one
  against its backdrop with `D.contrast_ratio(a, b)` before committing to it;
  under 3.0 it is not readable from a seat. `D.readable_on(bg)` picks for you.
- `D.table` draws rectangles and text rather than a real table shape: it takes
  per-cell colour without fighting a table style and survives a round trip
  unchanged. Use a real table only if the recipient must edit it as one.
- `est_lines` guesses how many lines a string will wrap to, for laying out
  before the document can measure. `office.template.text_overflow` measures for
  real — trust that one. `fit_size(s, width)` returns the largest point size at
  which `s` fits on one line, and `fit_lines(s, w, h)` returns `(size, lines)`
  wrapped on spaces — use them instead of guessing a size for a label whose
  length you do not control, which is where mid-word breaks come from.

### Lay out with `Box`, not with arithmetic on millimetres

**Every drawing call takes a `Box` where its geometry goes** — pass it whole,
there is nothing to unpack. Ask an area for the pieces you want, and the pieces
are right by construction:

```python
area = D.canvas(page, keep=(0, 1, 2, 3))     # the free area, as a Box
body, foot = area.cut_bottom(mm(12), gap=mm(6))
for col, step in zip(body.cols(4, gap=mm(8)), steps):
    D.rect(doc, page, col, fill=0xEEEEEE, radius=300)
    D.text(doc, page, col.inset(mm(5)), step, size=14, align="center")
```

`rect` · `text` · `chip` · `bullets` · `arrow` · `chevron` · `ellipse` ·
`table` · `venn` all take one. `ellipse` and `venn` read it as the area to
fill, so neither needs a radius worked out by hand. The four-number form still
works, and so does `*box` — use whichever reads better.

Returning **one** `Box`: `top(h)` · `bottom(h)` · `left(w)` · `right(w)` ·
`inset(all=, x=, y=, left=, top=, right=, bottom=)` · `square(scale)` for
anything drawn from a radius. Returning a **list**: `cols(n, gap)` ·
`rows(n, gap)` · `grid(rows, cols, gap)`. Returning a **pair**:
`cut_top/cut_bottom/cut_left/cut_right(size, gap)`. Plus `.centre`, `.cx`,
`.cy`, `.x2`, `.y2`. `D.page_box(page)` is the whole slide, for when you are not
on a canvas.

**The `cut_*` family returns two boxes and must be unpacked.** They come in
reading order, so which half is the band changes with the direction:

```python
head, body = area.cut_top(mm(14))      # the band is FIRST
body, foot = area.cut_bottom(mm(14))   # the band is SECOND
```

Getting that backwards is silent and it wrecks the slide: you fill the whole
panel instead of a 14 mm strip of it, nothing lands off the slide, and no text
overflows. **When you want only one of the two, name it** — `area.top(mm(14))`,
`area.bottom(mm(14))`, `area.left(mm(60))`, `area.right(mm(60))` return a single
`Box`, go straight into a drawing call, and have no order to get wrong.
Reach for those first; keep `cut_*` for when you genuinely need both halves.
`qa.py` reports the wreckage as an occlusion finding, but only after you look.

**A run of `mm(30 + i * 70)` is the tell.** It assumes a slide width, drifts out
of alignment the moment a column is added or removed, and leaves the content
sitting in the top half of the slide with a dead band underneath — the
commonest composition defect in a built deck, and one no automated check
reports. Slicing a `Box` consumes all of it, so the content fills the area the
template actually left.

### The canvas pattern

**When the catalog has no design for what you need, do not force the content
into the nearest one and do not start from a blank slide.** Take a designed
content slide, strip it back to its title and topic label, and draw on what is
left. You keep the master, background, chrome, fonts and margins — everything
that makes it look like the house deck — and you get an empty area to work in.

```python
def fill(doc, page, slot, design):
    area = D.canvas(page, keep=(0, 1, 2, 3))   # design 8: title, topic, source
    retext(page, {"Potentielle Navigatoren": "Why it cannot just be patched"})
    left, right = area.cols(2, gap=mm(10))     # draw in what the body left
```

`canvas(page, keep=(...))` removes every shape whose index is not in `keep`;
`canvas(page, remove=(...))` removes only the ones you name. **Run `probe.py` on
the design first** — the indices differ per design and guessing them deletes the
title.

**It returns the free area as a `Box`**, measured from the shapes it removed —
the body placeholder's own rectangle, which is the designer's answer to where
content goes. Use it rather than hardcoding an area: it is exact, and it is
right on a design you have not probed.

On the house template **design 8 is the canvas**: `[0]` a think-cell stub, `[1]`
title, `[2]` topic label, `[3]` source line, `[4]` the sample body text and
`[5]`+ the navigator chevrons. Keep `(0, 1, 2, 3)` and `canvas` hands you
`Box(x=2031, y=3545, w=27972, h=12157)`. See `references/company-template.md`.

## Composition

**This applies to every deck, template or not.** A brand template decides the
palette, the fonts and the chrome; it does not decide whether your slide is
worth looking at. Choosing the palette for a deck that has no template is a
separate job — `references/design.md`.

**Every slide needs something to look at** — a diagram, chart, table, image,
icon or drawn shape. A slide that is a title and a bullet list is forgettable,
and a deck of twelve of them has failed even when every automated check passes.
That is the single commonest way a build goes wrong now that `office.draw`
exists: reaching for `bullets()` when the content wanted a diagram.

Patterns worth reaching for, roughly in order of how often they earn their keep:

- **Comparison columns** — before/after, pros/cons, option A vs B
- **Timeline or process flow** — numbered steps, chevrons, arrows
- **A table** — when the content really is a matrix, say so with one
- **Large stat callouts** — 60–72pt numbers with small labels beneath
- **Icon + text rows** — icon in a filled circle, bold header, description below
- **2×2 or 2×3 grid** of content blocks
- **Two-column** — text left, illustration right
- **A diagram** — Venn, quadrant, flow, hierarchy; `office.draw` has the parts

Then:

- **Vary the layout.** Not every section belongs on title-and-bullets. Repeating
  one design down the whole deck reads as a form someone filled in.
- **Size contrast.** Titles well above body text; see [Typography](#typography).
- **Left-align body text.** Centre titles and short labels only.
- **Hold one spacing rhythm** — pick a gap and keep it. Minimum margins about
  0.5", 0.3–0.5" between blocks.
- **Fill the area you were given.** Content in the top two-thirds with an empty
  band underneath is the signature of hardcoded offsets, and it is invisible to
  `qa.py` — nothing overflows, nothing collides, the slide is just badly placed.
  Lay out from the `Box` `D.canvas` returns; if a band is genuinely meant to be
  empty, it should be at least as wide as the content, not a leftover.
- **Contrast on everything**, icons as well as text. A card flipped to a dark
  fill needs `set_char_color(shape, 0xFFFFFF)` on its label.
- **Never ship text that overflows its shape** — shrink it, split it, or make
  the shape bigger. `qa.py` finds these.
- **Never draw an accent line under a title**, and **never add decorative colour
  bars or stripes** — header/footer bars, sidebar stripes, thin rules along a
  card edge. They are the clearest tell of a generated deck. To set a card apart
  use a background tint or an icon.

## Adapting a template deck

The common job: a brand deck arrives with one slide per layout, and you have to
produce fifteen slides of your own content in that design.

```bash
scripts/thumbnail.py brand.pptx        # which layout is which slide
scripts/probe.py brand.pptx            # shape indices + geometry + fonts
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

**Text inside groups.** A designed slide keeps a lot of its wording inside
grouped shapes, and `shapes(page)` cannot see any of it. On the house template
**11 of 21 designs carry text inside groups**; on the phase-plan design 13, 17
shapes are top-level and 13 more sit inside five groups, 7 of them holding the
phase labels and durations. Design 10 is worse: 16 grouped text shapes against
11 top-level, so most of the slide is invisible from the top level.
`probe.py` lists them indented, addressed `[6.2]` for the third shape in the
group at `[6]`, and
`retext`, `find_shapes`, `shape_by_text` and `qa.py` all reach them. Only
`shapes(page)[i]` does not — one more reason to retext by wording rather than by
index.

**Excerpts in `probe.py` are cut short with `…`.** A `retext` key is a
substring, so copy a short distinctive fragment rather than a whole line. A
trailing ellipsis is stripped if you do paste one.

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
  argument per paragraph, and each line keeps the formatting of the template
  paragraph in the same position — a designed body that opens with a coloured
  heading and continues in body text stays that shape. `retext` calls it for
  you.
- **A `.pptx` cannot hold two text colours in one shape.** The export discards
  per-paragraph character colour and writes a single colour for the whole
  shape. Measured on the house template: three paragraphs written red, blue and
  green come back all one colour, while their per-paragraph **sizes and weights
  survive intact**. An `.odp` keeps all of it. So a heading in one colour above
  body text in another has to be **two shapes** in a `.pptx` — `office.draw.text`
  makes that cheap. The template's own multi-coloured bodies were authored in
  PowerPoint and are fine until you retext them.
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

### A title too long for the slot it was drawn for

Going the other way — the box is fixed and the words are too many — shrink the
type rather than move the box:

```python
retext(page, {"Formelemente, Navigator": "The Lethal Trifecta in Agentic AI"})
shape = find_shapes(page, "Lethal Trifecta")[0][1]
shrink_to_fit(doc, page, shape, max_lines=1)     # 60pt -> 47pt, still one line
```

**`max_lines=1` is the one that matters on a decorated divider.** A headline box
is drawn for a set number of lines at a set size; a longer title does not clip,
it flows on over whatever the *master* has behind it — and master decoration is
not a shape on the page, so nothing in `qa.py` can see the result. On the house
dividers that decoration is a large triangle, and the second line comes out with
a line struck through it. Budgets per design: `references/company-template.md`.

If the size it settles on looks wrong, the title is too long for that design
rather than too big — shorten it, or build the slide on a canvas.

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

**A deck is not finished until this has run and you have looked at the
renders.** Every deck that shipped broken had skipped one of the two.

### Automated QA

`make_deck.py` runs all of this and renders too — prefer it. `qa.py` is the same
checks without the renders:

```bash
scripts/make_deck.py build.py out.pptx --template brand.pptx   # do this
scripts/qa.py out.pptx --template brand.pptx                   # checks only
```

`make_deck.py` writes both its renders (`render/`) and its findings
(`<deck>-report.md`) **next to the deck**, not next to wherever you ran it. If a
build ends between "here is what is wrong" and "fixed it", that report is what
says so; a clean run overwrites it with a clean one.

It reports what a reader would see as broken and nothing else:

- text overrunning its box *into* other text or off the slide, and a lone word
  broken across lines because its box is too narrow;
- **two blocks of text printed over each other** — a caption placed where a
  diagram's labels already are, a Venn label on the note beneath it, a title
  long enough to reach the list beside it. Measured on the ink rather than on
  the boxes, so a big text box holding two words does not report;
- **a shape drawn past the edge of the slide**;
- wording still shared with the template, and leftover placeholders;
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

then runs `validate.py` on a `.pptx`. Exit code 1 means findings. Pass
`--template` for any template-derived deck; without it neither the
leftover-wording nor the stray-picture check can run, and a bleed the template
does on purpose cannot be told from one you introduced. For a deck built from
the house default, the template to pass is `assets/company-template.pptx`.

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

### Visual QA

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
scripts/dump.py deck.odp | grep -niE "lorem|ipsum|\bTODO|\[insert|platzhalter"
```

### File QA (.pptx only)

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
  baselines against it so the template's own defects do not read as yours. For a
  house-default deck that baseline is `assets/company-template.pptx`.
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
   content to **design numbers** — the slide numbers in the asset — as a
   decision table applied top to bottom, first match wins, and describes what
   each design holds. Read it first, pick a design per section, and vary them
   per its variety rule. Do not eyeball the thumbnail and guess.

   **"Design 8" always means the catalog's design 8**, which is the number you
   put in `build_deck`'s list. It is never an Impress autolayout number: the
   autolayout `probe.py` prints is `20` for every slide in this template and
   identifies nothing.
2. **Respect the catalog's capacities.** Each entry states how many bullets a
   body holds, how many columns, how many images. When content exceeds that,
   split the slide rather than overflow it — the numbers are there so you do not
   have to discover the limit in the render.

   **Take `retext` keys from the entry's *Retext keys* line, never from
   memory.** They differ per design, and reusing design 8's
   `"Potentielle Navigatoren"` on design 5 or 12 is the single commonest way a
   build script dies. A slot that is absent from the list is *empty* in the
   template — `retext` cannot match an empty shape, so write it with
   `set_text(shapes(page)[i], ...)` or draw your own.
3. **Anything the table has no design for — a diagram, a chart, a comparison you
   want to draw — goes on design 8 stripped to a canvas.** Do not force it into
   the nearest design's bullet list, and do not start from a blank slide. See
   [The canvas pattern](#the-canvas-pattern).
4. **Build with `build_deck`** against the asset — see [Building from a template
   at full fidelity](#building-from-a-template-at-full-fidelity). Pass the design
   numbers you picked in final order, and retext *and draw* from its `fill`
   callback. It duplicates rather than rebuilds, so this template's bullets,
   inline colour and think-cell objects survive; `clone_slide` loses all three
   here.
5. **Check and look:**
   `scripts/make_deck.py build.py out.pptx --template assets/company-template.pptx`,
   then open every PNG it prints.

### Maintaining the default

The catalog and the asset are generated from the real template **once, on
internal infra** — never send company slides to a hosted model. The pipeline:

```bash
scripts/catalog_extract.py company.pptx -o catalog/   # structure + renders, no model
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
scripts/thumbnail.py template.odp     # contact sheet, labelled by master
scripts/dump.py template.odp          # which slots each layout offers
scripts/probe.py template.odp         # shape indices, geometry, fonts
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
- Legacy `.ppt`: `scripts/convert.py old.ppt new.pptx` first.
- A template installed system-wide is usually read-only. Copy it before editing;
  `load_document` says so rather than failing obscurely.

## Design

**No template — inventing a look?** Read **`references/design.md`** before
choosing anything: how to pick a palette that fits the content, colour
dominance, dark/light structure and committing to one motif.

Filling a brand template does not need it: the template already chose the
palette, the fonts and the chrome, and your job is to follow them. That includes
the house default — build from `assets/company-template.pptx` and follow the
catalog, not `design.md`.

**[Composition](#composition) is not part of that split.** What goes on a slide,
how much variety a deck needs and the visual-per-slide rule apply to a template
build exactly as much as to a designed one — a brand template will happily let
you ship twelve identical bullet slides.

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
