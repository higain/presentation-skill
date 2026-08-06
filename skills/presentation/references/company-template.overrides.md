# delta catalog — prose overrides

Everything here is **judgement that no script can measure**, spliced into the
generated catalog by `catalog_build.py --overrides`. Each `## anchor` section is
inserted verbatim at that point in `references/company-template.md`; the rest of
that file is generated from `catalog/` and must not be edited by hand — a
rebuild overwrites it.

Anchors: `header`, `after-decision-table`, `before-design-entries`,
`design-entries-intro`, `footer`, and `design-N` for any design. An anchor that
does not land prints a warning and shows up as a ⚠ line at the foot of the
catalog, so prose can never go quietly missing.

Blocks **add to** the generated text at that point rather than replacing it, so
write them as additions and do not restate what the generator already says.

## header

**Asset detail the generator does not measure.** The deck is 16:9. Its masters
carry roles: `AT_ST` is the content master, `Section_dark` the dividers, and
`Default 36` belongs to design 4 alone. The slide size matters more than it
looks — at 32001 × 18000 rather than LibreOffice's 28000 × 15750, every offset
computed against the default lands about 14% too far left. `build_deck`
duplicates rather than rebuilds, which is what keeps this template's bullets,
inline colour and **think-cell objects** alive; `clone_slide` loses all three —
the recipe is in `SKILL.md` and the detail in `references/templates.md`.

**Design 20 is a real, densely filled two-column slide, not a blank one.** Every
slide of this deck reports `layout: 20`, and that coincidence has been read as
"design 20 is the empty layout" more than once.

**Every retext key of a design is a key you have to replace, not just the
title.** Design 8's `[2]` is the label `"Navigator (1/2)"`, and two runs
replaced `[1]` on every slide and `[2]` on none — shipping that German label on
13 of 15 and 9 of 10 finished English slides. Read the *Retext keys* line of the
design you are using and account for all of them: replace the ones you want, and
`strip_annotations` or draw over the ones you do not.

**The footer is a page property, not a shape and not the master.** `Gegenstand
| Titel | Ort | Datum` renders along the bottom of every content slide. It is a
`ftr` placeholder that Impress exposes as `page.FooterText`, so `walk_shapes`
never reaches it and neither does `retext`, `find_shapes` or anything that
searches the slide's text.

**It is German, and a deck in any other language must set or hide it.** Four
runs shipped it under English slides — every slide of every deck — and one of
them searched the shapes for the string, found nothing, and concluded the
reviewer that had flagged it was wrong. Set it once for the whole deck, or per
page in `fill`:

```python
build_deck("assets/company-template.pptx", "out.pptx", designs, fill=fill,
           footer="The Lethal Trifecta · Board briefing · 2026")   # None hides it

from office.template import set_footer, footer_text
set_footer(page, "…")          # per page, inside fill
```

`qa.py` reports it when it is still the template's, so this is a finding now
rather than a thing to remember.

## after-decision-table

The canvas row's `keep=(0, 1, 2, 3)` on design 8 is the think-cell stub, the
title, the topic label and the source line — see "The canvas pattern" in
`references/drawing.md` for how to draw in the area that frees up.

**There is no cover design and no closing design.** Design 7 is a dark section
divider pressed into service for both, and a real deck title with a subtitle and
a date does not fit it — see the divider budget below. A title slide has to be
built: design 8 as a canvas, with title, subtitle and date drawn in the area the
body text vacates. Do not leave the deck without one.

**Neither is there a design for a diagram.** That is what the canvas row at the
bottom of the table is for, and on a deck with any analytical content you should
expect to use it several times — several, not every time. `build_deck` records
the design list on the file and `qa.py` prints the spread, so a deck that is
100% canvas says so in its own report.

## before-design-entries

**Why the divider budgets are hard limits.** On designs 1, 5 and 7 the title
sits in the gap between the two edges of a big white triangle. One line lands in
the gap; a **second line crosses the triangle**, and the words come out with a
stroke drawn through them. The triangle is master decoration rather than a shape
on the slide, so nothing collides and `qa.py` reports clean.

`shrink_to_fit(doc, page, shape, max_lines=1)` comes from `office.template` and
reduces the type until the title is one line again — a 38-character title on
design 7 settles at 47pt and looks right. Do not add a second text box beside
the placeholder: that is how a title and its subtitle both end up written
through the triangle.

## design-entries-intro

Every failed `retext` in testing was a key remembered from a *different* design,
so copy the key from the entry you are actually using. Text is shown truncated
and any distinctive substring matches — `"Potentielle Navigatoren"` finds the
full string. A design's title or source slot that is **absent from its key list
is empty in the template**: `retext` cannot find an empty shape, so write it with
`set_text(shapes(page)[i], ...)` or draw your own.

## design-8

- **Watch `[2]`:** the `"Navigator (1/2)"` label is the key most often left
  behind on this design. Retext it or draw over it on every slide that uses it.

## footer

## Working with this asset

Facts established by inspecting the file itself (`probe.py`, zip inspection),
not by reading the renders. A vision model cannot see any of these.

### think-cell is embedded throughout

17 of the 21 slides carry a hidden OLE shape literally named
`think-cell data - do not delete`; the deck holds 62 `oleObject*.bin`
embeddings in total. **Do not delete these shapes**, and expect them not to
survive a UNO round trip intact — a rebuilt deck loses the think-cell linkage
even when the visible slide looks identical. If a build must keep live
think-cell charts, edit the original in PowerPoint instead of rebuilding here.

Slide 11 is the exception worth knowing: its waterfall is a genuine OOXML
`barChart` (`ppt/charts/chart1.xml`) with an embedded workbook — two series,
the lower one invisible to fake the waterfall — so it *is* editable as a real
chart. It carries no think-cell object.

### Build with `build_deck`, not `clone_slide`

**Use `office.deckbuild.build_deck` for this template.** It duplicates designs
instead of rebuilding their shapes, then sets the slide order by permuting
`<p:sldIdLst>` in the saved file. Measured over designs 7, 8, 13, 20 and 1
against their source renders: mean RMS **7.8 versus 22.8** for `clone_slide`,
better on every slide.

That matters here specifically because this template leans on three things a
shape rebuild destroys: **bullet markers**, **per-run colour inside a paragraph**
(the blue-on-black body text is everywhere), and **think-cell's OLE objects** —
6 references survive a `build_deck` build, none survive `clone_slide`. The
divider footer line survives too.

The rest of this section describes `clone_slide`'s behaviour, which still
applies if you use it — for an `.odp`, or when editing a deck in place.

### Cloning dividers with `clone_slide` — fixed, but know what it cost

Cloning a `Section_dark` divider used to produce a white slide with black text.
Two separate faults, both now fixed in `office/template.py`:

1. **The master was not carried.** A new page comes up on its neighbour's
   master, so the divider landed on `AT_ST` and lost the dark background.
   `clone_slide` now assigns `src_page.MasterPage` before copying anything.
2. **A cloned placeholder loses its direct run colour on export.** This one is
   the subtle half. The write lands — read the portion back in the same session
   and it is white — but LibreOffice's `.pptx` export discards direct run
   formatting on a *placeholder* and re-inherits the master's text style, so
   the run leaves as `<a:schemeClr val="dk1"/>` where the source had `lt1`.
   Reassigning the master, rebinding `Style` to the source master's family, and
   forcing a non-theme colour (`0xFEFEFE`) all failed to prevent it.

`clone_slide` now rebuilds such placeholders as plain `drawing.TextShape`s,
which have no style to re-inherit from, so the colour exports as an explicit
`<a:srgbClr val="ffffff"/>`. The trigger is narrow: a placeholder is flattened
only when a run's colour differs from its own presentation style's. Across this
template that is the divider titles and the blue topic label on every content
slide — the black action titles keep their placeholder.

**The blue topic label was silently going black too**, on all 18 content
slides; the same fix repairs it.

Verified against the renders: clones of slides 1, 5 and 7 now match their
sources exactly (identical near-white pixel counts, RMS ≈ 3, antialiasing
only), and slides 13 and 20 are byte-for-byte unchanged from before the fix.

The cost of flattening: those shapes are no longer placeholders, so they will
not pick up later layout changes, and per-run formatting *within* a paragraph
collapses to that paragraph's first run.

### Inline formatting in body text boxes — avoided, not repaired

The body text on these slides sits in ordinary `ooxml-rect` shapes, not
placeholders, so `clone_shape` handles it — and it loses bullet markers and
inline run colour, flattening a paragraph that mixed blue and black to one
colour. Compare a `clone_slide` copy of slide 20 against the original and the
bullets are visibly plainer.

This was never fixed in `clone_slide`; `build_deck` sidesteps it by not
rebuilding shapes at all (slide 20: RMS 2.5 via `build_deck`, 31.5 via
`clone_slide`). If you do use `clone_slide` on this template, budget for fixing
body text by hand and check the render before shipping.

### Round-trip validation noise

`validate.py` reports the untouched asset clean, but a UNO round trip of it
raises 22 findings: 6 broken `customXml/item*.xml` relationships (reported
CRITICAL) and 16 undeclared `.wmf` content types. The resulting file opens and
renders correctly regardless, so treat these as round-trip noise on this
template rather than as build failures — but they will mask any *real*
validator finding, so compare against a baseline run rather than expecting zero.

### The AT_ST text grid

Every `AT_ST` content slide (2–4, 6, 8–21) repeats the same three text slots:

| Slot | Position | Size | Role |
|---|---|---|---|
| Action title | y = 498 | 16 pt, black | The slide's **assertion** — a full sentence, not a label |
| Topic label | y = 2302 | 16 pt, blue | Short chapter/section name |
| Source line | y = 16003 | 14 pt | Empty footnote/source line, usually left blank |

The action title carries the argument; German consulting decks read as a chain
of these sentences. When retexting, keep it a full assertive sentence.

### Authoring annotations that must be deleted

The deck is a pattern collection, so several slides carry rotated coloured
flags that are instructions to the author, never content:

- Orange/green/magenta `Beispiel`, `Layout`, `Alternatives Design zum Master`
  flags — designs 6, 10, 11, 12, 15. Design 10 keeps its `Beispiel` **inside a
  group**.
- Design 4's yellow `Vorgehen:` block (four numbered build steps) and its orange
  toolbox tip box.
- Yellow highlight marks `Kundenlogo` (designs 2, 3), `Kunde`, `XX.XXX`, `XX`,
  `450` — fill-in markers, replace with real values or the customer's logo.

**`strip_annotations(page)` deletes them for you** — call it in `fill`, before
retexting:

```python
from office.template import strip_annotations, retext

def fill(doc, page, slot, design):
    strip_annotations(page)          # returns what it removed
    retext(page, {...})
```

It matches whole strings, so design 17's title *Vorgehen zum Bilden eines
Modells* and design 20's `Vorgehen` column header are correctly left alone. The
fill-in markers (`XX.XXX`, `450`) are deliberately **not** deleted — they are
values to replace, not flags to drop.

`qa.py` will not flag any of this; they are ordinary shapes carrying short
strings.

### Gaps in the collection

**No cover, no closing/contact slide, no diagram design** — see the notes under
the decision table. Designs 1, 5 and 7 are all section dividers on
`Section_dark`, which is why "deck opener" and "closing slide" both point at
one.

**Slides 8 and 9 are documentation, not designs.** They explain how the
Navigator element works and what it is for. The decision table reaches for
slide 8 as the default bulleted-content and agenda layout; that works because
its body is a plain text box, but the sample text on it is meta-commentary
about the template and must be fully replaced.

### How this catalog was first built

Generated 2026-07-28 from `delta-sammlung-21f.pptx` (21 slides):

1. `scripts/catalog_extract.py` — per-slide `render.png` + `structure.json`,
   no model involved.
2. `scripts/catalog_gemma.py --openrouter --model google/gemma-4-31b-it` —
   per-layout pass, then the aggregation pass for the decision table.

3. The measured fields — retext keys, empty slots, group counts, canvas
   indices, title budgets — were then added by hand. `catalog_extract.py`
   extracts all of them now and `scripts/catalog_build.py` assembles the whole
   file, so a regeneration no longer loses them.

To regenerate after the template changes, re-run the pipeline in README.md
(extract → model pass → `catalog_build.py`) and re-review every capacity figure
and decision-table row against the renders.
