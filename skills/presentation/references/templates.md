# Working with a template deck

For the installed pack see its `catalog.md` and the recipe in SKILL.md.
**This file is for any *other* template a user brings**, and for the
mechanics underneath both: how slides get copied, how text gets replaced, and
what breaks when you index by hand.

```bash
scripts/thumbnail.py brand.pptx        # which layout is which slide
scripts/dump.py brand.pptx             # which slots each layout offers
scripts/probe.py brand.pptx            # shape indices + geometry + fonts
```

Map each content section onto a template slide and **vary the layouts** — not
every section belongs on title-and-bullets.

- **Template slots ≠ your content.** If the template shows 4 team members and
  you have 3, delete the 4th member's whole group — image and text boxes — not
  just its text, then check the render for orphaned visuals. `page.remove(shape)`
  deletes; do it before you index into `shapes(page)` again.
- To reuse a template's icon or image, clone or duplicate a slide that has it.
- `.otp`/`.potx` templates open and save like any deck; keep the extension.
- Legacy `.ppt`: `scripts/convert.py old.ppt new.pptx` first.
- A template installed system-wide is usually read-only. Copy it before editing;
  `load_document` says so rather than failing obscurely.
- A template carries authoring notes addressed to whoever builds from it — an
  "example" flag, a customer-logo marker, a block of build instructions.
  `office.template.strip_annotations(page)` deletes them, reaching inside
  groups. The strings for an installed pack live in its `profile.json`
  (`presentation-setup` writes them); for another template pass
  `profile={"exact": [...], "prefixes": [...]}`.

## Reading the probe before you index anything

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
grouped shapes, and `shapes(page)` cannot see any of it. On one 21-design
corporate template **11 designs carry text inside groups**; on its phase-plan
design, 17 shapes are top-level and 13 more sit inside five groups, 7 of them
holding the phase labels and durations. Its worst design has 16 grouped text
shapes against 11 top-level, so most of the slide is invisible from the top
level.
`probe.py` lists them indented, addressed `[6.2]` for the third shape in the
group at `[6]`, and `retext`, `find_shapes`, `shape_by_text` and `qa.py` all
reach them. Only `shapes(page)[i]` does not — one more reason to retext by
wording rather than by index.

**Excerpts in `probe.py` are cut short with `…`.** A `retext` key is a
substring, so copy a short distinctive fragment rather than a whole line. A
trailing ellipsis is stripped if you do paste one.

`probe.py --json` emits the same survey as data when a script needs to pick
shapes itself.

## Building every slide in one pass

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
    save_as(doc, "out.pptx")                       # one .pptx save
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
  shape. Measured on a corporate template: three paragraphs written red, blue and
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

## Duplicating a slide, and why order is the hard part

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
better one**; for everything else — `.odp`, or copying between two open
documents — use `clone_slide` above, which rebuilds each shape and accepts the
losses that come with that.

## `build_deck` — full fidelity, any order

**`.pptx` from a template: reach for this first.** It duplicates the designs you
name — a real Impress duplicate, so nothing is rebuilt — drops the template
originals, and then puts the slides in your order by permuting `<p:sldIdLst>` in
the saved file. Slide order in a `.pptx` is only that list; the slide parts do
not care.

```python
from office.deckbuild import build_deck
from office.template import retext
from office import draw as D

def fill(doc, page, slot, design):
    if design == 7:
        retext(page, {"Chapter title": "Chapter 1: Method"})
    if slot == 3:
        D.rect(doc, page, mm(20), mm(60), mm(80), mm(40), fill=D.color("accent"))

build_deck("templates/acme/template.pptx", "out.pptx",
           [7, 8, 13, 20, 1],          # catalog design numbers, in final order
           fill=fill)                   # called per page, before the save
```

`designs` may repeat a design. **`fill(doc, page, slot, design)` is where all
the content goes** — retexting *and* drawing. It runs after the originals are
gone, so `slides(doc)` is already the deck you asked for, and it hands you `doc`
because every shape constructor needs it.

**Do not reopen the saved file to add shapes.** That is a second `.pptx` save
(the masters grow — see `references/qa.md`), and because the second pass sees a
deck you have already stopped thinking about, it is how drawn shapes end up
layered on top of template text nobody replaced. Everything happens in `fill`.

`build_deck` QAs its own output and prints what it finds; pass `qa=False` to
silence that. Under `make_deck.py` it is skipped automatically, because that
command runs the same check itself.

Why it beats `clone_slide` on a template: measured against a corporate asset
over five of its designs, diffing each result against its source render, mean
RMS **7.8 versus 22.8**, better on every slide. Bullet markers, per-run colour
inside a paragraph and think-cell's OLE data all survive here and do not survive
a rebuild — 6 think-cell references kept versus none. Same output size, same
validator findings.

Limits: `.pptx` only, since the reorder is an OOXML step, and it builds a *new*
deck from a template rather than editing one in place. For those, use
`clone_slide`.

## Retexting a chart

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

## Sizing a badge to new text

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

## A title too long for the slot it was drawn for

Going the other way — the box is fixed and the words are too many — shrink the
type rather than move the box:

```python
retext(page, {"Chapter title": "The quarter in three numbers and one decision"})
shape = find_shapes(page, "three numbers")[0][1]
shrink_to_fit(doc, page, shape, max_lines=1)     # 60pt -> 47pt, still one line
```

**`max_lines=1` is the one that matters on a decorated divider.** A headline box
is drawn for a set number of lines at a set size; a longer title does not clip,
it flows on over whatever the *master* has behind it — and master decoration is
not a shape on the page, so nothing in `qa.py` can see the result. On one
template's dividers that decoration is a large triangle, and the second line
comes out with a line struck through it. Budgets per design: the pack's `catalog.md`.

If the size it settles on looks wrong, the title is too long for that design
rather than too big — shorten it, or build the slide on a canvas.
