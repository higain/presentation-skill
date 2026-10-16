# Drawing on a slide — `office.draw`

**Read this before drawing anything.** A template gives you designed slides. It
does not give you a Venn diagram, an attack flow, a comparison matrix or a bar
chart of your numbers — and the difference between a deck that lands and a deck
of bullet lists is almost always whether those got drawn.

The five that matter most, if you read nothing else:

1. **Every drawing call takes a `Box`** — pass it whole, in place of the
   coordinates. `D.rect(doc, page, col, fill=…)`.
2. **`cut_*` returns a pair and must be unpacked**; `top()`/`bottom()`/`left()`/
   `right()` return one box and have no order to get wrong. Reach for those.
3. **A Venn is `D.venn`, never three `D.ellipse` calls**, and it takes the area.
4. **`D.table` takes the area and *proportional* widths.** Absolute widths are
   how a table ends up 1 mm wide with every label broken mid-word.
5. **Colours are ints from the template**, and `D.text`'s default colour is
   whatever reads on the background. Never assume black.

```python
from office.impress import mm
from office import draw as D

D.rect(doc, page, mm(20), mm(40), mm(60), mm(30), fill=D.color("accent"),
       radius=200)
D.ellipse(doc, page, mm(50), mm(55), mm(15), fill=D.color("muted"),
          transparence=60)
D.text(doc, page, mm(20), mm(40), mm(60), mm(30), "Phase I", "eight weeks",
       role="body", color=D.WHITE, weight=D.BOLD, align="center",
       valign="center")
D.arrow(doc, page, mm(85), mm(50), mm(8), mm(10))
D.chevron(doc, page, mm(20), mm(80), mm(50), mm(14), fill=D.color("accent"))
D.table(doc, page, area, widths=[3, 2, 2], rows=rows,
        head_fill=D.color("accent"))
D.venn(doc, page, area, D.series(3),
       labels=["Private data", "Untrusted content", "External comms"],
       centre=D.color("alert"), centre_text="Exfiltration")
```

Full list: `rect`, `ellipse`, `line`/`hline`/`vline`, `polygon`, `arrow`,
`chevron`, `text`, `chip`, `bullets`, `table`, `cards`/`steps`/`kpis`/
`statement`, `venn` (+ `venn_geometry`, `venn_fit`, `venn_bounds`,
`pair_spots`), `canvas`, `Box`/`page_box`, `restyle`, `walk`/`walk_page`,
`est_lines`/`fit_size`/`fit_lines`/`wrap_words`/`balance_lines`,
`contrast_ratio`/`readable_on`/`bg_color`, `palette`/`color`/`series`,
`type_scale`/`size_for`, plus `move`, `resize` and `set_text` re-exported from
`office.template`.

**`D` draws; `office.template` edits what the template already put there.**
`retext`, `find_shapes`, `shape_by_text`, `strip_annotations`, `shrink_to_fit`
and `fit_pill` are all in `office.template`, and reaching for them on `D` is
the commonest wrong guess in a build script — so `D.find_shapes` raises an
error naming the module and the import to write. Whatever you guess wrong, the
error says what to write instead: that is true of every keyword
(`bold=` → `weight=D.BOLD`) and every `Box` method (`box.split_vertical(2)` →
`box.rows(2)`). **Read the message before changing the call.**

## The primitives, and what goes wrong with each

- **Ellipses take a centre and radii**, not a bounding box. Pass a `Box` and it
  is read as the area to fill, which is usually what you meant.
- **A Venn is `D.venn`, and it takes the area.** Give it the `Box` and it
  solves for the largest figure whose labels still fit, then centres it; give it
  a radius and it draws the radius you guessed — one run's hero slide put its
  Venn in a quarter of the canvas and shipped, so passing `cx, cy, r` now prints
  a nudge. Labels are set **against the edge of their own lobe** — above the top
  one, beside the lower two — rather than pushed out along a diagonal, which is
  how three labels ended up closer to the slide's corners than to the circles
  they name. `centre_text` takes a string (wrapped to fit the disc) or a **list,
  one line per entry**; if the wording cannot be read at the size the disc
  allows, the call says so and tells you to shorten it. `D.venn_fit(box)` gives
  you `(cx, cy, r)` alone if you want to place the parts yourself. `qa.py`
  cannot see any of this; it is a rendering defect and it needs eyes.
- **`D.table` takes the area and proportions**: `widths=[3, 2, 2, 5]` scaled to
  the box, which cannot overflow. Absolute widths in 1/100 mm still work
  (`widths=[mm(60), mm(40)]`) but a list that looks like points or percentages —
  `[80, 60, 40, 120]` — is columns under 1.2 mm wide, and it is refused now with
  a message rather than drawn. That one call shape produced 54 QA findings in a
  single build. With a box and no `row_h`, the rows divide the box height.
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
- **Take the colours and the sizes from the deck, by name.** The template's own
  palette and type scale are in `office.draw`:

  ```python
  D.color("accent")   D.color("alert")   D.series(3)     # colours
  D.text(..., role="footnote")          D.size_for("body")   # sizes
  ```

  Four separate runs against a corporate template wrote their own hex —
  `0xE8730C` for orange, `0x6B3FA0` for purple — and neither colour appears
  anywhere in the deck they were "matching". Four corporate decks, off-brand,
  and nothing could see it. `D.color` raises with the list of real names;
  `D.palette()` is the whole dict and the pack's `catalog.md` prints
  it with each colour's contrast on white.
- **Most brand colours are fill colours, not text colours.** One template's
  brand cyan is 3.0:1 on white and its orange is 2.3:1 — fine behind white
  text, not fine *as* text. Check with `D.contrast_ratio(a, b)` before committing; under 3.0 is
  not readable from a seat, and `qa.py` reports it. `D.readable_on(bg)` picks
  black or white for you.
- **One scale, not a size per call.** `role=` keeps a footnote the same size on
  every slide; anything under `D.MIN_READABLE_PT` (the smallest size the
  template itself types at) is a QA finding. Colours are still plain ints when
  you need one — `probe.py` prints every fill and line colour in a deck.
- `D.table` draws rectangles and text rather than a real table shape: it takes
  per-cell colour without fighting a table style and survives a round trip
  unchanged. Use a real table only if the recipient must edit it as one.
- `est_lines` guesses how many lines a string will wrap to, for laying out
  before the document can measure. `office.template.text_overflow` measures for
  real — trust that one. `fit_size(s, width)` returns the largest point size at
  which `s` fits on one line, and `fit_lines(s, w, h)` returns `(size, lines)`
  wrapped on spaces — use them instead of guessing a size for a label whose
  length you do not control, which is where mid-word breaks come from.

## Lay out with `Box`, not with arithmetic on millimetres

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

**Never pass a box *and* a coordinate.** `D.text(doc, page, area, "…",
y=mm(60))` is the one that gets written: the box already carries x, y, w and h,
so there is nothing for `y=` to mean. To put the content somewhere else in the
area, **slice the box** — `area.bottom(mm(20))`, `area.inset(mm(6))`,
`D.Box(area.x, area.y + mm(60), area.w, mm(20))` — or drop the box and pass four
numbers.

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

## The canvas pattern

**When the catalog has no design for what you need, do not force the content
into the nearest one and do not start from a blank slide.** Take a designed
content slide, strip it back to its title and topic label, and draw on what is
left. You keep the master, background, chrome, fonts and margins — everything
that makes it look like the house deck — and you get an empty area to work in.

```python
def fill(doc, page, slot, design):
    area = D.canvas(page, keep=KEEP)           # the catalog's keep tuple
    retext(page, {TITLE_KEY: "Why it cannot just be patched"})
    left, right = area.cols(2, gap=mm(10))     # draw in what the body left
```

`canvas(page, keep=(...))` removes every shape whose index is not in `keep`;
`canvas(page, remove=(...))` removes only the ones you name. **Run `probe.py` on
the design first** — the indices differ per design and guessing them deletes the
title. The catalog's entry for each design records the keep tuple that works.

**`canvas(page)` with neither is refused**, because it would remove nothing and
hand back an area guessed from the page — so you would draw on top of the
template's own text, and the slide renders in two languages. It returns **one**
`Box`, not a pair: `D.canvas(page)[1]` is not a thing.

**Do not wrap `mm()` around geometry you read back.** `area.y`, `page.Width`,
anything from `probe.py` and every `Box` field are *already* in 1/100 mm;
`mm(26000)` is 26 metres and `mm()` now refuses it. Use `mm()` only on real
millimetres — `area.y + mm(10)`, `D.Box(area.x, area.y, area.w, mm(20))`.

**It returns the free area as a `Box`**, measured from the shapes it removed —
the body placeholder's own rectangle, which is the designer's answer to where
content goes. Use it rather than hardcoding an area: it is exact, and it is
right on a design you have not probed.

**Which design is the canvas, and which indices to keep, is in the pack** —
its catalog's *Designs by role* table names the design, and the design's own
entry gives the keep tuple and the free box it hands back. A typical one keeps
four shapes: the title, a topic label, a source line and whatever decoration
the master puts in a corner, and drops the body placeholder. Do not carry a
keep tuple from one template to another; the indices are not the same twice.

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

### Four of them are one call

`D.cards`, `D.steps`, `D.kpis` and `D.statement` are the layouts every build
draws by hand, drawn once. Reach for these before laying out rectangles: the
hand-rolled versions come out different every time, and the difference is always
in the rhythm — a note floating under one card and inside the next, arrows too
small for their gap, one title in the accent colour and two in black.

**QA watches this.** Every shape `office.draw` makes is named after what made
it. A slide of more than twelve shapes placed one by one (`D.text`, `D.rect`,
...) with nothing from a composition helper goes to the vision-agent first,
and hand-placed peers whose edges or gaps miss by 0.6 to 3 mm fail the gate
(`references/qa.md` §Hand-placed slides). A diagram no helper draws is fine
when its shapes come from one `Box` — `grid`, `cut_*` — because then they line
up exactly. One field run placed 570 shapes by page percentage instead, and
its rows drifted on slide after slide.

```python
D.cards(doc, page, area, [
    ("Inventory", "Classify every agent against A / B / C.",
     "Owner: ‹name› · By: ‹date›"),          # (title, body, note)
    ("Gate", "MCP installation becomes a privileged change."),
    ("Accept the trade-off", "Cap capability where all three coincide.",
     "Board decision · today"),
], numbered=True, accent=[CY, OR, PU], height="content")

D.steps(doc, page, area, steps, captions=[...], accent=[...])   # flow + arrows
D.kpis(doc, page, area, [("~47,000", "downloads"), ("3 hours", "on PyPI")])
D.statement(doc, page, area, "One sentence, as large as it fits.")
```

- **`height="content"` on `D.cards`** sizes every card to the longest one's
  content and centres the row. Use it whenever the cards hold different amounts
  of text; `"equal"` (the default) fills the area and is right only when they
  hold about the same.
- **The type grows to the space.** Without `size=`/`title_size=`, `D.cards`
  and `D.steps` take the largest step on the deck's type scale at which the
  tallest card still fills under 60% of its panel and every word still fits a
  line. Pass sizes only to override that.
- **`note` sits on the card's floor**, so a row of cards has its owner/date line
  on one baseline instead of three.
- **The accent titles a card only when it can be read on it.** Brand colours are
  picked against white and a panel fill is not white — one brand's cyan on its
  own light grey is 2.7:1. If any accent in the row fails, every title in the row
  falls back to black, because peers styled two ways is a defect nothing
  measures. The colour still marks the card: it is the number chip's fill.
- **`D.statement` chooses the size and the line breaks together**, then checks
  them against the layout engine rather than against its own estimate. Do not
  hand-break the sentence with `\n` — it refuses, because hand-breaking is what
  leaves a slide ending on a line reading `an`.

Patterns worth reaching for, roughly in order of how often they earn their keep:

- **Comparison columns** — before/after, pros/cons, option A vs B
- **Timeline or process flow** — numbered steps, chevrons, arrows
- **A table** — when the content really is a matrix, say so with one
- **A chart or a table takes half the slide or all of it** — beside the text
  (about 40/60) or alone on the slide, never squeezed into a band under a
  paragraph
- **Large stat callouts** — 60–72pt numbers with small labels beneath
- **Icon + text rows** — icon in a filled circle, bold header, description below
  (`D.icon(..., backdrop="accent")`, see Icons below)
- **2×2 or 2×3 grid** of content blocks
- **Two-column** — text left, illustration right
- **A diagram** — Venn, quadrant, flow, hierarchy; `office.draw` has the parts

Then:

- **Vary the layout.** Not every section belongs on title-and-bullets. Repeating
  one design down the whole deck reads as a form someone filled in.
- **Size contrast.** Titles well above body text; see Typography in
  `references/authoring.md`.
- **Left-align body text.** Centre titles and short labels only.
- **Hold one spacing rhythm** — pick a gap and keep it. Minimum margins about
  0.5", 0.3–0.5" between blocks.
- **Fill the area you were given.** Content in the top two-thirds with an empty
  band underneath is the signature of hardcoded offsets, and it is invisible to
  most of `qa.py` — nothing overflows, nothing collides, the slide is just badly
  placed. It is no longer invisible to all of it: **`office.qa.layout.balance_findings`
  fails a build on a content slide that paints real ink and then leaves its
  bottom half bare** (see `references/qa.md` §Top-heavy). Lay out from the
  `Box` `D.canvas` returns; if a band is genuinely meant to be empty, it should
  be at least as wide as the content, not a leftover.
- **Contrast on everything**, icons as well as text. A card flipped to a dark
  fill needs `set_char_color(shape, 0xFFFFFF)` on its label.
- **Never ship text that overflows its shape** — shrink it, split it, or make
  the shape bigger. `qa.py` finds these.
- **Never draw an accent line under a title**, and **never add decorative colour
  bars or stripes** — header/footer bars, sidebar stripes, thin rules along a
  card edge. They are the clearest tell of a generated deck. To set a card apart
  use a background tint or an icon.


## Consulting components

The compositions consulting decks reuse, each laid out completely. **Reach
for these before drawing any panel, row or bar by hand.** They choose the
type size for the space, keep peers on one size and one baseline, and
balance the fill. When the content cannot fit at the template's smallest
size, they raise an error that says what to change. Never shrink text past
the template's smallest size yourself.

```python
D.kicker(page, "SDLC · Phase 1 von 6 · Planung & Definieren")  # small line above the title
D.sticker(doc, page, "ÜBUNG · 25 MIN", fill="orange")          # tag top right, clear of the logo

band, area = area.cut_top(mm(13), gap=mm(4))
D.tracker(doc, page, band, PHASES, current=1, done=1)          # phase bar, current one in the accent

D.pairs(doc, page, area, [(risk, risk_text, guard, guard_text), ...],
        heads=("RISIKO DURCH KI", "LEITPLANKE / GATE"),
        tags=["UNSICHERER OUTPUT", ...])                       # risk -> guardrail rows
D.levels(doc, page, area, [dict(label="Ebene 1", title=..., body=...,
                                plus=..., minus=...), ...])    # stages rising left to right
D.matrix2x2(doc, page, area, ["Zuerst", "Planen", "Mitnehmen", "Zurückstellen"],
            "Aufwand", "Wirkung", highlight=0)
rest = D.quote(doc, page, area, "Words of the quote", "Author, source")  # returns the area under it
D.cards(doc, page, area, [
    {"title": "Erkennung", "icon": "ep:filter", "bullets": [...],
     "figure": ("> 90 %", "der Abwehrmaßnahmen umgangen"),
     "figure_color": D.color("alert")},
    {"title": "Architektur", "icon": "ep:files", "fill": D.tint("ok", 0.14),
     "accent": D.color("ok"), "bullets": [...]},
], height="equal")
```

- **`D.emphasize(shape, phrase)`** sets one phrase in bold inside a text:
  the "Grundlage:" before a list of references, a key figure inside a
  sentence. `D.bullets` items can be `(lead, rest)` pairs, with the lead in
  bold, the way consulting decks open a point.
- **Adapt a design before drawing over it.** A design with roughly the right
  structure, such as three columns or four steps, is often closer than a
  canvas. Retext it, `move` and `resize` its shapes, and add icons. Its
  own decoration then stays the template's.
- **`D.tint(name, share)`** is a pale tint of a palette colour: the panel
  fill consulting decks use, one hue family plus its tint. It is derived
  from the palette, not invented.
- **`D.cards`** items take an `icon` (on the title's first line, never over
  the title), `bullets` (square markers in the accent) and a `figure`
  (`(value, caption)`, a big number with its caption).
- **When a component refuses**, follow its message. Typical ones:
  - a taller band for the tracker;
  - the rows split over two slides with "(1/2)" and "(2/2)" in the title;
  - shorter text.

**Size estimates are calibrated to the font the deck really renders in.**
The first drawing call measures how wide the template's font sets on this
machine, regular and bold. Every fitted label and every wrapped line uses
that measurement, not Arial's widths. A template whose font is substituted
(not installed) still gets one-line labels that stay on one line.

## Charts, tables and sources

**Never draw a chart.** The numbers and the chart are the author's: built in
Excel and pasted into the slide. What a build does is everything consulting
decks put around a chart (measured on 48 of them), plus a reserved place of
the right size:

```python
area = D.canvas(page, keep=(0, 1, 2))
area = D.source(doc, page, area, "Quelle: Pilotauswertung, 2026",
                notes=["1) ohne Sondereffekte"])   # returns the area above it
D.chart_with_comment(doc, page, area, "Review-Zeit je Pull Request",
    points=["Plattform-Team halbiert die Wartezeit", "Kein Team wird langsamer"],
    unit="Median in Stunden, vor und nach dem Pilot",
    brief=["Gruppierte Balken: vorher / nachher je Team",
           "Hervorheben: Plattform-Team (Akzentfarbe), Rest grau",
           "Werte an den Balken, keine Gitterlinien"],
    heading="Kernaussagen")
```

- **`D.chart_placeholder`** draws the two-line chart header, with the metric
  in bold and unit, scope and period in grey. The chart itself then needs no
  title and no unit on its axis. Below it is a dashed frame holding the
  `brief`: what the author is to build, written so that the chart argues:
  - the chart type and the data;
  - the one bar or series the slide's title talks about, in the accent
    colour with everything else grey;
  - bars sorted, values written on the bars, no gridlines, no legend box
    where a direct label fits.

  QA lists the slide as *a chart placeholder to fill*, never as a defect.
- **`D.chart_with_comment`** puts the placeholder on the left two thirds and
  a comment column on the right third: an optional bold heading over 3-5
  short points saying what the chart shows. That split is what the decks
  use (chart 65-70 %, comment 25-33 %).
- **`D.source`** puts footnotes and the source line small and grey at the
  bottom left of an area, and returns the area above it. Every slide with
  figures needs one. QA lists a slide with three or more figures and no
  source line. When the design has a source line of its own under the area
  (the canvas usually keeps it), `D.source` writes into that one, so its
  sample text never ships.
- **`D.table`** draws the consulting table by default:
  - a solid header row;
  - a hairline under every body row, no zebra stripes, no vertical rules;
  - every column of numbers right-aligned, its header too.

  `zebra=` brings the stripes back.

## Icons

```python
D.icon(doc, page, box, "checkmark")                    # accent, centred in box
D.icon(doc, page, box, "lock", backdrop="accent")      # white on a filled circle
D.icon(doc, page, box, "warning", color="alert")       # any palette name or int
```

`D.icon` looks a name up in order, first match wins: the pack's own icons
(`templates/<pack>/icons/`, harvested from the template by
`presentation-setup`, and listed in the pack's `catalog.md`), the directories
in `$PRESENTATION_ICONS`, then **Element Plus** (`ep`, 293 icons, MIT,
shipped with the skill in `icons/ep/`). `"ep:lock"` picks a set explicitly
when the pack has an icon of the same name. Never install an icon set: when
nothing fits, pick the closest icon that exists.
`scripts/icons.py search <word>` lists matches, pack icons first; names are
English (`lock`, `coffee-cup`, `user`, `monitor`).
An unknown name raises with the close ones.

- **Every icon is drawn in one palette colour**: its fills and strokes are
  rewritten. With `backdrop=` the icon turns white or accent, whichever reads.
  A pack icon marked multi-coloured is placed in its own colours.
- **Prefer the pack's icons.** They are the brand's style; a general set is
  for what the brand has none of. Do not mix two sets on one slide.
- The box is the icon's space; the icon keeps its proportions inside it. Give
  it a square box and size it from the layout, as for any other shape.
- **Never draw an icon out of rectangles and lines, and never use a Unicode
  symbol or emoji for one.** One run drew 45 that way for a 58-slide deck; on
  the slide they read as "II", "#" and three dots.
- In the `.pptx` the icon is an SVG with a PNG fallback, the way PowerPoint
  stores its own, so it stays sharp and can be recoloured there.
