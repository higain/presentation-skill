# Accents and hero visuals: the look, chosen rather than invented

Loaded at stage 4 of `references/storyline.md`, and any time a deck should look like more
than the standard master. Two closed catalogues: **six accent packs** and **twelve hero
visuals**. You pick from them. You do not add to them mid-deck.

**The one hard rule: no new colours, ever.** Every colour here is one this template
already fills a shape or sets text in, reached by name through `D.color(...)`. Four
separate runs against one corporate template wrote their own hex — `0xE8730C` for
orange, `0x6B3FA0` for purple — and neither appears anywhere in the deck they were "matching".
Four corporate decks, off-brand, and no check in the pipeline could see it. A pack is a
**selection and a discipline**, not a palette.

What that buys: a deck that does not look like every other deck off this template, from
a run that never made an aesthetic decision. Every deviation below was authorised once,
by a person; the run only chooses between them.

---

## 1. The six packs

A pack is five things: an **accent**, a **muted** tone behind neutral content, one
**rule** about where the accent may appear, a **motif**, and the **type emphasis**.
Everything else — fonts, margins, chrome, the master — belongs to the template and is
not yours.

Copy the `PACK` block into the build script and read from it. One pack for the whole
deck; a deck that changes pack halfway is a deck built by two people.

### Signal — the house look, sharpened

```python
PACK = {"accent": D.color("accent"), "muted": D.color("muted"), "radius": 200}
```

- **Rule:** the accent marks **exactly one element per slide** — the thing the title's
  claim is about. Two accented things on one slide means neither is the point.
- **Motif:** rounded blocks (`radius=200`) and numbered chips.
- **Type:** titles in the template's text colour, captions `D.color("neutral")`, numbers
  at `role="display"`.
- **Use for:** anything, and by default. **Check the accent's contrast first** —
  `D.contrast_ratio(D.color("accent"), D.WHITE)`. Most brand accents come in under 4.5:1,
  which means headline sizes only and never body text.

### Ember — the same colour, used only as area

```python
PACK = {"accent": D.color("accent"), "muted": D.color("muted"), "radius": 0}
```

- **Rule:** the accent is **fill only** — it can never be a word. So it always arrives as
  a block with text on it, which is why this pack reads harder than Signal without a
  single new colour. This is the pack for an accent too low-contrast to set type, and
  most of them are.
- **Motif:** square-edged solid blocks. `D.readable_on(PACK["accent"])` picks the text
  colour; do not assume white.
- **Type:** as the template's scale. The emphasis comes from area, not weight.
- **Use for:** cost of inaction, an incident, a deadline.
- If the palette has a second warm colour — `D.series(2)[1]` — this is the pack to spend
  it in. Never a colour from outside the palette.

### Meadow — the same structure, read as progress

```python
PACK = {"accent": D.color("ok"), "muted": D.color("muted"), "radius": 900}
```

- **Rule:** `ok` fills. If the palette has a darker partner for it that clears 4.5:1,
  that one — and only that one — sets the single figure or word per slide allowed to be
  coloured type. Never the reverse.
- **Motif:** discs — `D.ellipse` for step numbers where Signal would use a chip.
- **Type:** as the template's scale.
- **Use for:** a method being taught, a rollout that is going well, a workshop.

### Slate — no colour until it means something

```python
PACK = {"accent": D.color("neutral"), "muted": D.color("muted"), "radius": 0,
        "once": D.color("alert")}
```

- **Rule:** neutrals throughout. `alert` appears **once in the whole deck**, on the single
  number that matters most — not once per slide. That is the entire effect, and it
  collapses the moment it is used twice.
- **Motif:** outline instead of fill — `fill=D.WHITE, line=D.GREY_BD`.
- **Type:** titles `weight=D.BOLD`, everything else quiet.
- **Use for:** boards, audits, anything where restraint is the message.

### Spectrum — colour carries data, and nothing else

```python
PACK = {"accent": None, "muted": D.color("muted"), "radius": 200,
        "figure": D.series}          # D.series(n) for the nth thing on a figure
```

- **Rule:** colour only ever encodes a category on a figure. **If nothing on the slide is
  being distinguished by colour, there is no colour on the slide.** Titles and body stay
  in the template's text colour.
- **Motif:** legend chips (`D.chip`) beside the figure, in `D.series` order.
- **Type:** as the template's scale.
- **Use for:** analysis-heavy decks, several charts, comparisons across categories.

### Contrast — two colours, fixed meanings, whole deck

```python
PACK = {"accent": D.color("accent"), "gap": D.color("alert"),
        "muted": D.color("muted"), "radius": 200}
```

- **Rule:** `accent` is always *how it is today*; `alert` is always *the gap, the
  failure, the thing that is not covered*. Neither is ever used for anything else,
  anywhere. A reader learns the code on slide 3 and reads the rest of the deck faster.
- **Motif:** paired blocks — the two colours next to each other, same size, on every
  slide that has both.
- **Type:** as the house scale. `red` is 4.9:1 and may be type.
- **Use for:** the **Rollout** scaffold, before/after, current control vs. residual risk.

### How the accent actually reaches a card

**`D.cards(..., accent=[…])` alone will not show it. Pass `numbered=True`.**

Measured on this template: `D.cards` drops the *whole row* of card titles to black when
any single accent in the row fails contrast against the panel fill — peers styled two
ways is a defect nothing measures, so it styles none of them. The house cyan is 2.9:1 on
`GREY_LT`, so it always fails. Four of the six packs above have an accent that can never
be a card title (Signal and Contrast's cyan, Meadow's lime, Ember's orange), and in
Contrast's case the cyan's failure also suppresses the red that would have passed.

With `numbered=True` the accent becomes the **number chip's fill**, which is a filled
block and has no contrast problem. That is the carrier for every pack. Only Slate's greys
and Contrast's red can also be title colour — and mixing the two carriers in one deck is
what the row-wide fallback exists to prevent.

The same applies to `D.steps`, where `captions=` are colour-matched to `accent=` and are
the cheapest way to let a flow say which phase belongs to which colour without a legend.

### Which pack, if the user has no opinion

| Scaffold (`spine.md`) | Default pack |
|---|---|
| Case | Signal |
| Rollout | Contrast |
| Report-out | Spectrum |
| Keynote | Ember |
| Workshop | Meadow |
| Board | Slate |
| Close | Signal |

This is a starting point for the preview, not a decision. The user makes the decision.

---

## 2. Show three, ask which

**Do not ask a colleague what colours they like.** Non-designers cannot specify a look
and reliably recognise one. Render their own strongest title in three packs and let
them point.

1. Take the highest-`proof` slide from `plan.md` — the one with the real number.
2. Build it three times, one pack per slide, on the canvas design.
3. Render, look at the three images, and show them.
4. Ask: **"which of these three feels right for this room?"**
5. Write the answer to `look.md`: the pack name, and nothing else.

```python
# preview.py — three packs, one real slide, three renders
from office.deckbuild import build_deck
from office.template import retext
from office.impress import mm
from office import draw as D

TITLE = "Auf 42 geprüfte Bibliotheken kommen 1.447 installierte."
PACKS = [("Signal",   D.color("accent"),  200),
         ("Contrast", D.color("alert"),   200),
         ("Slate",    D.color("neutral"),   0)]

def fill(doc, page, slot, design):
    name, accent, radius = PACKS[slot]
    area = D.canvas(page, keep=KEEP)                  # from the catalog
    retext(page, {TITLE_KEY: TITLE, LABEL_KEY: name})
    top, bottom = area.cut_top(mm(48), gap=mm(8))
    D.kpis(doc, page, top, [("42", "geprüft"), ("1.447", "installiert")],
           color=accent)
    D.cards(doc, page, bottom,
            [("Geprüft", "Was im Review tatsächlich gelesen wird."),
             ("Installiert", "Was zur Laufzeit im Prozess landet.")],
            fill=D.color("muted"), accent=[accent, accent], radius=radius,
            height="content")

build_deck("templates/acme/template.pptx", "preview.pptx", [8, 8, 8],
           fill=fill, footer=None)
```

```bash
scripts/make_deck.py preview.py preview.pptx --template templates/acme/template.pptx
```

**If you cannot see images, say so and skip this stage** — do not fake it.
`scripts/wireframe.py` prints ink, not colour, so it cannot tell three packs apart and a
verdict drawn from it would be invented. Show the user the pack table in §1 instead, name
the default for their scaffold, and ask them to confirm it in words. A stage that was
skipped and said so is fine; a preview nothing looked at is how an off-brand deck ships.

---

## 3. Hero visuals

**The `visual:` field in `plan.md` names one of these.** Twelve names, a closed
vocabulary. "A diagram showing the relationship between…" is not a value this field
accepts — it is how a run talks itself into a slide it cannot build.

Four of them are already one call in `office.draw`; the rest are a fixed slice of the
`Box` that `D.canvas` returns. **Never lay one out with arithmetic on millimetres** — a
run of `mm(30 + i * 70)` assumes a slide width and leaves a dead band under the content
that no check reports. See `references/drawing.md`.

| `visual:` | When the slide is about | How | Design |
|---|---|---|---|
| `statement` | one sentence that has to land | `D.statement` | canvas |
| `single-number` | one figure and what it means | `D.kpis` with one item | canvas |
| `stat-strip` | two to four numbers together | `D.kpis` | canvas |
| `cards` | 2–4 parallel things of equal weight | `D.cards(..., height="content")` | canvas, or a two-column design with image slots |
| `steps` | a sequence with direction | `D.steps` | a phase or process design if the catalog has one, else canvas |
| `timeline` | a sequence with dates | `D.steps(..., captions=dates)` | a schedule design if the catalog has one, else canvas |
| `table` | the content really is a matrix | `D.table(area, widths=[3, 2, 2], rows=rows)` | canvas |
| `venn` | overlapping sets, and the overlap is the point | `D.venn(doc, page, area, D.series(3), labels=[…])` | canvas |
| `area-compare` | two quantities whose ratio is the argument | §3.1 | canvas |
| `before-after` | one thing, two states | §3.2 | canvas |
| `matrix-2x2` | two axes, and one quadrant is where you are | §3.3 | canvas |
| `layered-stack` | tiers, defences, or a plan with levels | §3.4 | canvas |

Every canvas row is the same opening line, with the keep tuple from the catalog:

```python
area = D.canvas(page, keep=KEEP)             # the box the body vacates
```

### 3.1 `area-compare`

Two rectangles whose **areas** carry the ratio. The commonest way this goes wrong is
scaling only the height, which understates a 34× difference by a factor of 34.

```python
small_v, big_v = 42, 1447
k = (small_v / big_v) ** 0.5                      # both dimensions, so area is right
left, right = area.inset(mm(10)).cols(2, gap=mm(16))
D.rect(doc, page, D.Box(left.x2 - left.w * k, left.y2 - left.h * k,
                        left.w * k, left.h * k), fill=PACK["muted"])
D.rect(doc, page, right, fill=PACK["accent"], radius=PACK["radius"])
D.text(doc, page, left.bottom(mm(8)), "42 direkt geprüft", role="caption")
D.text(doc, page, right.bottom(mm(8)), "1.447 installiert", role="caption",
       color=D.readable_on(PACK["accent"]))
```

### 3.2 `before-after`

Two columns, an arrow in the gap between them — the gap the `cols` call already left.

```python
left, right = area.inset(mm(8)).cols(2, gap=mm(26))
D.cards(doc, page, left,  [("Heute", "Jede Abhängigkeit landet ungeprüft im Build.")],
        fill=D.GREY_LT, height="content")
D.cards(doc, page, right, [("Nach dem Gate", "Neue Abhängigkeiten sind ein Change.")],
        fill=D.GREY_LT, accent=[PACK["accent"]], height="content")
D.arrow(doc, page, D.Box(left.x2, area.cy - mm(7), mm(26), mm(14)),
        color=PACK["accent"])
```

### 3.3 `matrix-2x2`

Four cells, **one** of them accented. If two are accented the slide has no point.

```python
plot = area.inset(left=mm(26), bottom=mm(14), top=mm(4), right=mm(4))
cells = [c for row in plot.grid(2, 2, gap=mm(4)) for c in row]   # grid gives ROWS
for cell, (label, body), here in zip(cells, quadrants, [0, 0, 1, 0]):
    fill = PACK["accent"] if here else D.GREY_LT
    D.rect(doc, page, cell, fill=fill, radius=PACK["radius"])
    D.text(doc, page, cell.inset(mm(6)), label, body,
           role="body", color=D.readable_on(fill))
D.text(doc, page, area.left(mm(24)), "Wirkung", role="caption", align="center",
       valign="center")
D.text(doc, page, area.bottom(mm(12)), "Aufwand", role="caption", align="center")
```

**Both axis labels go in the strips the `inset` reserved — `area.…`, never
`plot.…`.** `plot.bottom(mm(12))` is the bottom 12 mm *of the plot*, so the label lands
on top of the two lower cells; it renders clean through every check, because text over a
filled rectangle is not an overflow and not a collision.

### 3.4 `layered-stack`

Horizontal bands, top to bottom, one highlighted — tiers of a plan, layers of a defence,
levels of a maturity model.

```python
bands = area.inset(mm(10)).rows(len(tiers), gap=mm(3))
for band, (name, note), here in zip(bands, tiers, marks):
    fill = PACK["accent"] if here else D.GREY_LT
    D.rect(doc, page, band, fill=fill, radius=PACK["radius"])
    D.text(doc, page, band.inset(x=mm(8)), name, note, role="body",
           color=D.readable_on(fill), valign="center")
```

### When the number is still `[TBD]`

Three of the twelve are drawn **from** a figure, so a `[TBD]` in `plan.md`'s `proof`
field is not a gap in the text — it is a missing dimension. Decide which case you are in
before the build, not during it:

| `visual:` | With the number unknown |
|---|---|
| `single-number` | **Draw the unknown.** `D.kpis(doc, page, area, [("[TBD]", "direct vs. transitive, from the SBOM")])` sets it at display size like any other figure. A slide reading *"We don't know"* in 54pt is a finding stated at full volume, and it is often the strongest slide in the deck. Do not shrink it and do not apologise for it in a caption. |
| `stat-strip` | Same, per cell. A strip of three where one is `[TBD]` reads correctly: two answers and an open question. |
| `area-compare` | **Cannot be drawn.** Its whole argument is the ratio of the two areas, and there is no ratio. Equalising the blocks does not defer the claim — it *makes the opposite one*. Either get the figure before the build, or swap the slide to `single-number` on the `[TBD]` and let that carry the point. |

The other nine are qualitative and build unchanged; a `[TBD]` lands in their text.
`table` is the useful case — **leave the cells blank rather than dropping the rows.** A
table of question marks argues for the inventory better than a sentence saying the data
is missing.

### Rules that hold across all twelve

- **One hero per slide.** A hero visual plus a bullet list is a slide that could not
  decide.
- **The figure fills the area `D.canvas` gave you.** Content in the top two-thirds with
  an empty band underneath is the signature of hardcoded offsets and no check reports it.
- **`D.readable_on(fill)` picks the text colour.** Never assume black — a headline drawn
  in black onto a mid-blue block renders clean through every check and is unreadable.
- **Caption every number** with `role="caption"`, and put the source on the slide.
- **No decorative rules, bars or stripes** — not under the title, not along a card edge.
  They are the clearest tell of a generated deck. To set something apart, tint it or
  give it the accent.

---

## 4. What stays off-limits

Not because none of it can work, but because when it fails it fails invisibly and lands
in front of a customer:

- **Inventing a colour.** Any hex not in `D.palette()`. This is the one that has actually
  happened, four times.
- **Restyling the master** — background, chrome, logo placement, the decorative triangle.
- **Generating imagery.** The template's photo placeholders take the customer's own
  pictures or nothing.
- **A new type scale.** `role=` and `D.size_for` only; anything under
  `D.MIN_READABLE_PT` is a QA finding.
- **A second pack in the same deck**, or a pack chosen per slide.
- **A pack applied to a template with no profile.** Every colour above is a *role*, and
  the roles come from the pack. On a template the user brought that has no pack, read
  its colours with `probe.py` and name them explicitly — see `references/templates.md`.

---

## Attribution

The preset-catalogue pattern and the three-preview picker follow `high-quality-slides`
(MIT); the deck-type → look mapping follows `presentation-director` (MIT). The packs
themselves — expressed in the palette roles every profile defines — and the hero-visual
catalogue are ours. See `THIRD-PARTY.md`.
