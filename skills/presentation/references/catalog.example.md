# The catalog a pack carries (EXAMPLE / annotated)

**A specimen of what `catalog_build.py` writes**, from the public demo template
in `tests/fixtures/`, annotated with where each part comes from. Nobody writes
a catalog by hand: run the pipeline in `README.md` and review the result.

A pack's catalog is two kinds of file:

| File | Read | Holds |
|---|---|---|
| `catalog.md` | first, in full, for every deck | how to use it, the roles, the decision table, the design index, the palette |
| `designs/design-N.md` | for each design a deck uses | its retext keys, empty slots, budgets, use / avoid |

The split exists because a single file with every key of a sixty-design
template runs to sixteen thousand tokens, and a small model reading it loses
the decision table it opened the file for.

Every field is one of two kinds:

| Marked | Comes from | Trust |
|---|---|---|
| **[measured]** | `catalog_extract.py`, read out of the template file | exact |
| **[judged]** | the vision-agent (`vision_brief.py designs`), reviewed by a person | review it |

---

## `catalog.md`

**Design index** — one row per design. *Name* and *Content shape* are
**[judged]**; *Headline* is **[measured]**: `set_title` when the design's title
placeholder is empty, otherwise the retext key of the title placeholder or of
the largest text on the slide.

| Design | Name | Content shape | Headline |
|---|---|---|---|
| `1` | Title slide | Slide 1 of any deck, with a title of roughly 20-25 characters so it stays on one line at 99pt and a subtitle that reads as a single sentence. | retext `"Kite"` (the largest text) |
| `2` | Brand reference — **canvas** | Onboarding, brand or handover decks where the audience has to leave with the palette, the type pairing and the logo rules. | retext `"One clear, friendly style"` (the largest text) |
| `3` | Agenda list | Straight after the cover to show the running order, or for any short ordered sequence - three to five entries, each a phrase of up to about 40 characters. | retext `"Agenda"` (the largest text) |
| `4` | Section divider | Between chapters of a deck, with numbered sections and a title of up to about 30 characters so it holds one line at 72pt. | retext `"Getting started"` (the largest text) |
| `5` | Three-column narrative | A topic that needs argument rather than labels - background / finding / implication, or any three-beat structure - at roughly 150-250 characters per column, split into one or two short paragraphs. | retext `"Room for the full story"` (the largest text) |

**Decision table** — **[judged]** rows mapping a content situation to a
design, then a **[measured]** canvas row last: the design with the largest
free area, its `keep` tuple and the area `D.canvas` hands back. When the canvas
keeps a shape that still carries sample text, the catalog says so under the
table — a build must retext it on every canvas slide.

**Palette and type scale** — **[measured]** by `count_style.py`, named by a
person in `presentation-setup` stage 2.

## `designs/design-7.md`

```markdown
# Design 7 — Tagged card grid
- **Purpose:** A 2x2 grid of rounded cards, each with a coloured pill kicker, a short heading and two lines of body; the cards come in white, ink and blush so one can be flipped dark for emphasis.
- **Content types:** `cards`, `comparison`
- **Capacity:** headline ✗ · body slots 0 · images 0 · chart ✗
- **Headline:** retext `"Tagged cards for quick points"` (the largest text)
- **Retext keys:** `[0]` "KICKER BOXES" · `[1]` "Tagged cards for quick points" · `[4]` "SPEED" · `[5]` "Fast to assemble" · `[6]` "Duplicate a card, swap the kicker, done. The" · `[9]` "CLARITY" · `[10]` "Easy to scan" · `[11]` "The colored tag tells the audience what kind" · `[14]` "CONTRAST" · `[15]` "Emphasis on demand" · `[16]` "Flip a card to ink to make the single most i" · `[19]` "RANGE" · `[20]` "Two, three, or four" · `[21]` "The same card works in any count. Use four f"
- **Use when:** Two to four parallel points that read better with a label on them - themes, risks, options, principles - or a two-card comparison side by side.
- **Avoid when:** More than four points, or body copy longer than two lines - the card is a fixed box and text overflows rather than shrinking. Flip only one card to ink, or the emphasis stops meaning anything.
- **Imagery:** none

Any distinctive substring of a key works; copy it from here, not from memory.
```

- *Retext keys* and *Empty slots* are **[measured]**: the exact strings the
  design carries, which is what `retext` matches.
- *Purpose*, *Content types*, *Use when*, *Avoid when*, *Content shape* and
  *Looks best* are **[judged]**.
- Notes nobody can measure — "this divider's budget is a hard limit" — go in
  the pack's `overrides.md` under `## design-7` and are spliced in on rebuild.
