# Company template — layout catalog (EXAMPLE / annotated)

**This file is a specimen of what `catalog_build.py` writes, annotated with
where each field comes from.** You do not fill a catalog in by hand any more —
you run the pipeline in `README.md` and review the result against the renders.
Keep this file as the answer to "is my generated catalog missing something?".

Two sources feed every entry:

| Marked | Comes from | Trust |
|---|---|---|
| **[measured]** | `catalog_extract.py`, read out of the template file | exact |
| **[judged]** | the vision-agent, from each slide's render | review it |

The measured fields are the ones a build script cannot work without, and exactly
the ones a model gets wrong when asked to invent them: the retext keys, the
empty slots, the group counts, the canvas indices, the title budgets.

---

## Header  **[measured]**

Asset path, design count, slide size in 1/100 mm, master names, and the reminder
that a design is identified by its **slide number** — never by the autolayout
number `probe.py` prints as `layout`, which is the same on every slide of a
designed template and so distinguishes nothing.

Plus a line naming the designs that carry text **inside groups**, because
`shapes(page)[i]` cannot reach any of it and a build that indexes by hand will
leave the template's own wording on the slide.

## Designs by role  **[measured]**

A three-row table naming the **canvas** design, the **divider** designs and the
**cover** design (or "none — build one"), so that the generic documentation can
say "the canvas design" and this pack says which number that is.

## Decision table — content → template slide

Applied top to bottom, first match wins.

The rows are **[judged]**: the model's aggregation pass, or the worksheet a
person filled. A row citing a slide the template does not have is written out
with a ⚠ rather than silently dropped, so a model that fell back to layout ids
is visible on the page.

The **last row is always the canvas row** **[measured]**: the design with the
largest free area that still keeps its own chrome, the `keep=` tuple that strips
it back, and the exact free `Box`. Everything the rows above have no design for
goes here rather than into the nearest bullet list. It is last on purpose: first,
"a diagram, a comparison, a matrix" was the first match for almost anything. A
`canvas` block in `decision_table.json` (written by hand, or from the worksheet's
`## Canvas` section) overrides the measured pick and survives every rebuild.

| When the content is… | Use slide | Design | Because |
|---|---|---|---|
| deck opener | `7` | Section divider dark | high-impact start |
| standard bulleted content | `8` | Content slide | headline plus one body |
| … | | | |
| **a diagram, chart, matrix — anything no row above fits** | `8` | **Canvas** | Strip design 8 with `draw.canvas(page, keep=(0, 1, 2, 3))`; free area **2031, 3545 to 30003, 15702** |

**Variety rule** **[judged]** — one sentence on which designs to interleave so
the deck does not read as one design repeated.

**No design for: …** **[measured, from the table]** — content situations the
decision table never covers. "There is no cover design and no closing design" is
exactly the sort of thing a deck-builder has to be told once, loudly.

**Display titles hold one line** **[measured]** — every title slot set at 32pt or
more: box width, size, and how many characters fit on one line. A display title
that wraps runs into the master's own decoration; nothing on the page collides,
and no automated check can see it.

## Palette and type scale  **[measured]**

Read from the pack's `profile.json`: every colour by name and role with its
contrast on white, and the seven type roles with their sizes. This is what
`D.color("accent")` and `role="body"` resolve to.

## Design entries

One per design, in slide order.

### Slide 5 — Section divider   *(name: [judged])*
- **Purpose:** _[judged]_ what this slide is for, one sentence.
- **Content types:** _[judged]_ from the fixed vocabulary in the prompt.
- **Capacity:** _[judged]_ headline ✓ · body slots 0 · max bullets 0 · images 1 ·
  chart ✗ — **the field the model most often overstates; check it against the
  render.**
- **Retext keys:** _[measured]_ `[0]` "Projektorganisation" · `[3.1]` "Phase 1" —
  every string the design carries, with the address that carries it. Copy a key
  from here rather than guessing one; any distinctive substring works. `[3.1]`
  is a shape inside the group at `[3]`.
- **Empty slots** _[measured]_ — shapes with no text, which `retext` can never
  match. Write them with `set_text(shapes(page)[i], ...)` or draw your own.
- **Text in groups** _[measured]_ — how many text shapes sit inside groups versus
  at the top level.
- **Title budget** _[measured]_ — characters on one line, for display slots only.
- **Use when / Avoid when:** _[judged]_ the trigger to reach for this design, and
  the content that does not belong on it.
- **Imagery:** _[judged]_ what each picture actually depicts and its role
  (decorative / photo-placeholder / icon / logo / background), cross-checked
  against the number of picture shapes in the file _[measured]_.
- **Notes:** _[judged]_ quirks a builder must respect — fixed decoration that
  must not be deleted, a slot that holds one line.

---

**Provenance** — a footer naming which half came from where, and the reminder
that the asset, the renders and the catalog are brand-internal and gitignored.

If a design's judged fields read _(not catalogued)_, the model pass did not
cover it. Run the per-layout pass for those slides and rebuild.
