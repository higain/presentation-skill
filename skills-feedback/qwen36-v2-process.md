# qwen36-v2 Session Retrospective

**Date:** 2026-07-23
**Task:** Read `Aurora_PM_Planning.xlsx`, analyze the project plan, then build a branded presentation from `Kite_brand_master_slide_deck.pptx` using the presentation skill.

---

## Overall Assessment

The session completed successfully but required **3 full rebuilds** of the presentation script due to indexing errors and leftover template text. The spreadsheet analysis was straightforward; the presentation build was the harder part.

---

## Spreadsheet Skill: Smooth

The spreadsheet skill worked well:

- `dump.py` gave a clean markdown dump with formulas and coordinates — exactly what was needed.
- `recalc.py` ran without issues (51 formulas, 0 errors).
- `openpyxl` with `data_only=True` provided computed values cleanly.
- **One gap:** `markitdown.sh` is documented in the skill but the script doesn't exist on disk (`FileNotFoundError`). Not critical — `openpyxl` filled the gap — but the docs should be updated.

**Rating:** Easy. No errors.

---

## Presentation Skill: Workable but Frictional

### What Worked Well

- `dump.py`, `probe.py`, and `thumbnail.py` gave complete visibility into the template structure.
- `clone_slide` reliably copies slides with all formatting intact.
- `set_text` preserves character formatting (font, weight, color, kerning) — critical for branded decks.
- `render.py` and `validate.py` provided solid QA. Validation passed cleanly against the original template.

### Errors Encountered

#### 1. Slide index mismatch (IndexError, 2 iterations)

**Problem:** I cloned 13 slides at the end of the document, then deleted 9 originals. I assumed the clones would occupy `A[0]` through `A[12]` in the order I called `clone_slide`. This was correct. But my first script used **5 dividers grouped together** at indices 2–6, pushing the kicker boxes to indices 7–8. When I tried to access `s[6]` on what I thought was a kicker box (22 shapes), it was actually a divider (6 shapes), causing `IndexError`.

**Fix:** Rebuilt the clone order so dividers and content slides interleave correctly: cover, agenda, div, kicker, div, kicker, div, kicker, div, content, data-viz, div, closing.

#### 2. `clone_slide` with `index` parameter — unexpected behavior

**Problem:** I tried using `clone_slide(doc, tpl[N], index=K)` to insert clones at specific positions. This shifted the original template slides around, making it impossible to predict which slide ended up where. The `index` parameter inserts *before* the slide at that index, but since the originals are still present, the final order after deletion is not what you'd expect.

**Fix:** Clone all slides at the end (no `index`), then delete originals high-to-low. The clone call order becomes the final slide order.

#### 3. Leftover template text (visual QA catch)

**Problem:** Three issues only visible in renders:
- Kicker box slides showed template title "Tagged cards for quick points" — the `s[0]` and `s[1]` shapes weren't updated.
- Content-rich slide had a leftover template paragraph ("Two shorter paragraphs read better than one long block on a projected slide") in the middle column at `s[7]`.
- The data-viz slide's x-axis labels ("BACKEND", "FRONTEND", "DATA", "QA") were too long for the bar widths and wrapped awkwardly.

**Fix:** Added `set_text` calls for the kicker titles and cleared `s[7]` on the content-rich slide. The x-axis labels are a template limitation — the bars are sized for short labels like "Q1"/"Q2".

### What Could Be Improved

#### In the skill scripts

1. **`markitdown.sh` is missing.** The skill docs reference it but it doesn't exist. Either ship it or remove the documentation.

2. **`clone_slide` `index` parameter is misleading.** It inserts at a position, but when combined with deleting originals, the final order is non-obvious. Consider documenting a clear pattern: "Always clone at end, delete originals, then reorder if needed" — or provide a `reorder_slides` helper.

3. **`probe.py` output could be machine-readable.** A `--json` flag would let scripts programmatically identify shapes by text content rather than hardcoding indices. Currently you have to manually map `probe.py` output to shape indices.

4. **Template text detection.** There's no way to programmatically find shapes that still contain template placeholder text. A helper like `find_shapes_with_text(page, pattern)` would catch leftover text before rendering.

5. **`set_text` on empty shapes.** When you call `set_text(s[7], "")` to clear a shape, it leaves an empty text box. A `remove_paragraph` or `clear_shape` helper would be cleaner.

#### In the prompts / workflow

1. **Always render before declaring done.** The skill docs emphasize visual QA, and it caught 3 issues that `dump.py` wouldn't show (template titles, leftover paragraphs, label wrapping). This should be a mandatory step, not optional.

2. **Probe the template once, save the mapping.** A good practice would be to run `probe.py` once and document the shape index → purpose mapping in a comment block at the top of the build script. This avoids guesswork on rebuilds.

3. **Build incrementally.** Clone and edit one slide type at a time, render, verify, then move on. Building all 13 slides in one pass means a single index error breaks everything.

---

## Bottom Line

The skills are capable and the scripts are well-designed for the common case. The main friction points are:

| Issue | Severity | Fix |
|---|---|---|
| `clone_slide` index behavior | Medium | Document the "clone at end, delete originals" pattern |
| Leftover template text | Medium | Render QA catches it; a detection helper would help |
| `markitdown.sh` missing | Low | Remove from docs or ship the script |
| No machine-readable probe output | Low | Add `--json` flag to `probe.py` |

With these improvements, a presentation build could likely be done in a single pass rather than 3 iterations.
