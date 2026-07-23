# Session process notes — z-ai/glm-5.2

**Task:** Read `Aurora_PM_Planning.xlsx` with the spreadsheet skill, document the
analysis in markdown, then build a branded presentation from
`Kite_brand_master_slide_deck.pptx` using the presentation skill.

---

## What went well

- **`dump.py` (spreadsheet) worked first try.** It is the rare reader that gives
  cell values *and* formulas *with coordinates* on both `.ods` and `.xlsx`, in
  one pass. I got everything needed to write the analysis in a single command —
  no `data_only=True` / formula-pass dance like openpyxl demands. The analysis
  markdown was written in one shot with zero rework.
- **`probe.py` (presentation) is the reason the deck was buildable at all.** It
  prints every shape's index, position, size, fill, line and character
  formatting. Retexting a template slide needs the *index* of each shape, not
  its text, and probe gives exactly that. Without it, editing a designed slide
  blind would be impossible.
- **The `clone_slide` → retext → delete originals LAST workflow** described in
  the skill is correct and important; following it verbatim produced a clean
  16-slide deck from 9 template slides with no ordering problems.
- **`validate.py --original`** baselined against the template and passed on the
  first valid build — good safety net for `.pptx` file integrity.
- **The "keep the format the deck arrived in" rule** is stated forcefully and
  prevented me from silently handing back an `.odp` for a `.pptx` job.

## Errors encountered

### 1. Write tool blocked on `/tmp/opencode/` despite bash allow rule
I first tried to write the build script to `/tmp/opencode/build_deck.py` (the
pre-approved temp dir per the bash tool's own description). The **Write** tool
rejected it with a permission rule error, even though the permission list
explicitly allows `/tmp/opencode/*` for bash. I fell back to writing the script
into the workspace (`build_deck.py`) and deleted it afterwards. The permission
model for the Write tool and for bash external-directory access appear to be
out of sync.

### 2. Miscounted shape indices on the 3-column content-rich layout (caught by QA)
This was the real bug. The template's slide 5 (content-rich, three columns) has
this shape layout according to `probe.py`:

```
[3] "Background"      (col1 header)
[4] long body...      (col1 body)
[5] "What we found"   (col2 header)
[6] paragraph 1...    (col2 body, first paragraph)
[7] paragraph 2...    (col2 body, SECOND paragraph — a second text box)
[8] "What it means"   (col3 header)
[9] long body...      (col3 body)
```

I assumed a clean 3×2 grid: `[3,4] / [5,6] / [7,8]`, and mapped col3's header
to shape `[7]`. In reality `[7]` is an *extra second-paragraph box in column 2*,
so my "col3 header" text landed in the wrong place and the original template
text ("Close with the implication and the recommended next step…") leaked
through as a leftover paragraph on slides 7, 13 and 15.

**How I caught it:** the post-build `dump.py` content grep for leftover
template phrases (`close with`, `what it means`, `two shorter`, etc.) matched
immediately. I cleared shape `[7]` and shifted col3 to `[8]`/`[9]`. Rebuilt,
re-dumped — clean. The bug never reached the user. This is exactly what the
skill's "Write 2–3 formulas first, dump them, check" advice is about, applied
to slides: verify a small slice before building the grid.

**Root cause:** column 2 happened to carry two body paragraphs in the template,
where columns 1 and 3 carried one. `probe.py` lists shapes in z-order, not in
reading order or column-grouped order, so the extra box was easy to overlook
in a long text dump.

## What was hard

- **No image input.** This model cannot view images. The skill's recommended
  visual-QA step ("Pass those paths to the view tool and look at every slide")
  is impossible for me to execute. I ran `thumbnail.py` and `render.py` as
  instructed, but they produced no usable information for me — the thumbnail
  step was effectively wasted work. I relied entirely on `probe.py` coordinates
  and `dump.py` text for every layout decision, and on `validate.py` for file
  integrity. Whether the resized data-viz bars look proportionally correct,
  whether any text overflows its box, and whether the dark-card text is legible
  — I genuinely cannot confirm. I flagged this to the user explicitly and asked
  them to eyeball `slide-04.png`, `slide-10.png`, `slide-06.png` and
  `slide-12.png`.
- **Editorial scope was undefined.** The user asked for "a presentation of
  this analysis and the project plan" with no guidance on length, audience or
  depth. I chose 16 slides (cover, agenda, 5 section dividers, 5 content
  slides, closing) — a middle ground. A one-line brief ("10-slide exec
  summary" vs "detailed 20-slide deep dive") would have saved the guesswork.

## What could be improved

### In the presentation skill

- **Add a column-aware or reading-order view to `probe.py`.** When a layout has
  repeated column blocks, grouping shapes by horizontal band (or emitting a
  one-line structural summary like "3 columns × {1 body, 2 body, 1 body}")
  would make the miscount I hit much less likely. The raw z-order list is
  faithful but not gestalt-friendly, especially for a model that cannot see the
  rendered slide.
- **`dump.py` on a template could label placeholder/structural roles.** Right
  now it dumps text in z-order with no grouping. A light structural annotation
  ("column 2, paragraph 2 of 2") would surface the extra-paragraph case that
  bit me.
- **Consider a non-visual QA signal for text overflow.** The skill correctly
  insists on visual QA, but for vision-less agents there is no fallback. Even a
  heuristic — flag a text box whose string length exceeds some function of its
  width and font size — would give a blind agent something to act on. This is
  a real gap, not a nice-to-have: the skill is currently unusable for
  high-confidence layout work without image input.
- **Document the "extra paragraph box" pattern explicitly.** Designed templates
  often split a column's body across two text boxes for leading-indent or
  spacing reasons. A one-line warning in the "Adapting a template deck"
  section ("column bodies may be split across multiple shapes; clear every
  shape in the column, not just the first") would have prevented my bug.

### In the task prompts

- **A one-line scope brief** (slide count, audience, depth) would remove the
  main guesswork in deck-building tasks. "Build a 12-slide executive review"
  vs "build a detailed walkthrough" produce very different decks from the same
  source material.
- **The "copy and edit the slides from the master deck" phrasing was good** —
  it maps cleanly to the `clone_slide` workflow and left no ambiguity about
  whether to start from the template or from `new_document()`. No change
  needed there.

### In the environment

- **Reconcile the Write-tool and bash external-directory permissions for
  `/tmp/opencode/`.** Bash is allowed to use it; the Write tool is not. Either
  allow both or document that the Write tool is restricted to the workspace,
  so the agent doesn't waste a round-trip discovering the mismatch.

## Net assessment

The spreadsheet skill was friction-free — `dump.py` is a genuinely good tool and
the analysis was a single-shot job. The presentation skill is well-engineered
for a sighted agent: `probe.py` + `clone_slide` + `set_text` + `validate.py` is
a coherent, correct pipeline. For a vision-less agent like this one, the
pipeline is *buildable but not fully QA-able* — I can produce a valid,
content-correct deck but cannot personally vouch for its visual quality, and
had to lean on the user for the final visual sign-off. The one real bug
(shape-index miscount) was caught by the skill's own content-QA step, which is
a point in the skill's favour.
