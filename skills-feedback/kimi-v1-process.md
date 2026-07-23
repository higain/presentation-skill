# Session Retrospective — kimi-v1

**Date:** 2026-07-23
**Agent:** Kimi (moonshotai/kimi-k3), running in OpenCode
**Working dir:** `/home/maxi/workspace/anthro-skills/testdir-kimi3`

## What was accomplished

1. **Spreadsheet analysis** — Read `Aurora_PM_Planning.xlsx` (5 sheets, formulas included) with the
   `maxi-spreadsheet` skill and wrote a full programme analysis to `Aurora_PM_Planning_Analysis.md`
   (portfolio roll-up, critical dependency chain, team workload, budget vs. estimate, risks).
2. **Presentation build** — With the `maxi-presentation` skill, adapted the
   `Kite_brand_master_slide_deck.pptx` template into a 9-slide deck
   (`Aurora_Programme_Review_kimi-v1.pptx`) via `clone_slide` / `set_text`, with rescaled bar
   charts, re-fitted kicker pills, an ink-flipped emphasis card, and speaker notes on every slide.
3. **Full QA on both deliverables** — formula-error check on the workbook, rendered-PNG visual
   inspection of all 9 slides, `validate.py --original` schema check (passed), placeholder-text
   grep (clean).
4. Renamed the final deck to contain `kimi-v1` as requested.

## Errors encountered

Only one real error in the whole session:

- **Write to `/tmp/opencode` was denied.** The bash-tool guidance says to use `/tmp/opencode` for
  temporary work, and the permission rules list `/tmp/opencode/*` as allowed — but the `write`
  tool call was still blocked by the user's permission rules. Fix was trivial (wrote the build
  script into the working directory instead), but it cost one failed tool call.

Everything else worked on the first attempt: the spreadsheet dump, the probe-driven template
adaptation, the one-pass deck build, and all QA scripts.

## How easy were the skill scripts to work with?

**Largely easy — the skills are unusually well documented.** Specifics:

- `dump.py` (spreadsheet) is excellent: cell coordinates, cached values **and** formulas in one
  pass. Planning the analysis required no second tool. The cross-check with a small pandas script
  (team workload) was a nice confirmation, not a necessity.
- `thumbnail.py` → `dump.py` → `probe.py` is a very effective template-survey workflow. The
  contact sheet made layout selection a 10-second decision, and `probe.py`'s shape indices,
  geometry and colors were exactly what the build script needed. The skill's claim that
  "`probe.py` is the step that saves the afternoon" held true.
- The "Adapting a template deck" recipe (`clone_slide`, `set_text`, `fit_width`, `move`,
  `resize`, `set_char_color`, delete templates last, single save) worked verbatim. No debugging,
  no second pass.
- The warnings in the skill docs earned their keep: "structural work before content edits",
  "one .pptx save, never a save loop", "never bare `setString` on a template shape", and
  "`Position`/`Size` are structs — use `move`/`resize`". Following them meant zero broken
  states. The note that a template chart "is just moving rectangles" directly enabled the two
  rescaled bar charts.
- The locale warning (German `de_DE`) and font-substitution table were good context even though
  this deck was English/Arial.

Net: the hard part of the session was the *content design* (mapping analysis → layouts, keeping
text within probed box capacities), not the tooling.

## What could be improved

1. **Fix the `/tmp/opencode` mismatch.** Either the environment's permission rules should allow
   the `write` tool there (not just bash), or the skill/system prompt should stop recommending
   it for scripts and say "write build scripts in the working directory" instead. This was the
   session's only failure and it is purely a docs/config inconsistency.
2. **State explicitly that `probe.py` indices == `shapes(page)` order.** The whole build relies
   on this; it worked, but the skill never says it outright. One sentence would remove the
   guesswork.
3. **Document the bar-chart recipe numerically.** I derived `baseline = y + h` and
   `height = value / vmax * max_height` from probe output. A tiny helper (`set_bar(shape,
   baseline, height)`) or a worked example in the skill would make chart retexting a
   copy-paste job.
4. **Provide a pill-fitting helper.** The padding constant (text width + 635) had to be inferred
   from three probe rows. A `fit_pill(doc, page, pill, tag, pad=635)` in `office.template`
   would remove per-template arithmetic.
5. **Spreadsheet `dump.py`: note the formula-syntax normalisation.** Cross-sheet references in an
   `.xlsx` print with LibreOffice dot syntax (`$'Backend API'.B4`) rather than Excel's `!`. It is
   harmless but worth one line in the docs so nobody "fixes" a file that isn't broken.
6. **Suggest `thumbnail.py` for QA of the output deck too.** I rendered 9 PNGs and read them one
   by one (9 tool calls). The skill mentions `thumbnail.py` only for templates; recommending a
   contact sheet of the *final* deck would make visual QA cheaper, with per-slide renders only
   for suspicious slides.
7. **Prompt-side:** the task chain (analyze → present) would benefit from the presentation
   request stating the desired slide count and audience up front. I chose 9 slides to mirror the
   template; an explicit "max N slides, for a steering committee" would tighten the design
   decisions.

## Bottom line

One trivial permissions hiccup, zero tooling failures, both skills delivered as documented. The
scripts are genuinely pleasant to work with; the improvements above are polish, not repairs.
