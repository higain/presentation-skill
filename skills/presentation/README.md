# Teaching this skill your company template — the catalog pipeline

**This is the build-time manual.** It is run once per template, on your own
infrastructure, and its output is two files the skill then uses forever after.
The runtime documentation — how a deck actually gets built — is `SKILL.md`.

**Privacy boundary — read this first.** Every step here runs on your own
hardware. The only model that ever sees your real slides is your internal vision
model. No slide, render, or template file is sent to a hosted model. The
extraction step uses no model at all — it is deterministic. Keep the template,
the renders and the finished catalog in your **internal fork only**; never push
them to a remote outside the company.

---

## What you are building

Three artifacts, dropped into the skill:

| Artifact | Path | What it is |
|---|---|---|
| **Asset** | `assets/company-template.pptx` | The template itself — one slide per design. `build_deck` duplicates slides out of it. |
| **Catalog** | `references/company-template.md` | Which design for which content, what each design holds, and the exact strings to `retext`. The skill reads this to choose designs. |
| **Profile** | `assets/company-template.profile.json` | The template's own authoring annotations — an "example" flag, a customer-logo marker — so `strip_annotations` can delete them without the engine knowing any company's wording. |

No wiring step: `SKILL.md` already points at those paths. Drop the files in and
the default is live.

## The pipeline at a glance

| Stage | You run | Model? | Produces |
|---|---|---|---|
| **A. Extract** | `catalog_extract.py` | no | `catalog/` — one dossier per design, with everything measurable already measured |
| **B. Judge** | `catalog_gemma.py` (or opencode) | **yes, locally** | per-design purpose, capacity, imagery + a decision table |
| **C. Assemble** | `catalog_build.py` | no | `references/company-template.md`, ready to review |
| **D. Calibrate** | `make_deck.py` on real briefs | no | a catalog that is actually good |

**The split between A and B is the point.** The first version of this pipeline
asked the model for the catalog wholesale. What made the resulting catalog
usable was everything added by hand afterwards — the retext keys, the empty
slots, the group counts, the canvas indices, the title budgets — and none of
that survived into the skeleton, so the next template would have lost it. All of
it is transcription or arithmetic, so stage A measures it and the model is never
asked.

---

## Stage A — Extract

```bash
scripts/catalog_extract.py company.pptx -o catalog/
```

Deterministic, no model, no network. It writes:

```
catalog/
├── manifest.json            one row per design: number, layout, master, render, slots
├── slide-01/
│   ├── render.png           the design rendered, for the vision model to look at
│   └── structure.json       slots, retext keys, empty slots, canvas hint, budgets
├── slide-02/ ...
```

Each `structure.json` carries, per design:

| Field | What it is |
|---|---|
| `slots` | every text-bearing shape: index, name, type, geometry, sample text, fonts |
| `retext_keys` | **every string on the slide with the address that carries it**, group children included (`6.2`), truncated the way a catalog quotes them |
| `empty_slots` | shapes with no text — `retext` can never match these, and a build that assumes otherwise wastes an attempt finding out |
| `text_counts` | how many text shapes are inside groups versus at the top level |
| `title_budgets` | per title slot: box width, point size, and characters that fit on one line |
| `canvas` | the `keep=` tuple `draw.canvas` would use and the free `Box` it would return |
| `images` | picture shapes with geometry — the model only has to say what they depict |
| `reading_order` | shapes grouped into columns, since z-order is not reading order |

Hidden slides are rendered too (from an all-visible copy), so image N is always
design N.

**Sanity-check before spending model time:** open a couple of `render.png`s and
confirm `manifest.json` has one row per real design. If a design is missing or
doubled, fix the source template and re-extract — a wrong catalog starts here.

---

## Stage B — Judge, with a local vision model

Two passes: **per-design** (one call per slide, with its render and structure),
then **aggregate** (one call over all the per-design answers, producing the
decision table). The prompts are in
[`references/gemma-catalog-prompt.md`](references/gemma-catalog-prompt.md); the
driver below embeds the same text.

The model is asked for judgement only: name, purpose, content types, capacity,
use/avoid, what the pictures depict, and quirks. It is told explicitly **not**
to produce retext keys, indices, geometry or budgets — those are already
measured, and a model asked for them invents plausible wrong ones.

### Method 1 — batch driver against your model endpoint *(recommended)*

`catalog_gemma.py` talks to any **OpenAI-compatible** `/chat/completions`
endpoint — ollama, vLLM, llama.cpp-server and LM Studio all expose one. Stdlib
only; no pip.

```bash
# per-design pass: sends each render.png + structure.json to the model
scripts/catalog_gemma.py catalog/ \
    --base-url http://localhost:11434/v1 \
    --model gemma3:27b

# aggregate pass: folds catalog/entries/*.json into a decision table
scripts/catalog_gemma.py catalog/ --aggregate \
    --base-url http://localhost:11434/v1 \
    --model gemma3:27b
```

Flags: `--base-url` (default `http://localhost:11434/v1`), `--model` (the vision
model name your server knows), `--api-key` (only if your endpoint needs one).

**Verify the wiring before a full run:**

```bash
scripts/catalog_gemma.py catalog/ --dry-run
```

`--dry-run` writes the exact request bodies to `catalog/entries/*.request.json`
(image blob elided) and calls nothing, so you can confirm the endpoint path and
message shape are what your server wants.

Output:
```
catalog/entries/slide-NN.json   one entry per design
catalog/decision_table.json     the aggregate (from --aggregate)
```

**No local GPU?** `--openrouter` points the same two passes at OpenRouter,
reading `OPENROUTER_API_KEY` from the environment:

```bash
OPENROUTER_API_KEY=... scripts/catalog_gemma.py catalog/ \
    --openrouter --model google/gemma-4-31b-it
```

That is a hosted model seeing your slides, so it is outside the privacy boundary
above — a deliberate exception, for a template you are willing to send.

> **Model choice.** The default `gemma3:27b` is an ollama tag — change it to
> whatever your server calls its vision-capable model. It **must** accept
> images; a text-only model ignores the render and guesses layouts from
> structure alone, which is exactly the quality you are trying to avoid. If your
> endpoint speaks ollama's native `/api/chat` rather than the OpenAI schema,
> point `--base-url` at ollama's `/v1` path as shown — ollama serves both.

### Method 2 — drive it with opencode

If you would rather your agent handled the loop and the file IO, run opencode in
the `catalog/` directory and give it:

> Read `../references/gemma-catalog-prompt.md`. For each `slide-NN/` folder
> here, look at its `render.png` and `structure.json` and produce the per-layout
> JSON object the prompt specifies. Write each to `entries/slide-NN.json`. Then
> run the aggregation prompt over all entries and write `decision_table.json`.

More forgiving of a nonstandard serving setup, less repeatable than the script.

---

## Stage C — Assemble

```bash
scripts/catalog_build.py catalog/ -o references/company-template.md
```

Merges the measured fields from stage A with the judged fields from stage B and
writes the catalog in its shipped format: header, decision table (with the
canvas row first), variety rule, coverage gaps, display-title budgets, and one
entry per design. `references/company-template.example.md` shows that format
annotated with which half each field comes from.

It flags rather than hides what went wrong:

- a decision-table row citing a slide the template does not have gets a ⚠ row;
- a design with no model entry reads _(not catalogued)_ and is listed at the
  foot as incomplete;
- a design whose imagery count disagrees with the model's description says so.

**Then review it against the renders.** Every capacity figure and every
decision-table row. The model overstates capacity — that is its most common
failure and no script can catch it. This human pass is where the catalog earns
its keep.

Finally, place the other two artifacts:

```bash
cp company.pptx assets/company-template.pptx
$EDITOR assets/company-template.profile.json     # the template's own annotations
```

The profile is a short JSON file listing the strings the template uses to talk
to whoever builds from it — an "example" flag, a "customer logo here" marker, a
build-instruction block. `strip_annotations(page)` deletes exactly those, and
nothing else in the engine knows any of your wording. Stage A's `retext_keys`
across all designs is where you find them: a string that appears on many
unrelated designs is almost always an annotation rather than content.

---

## Stage D — Calibrate

Build three or four real briefs with the now-defaulted skill. For each:

```bash
scripts/make_deck.py build.py out.pptx --template assets/company-template.pptx
```

then open every render it prints. `make_deck.py` covers what `qa.py` and
`render.py` used to do separately, and writes `<deck>-report.md` beside the deck.

Calibrating is the pass most worth having a vision model on, because it is the
one where you are judging *looks* rather than correctness:

```bash
export PRESENTATION_REVIEW_MODEL=google/gemma-4-31b-it   # must accept images
export OPENROUTER_API_KEY=...
```

Every `make_deck.py` run then also sends the renders to that model and prints
what it saw. It is advisory and never fails a build — see *The vision reviewer*
in `references/qa.md`.

A calibration pass is finished when: every finding in the report is fixed and
the run comes back clean; every render has been looked at; and every wrong
design choice has been corrected **in the catalog**, not in the build script —
the whole point is that the next deck picks better without being told.

Two passes and the default is genuinely good.

---

## Privacy & git rules

- The template, the renders in `catalog/`, `references/company-template.md` and
  `assets/company-template.profile.json` are brand-internal. Commit them to the
  **internal fork only**.
- `catalog/` is a working directory — `.gitignore` it once the catalog and the
  asset are written; the committed artifacts are the asset, the profile and the
  catalog.
- Nothing in this pipeline calls a hosted model *by default*. Two steps can, and
  both are opt-in and explicit: `catalog_gemma.py --openrouter`, and
  `review_renders.py` once `$PRESENTATION_REVIEW_MODEL` is set. If any other
  step wants network access to anything but your own endpoint, stop.
- **The vision reviewer sends slide images off the machine** when it is pointed
  at a hosted endpoint — that is a rendered picture of the deck, content and
  all, going to a third party. For a confidential deck, point
  `$PRESENTATION_REVIEW_BASE_URL` at a local server (ollama, vLLM, llama.cpp)
  and the images never leave. Unset `$PRESENTATION_REVIEW_MODEL` and nothing is
  sent at all.
- Keys are read from the environment (`OPENROUTER_API_KEY`, `OPENAI_API_KEY`)
  or passed explicitly. No script here reads another tool's credential store.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `catalog_extract.py` crashes on `AttributeError` for a shape property | A shape type your template has that the probe did not guard. The known ones (`LineColor`, `FillColor`) are handled in `probe.py`; a new one wants the same `getattr(..., None)` guard. |
| `manifest.json not found` | Point `catalog_gemma.py`/`catalog_build.py` at the directory `catalog_extract.py` wrote, not at a `slide-NN/` subfolder. |
| Model reply is not valid JSON | `catalog_gemma.py` strips ``` fences and extracts the first JSON value. If it still fails it prints the reply head — usually the model added prose. Temperature is already 0; trim empty slots from the structure, or use a larger model. |
| Every design catalogued as "bulleted content" | The model is not seeing the images. Confirm with `--dry-run` that the image part is present and that `--model` names a vision model. |
| The vision review is generic, or praises slides that are broken | Same cause: a text-only model. It will answer confidently about an image it never received. Check the model's input modalities — `deepseek-v4-flash` is text-only, `gemma-4-31b-it` and `qwen3.6-27b` are not. |
| `make_deck.py` prints "No visual review" | `$PRESENTATION_REVIEW_MODEL` is unset. That is the default; set it to enable the reviewer. |
| `no per-layout entries` on `--aggregate` | Run the per-design pass first; aggregate reads `catalog/entries/*.json`. |
| Capacity figures look too generous | Expected. Fix them by hand in stage C against the renders. |
| A ⚠ row in the built catalog | The model cited a slide number the template does not have. Correct the row or delete it. |
| The canvas row picked a design you would not have | It picks the largest free area whose canvas keeps the design's chrome and drops one body shape. Override it by editing the row; if it is wrong for your template, say so and the heuristic can be tightened. |

---

## File map

| Path | Role | Runs when |
|---|---|---|
| `scripts/catalog_extract.py` | Stage A — template → `catalog/` dossiers | build-time |
| `scripts/catalog_gemma.py` | Stage B — dossiers → judged entries + decision table | build-time |
| `scripts/catalog_build.py` | Stage C — both halves → the catalog | build-time |
| `scripts/review_renders.py` | The vision reviewer — renders → advisory findings | runtime, when `$PRESENTATION_REVIEW_MODEL` is set |
| `references/gemma-catalog-prompt.md` | The prompt spec (human-readable source of truth) | reference |
| `references/company-template.example.md` | The finished format, annotated | reference |
| `references/company-template.md` | The catalog *(generated)* | runtime |
| `assets/company-template.pptx` | The template asset *(you copy it in)* | runtime |
| `assets/company-template.profile.json` | The template's authoring annotations | runtime |
| `SKILL.md` | The runtime documentation — how decks get built | runtime |
