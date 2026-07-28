# Building the company template default — catalog pipeline

This README covers the **build-time** process that teaches the presentation
skill your company's house template: how to turn your real `.pptx` template into
a *layout catalog* the skill consults to pick the right slide for each kind of
content. It is separate from the runtime skill (`SKILL.md`) — you run this once,
review the result, and the skill uses its output forever after.

**Privacy boundary — read this first.** Every step here runs on your own
infrastructure. The only model that ever sees your real slides is your internal
vision model (gemma4). No slide, render, or template file is sent to any hosted
model. The extraction step (`catalog_extract.py`) uses no model at all — it is
deterministic. Keep the template, the renders, and the finished catalog in your
**internal fork only**; never push them to a remote outside the company.

---

## What you are building

Two artifacts, dropped into the skill:

| Artifact | Path | What it is |
|---|---|---|
| **Asset** | `assets/company-template.pptx` | The template itself — one slide per layout. The skill clones slides from it. |
| **Catalog** | `references/company-template.md` | Prose + a decision table: which *slide of the asset* for which content, and what each holds. The skill reads it to choose designs. |

The skill wiring that consumes them already exists as a draft:
`SKILL-with-template.md`. You activate it in Stage C.

---

## The pipeline at a glance

| Stage | You run | Model? | Produces |
|---|---|---|---|
| **A. Extract** | `catalog_extract.py` | no | `catalog/` — one folder per layout |
| **B. Author** | `catalog_gemma.py` (or opencode) | **gemma** | per-layout entries + a decision table |
| **C. Assemble** | hand-edit + `mv` | no | the wired-in default |
| **D. Calibrate** | build real decks, `qa.py` | no | a catalog that's actually good |

---

## Stage A — Extract  *(done once; you have already run this)*

```bash
python3 scripts/catalog_extract.py company.pptx -o catalog/
```

Deterministic. It runs `probe.py` and `render.py` internally — you do **not**
run those separately. It writes:

```
catalog/
├── manifest.json            one row per layout: number, layout, master, render, structure, slot_count
├── slide-01/
│   ├── render.png           the layout rendered to an image (for the vision model to see)
│   └── structure.json       layout number, master, background, slots[], reading order
├── slide-02/ ...
```

Each `structure.json` `slots[]` entry carries what the model needs to judge
capacity: `index`, `name`, `type`, geometry (`x/y/width/height` in 1/100 mm),
`sample_text`, and `fonts`.

**Sanity-check before spending model time:** open a couple of `render.png`s and
confirm `manifest.json` has one row per real layout. If a layout is missing or
doubled, fix the source template and re-extract — a wrong catalog starts here.

---

## Stage B — Author the catalog with gemma  *(this is the "pass them to gemma" step)*

Two passes: **per-layout** (one call per slide, with its render + structure),
then **aggregate** (one call over all the per-layout answers → the decision
table). The prompts are specified in
[`references/gemma-catalog-prompt.md`](references/gemma-catalog-prompt.md); the
driver below embeds the same text.

You have two ways to run it. Use whichever matches your setup.

### Method 1 — batch driver against your model endpoint  *(recommended)*

`catalog_gemma.py` talks to any **OpenAI-compatible** `/chat/completions`
endpoint — which is what ollama, vLLM, llama.cpp-server and LM Studio all
expose. Stdlib only; no pip.

```bash
# per-layout pass: sends each render.png + structure.json to gemma
python3 scripts/catalog_gemma.py catalog/ \
    --base-url http://localhost:11434/v1 \
    --model gemma3:27b

# aggregate pass: folds catalog/entries/*.json into a decision table
python3 scripts/catalog_gemma.py catalog/ --aggregate \
    --base-url http://localhost:11434/v1 \
    --model gemma3:27b
```

Flags: `--base-url` (default `http://localhost:11434/v1`, i.e. ollama),
`--model` (the vision model name your server knows), `--api-key` (only if your
endpoint requires a bearer token).

**Verify the wiring before you commit to a full run:**

```bash
python3 scripts/catalog_gemma.py catalog/ --dry-run
```

`--dry-run` writes the exact request bodies to `catalog/entries/*.request.json`
(image blob elided) and calls nothing — so you can confirm the endpoint path and
message shape are what your server wants.

Output:
```
catalog/entries/slide-NN.json   one catalogue entry per layout
catalog/decision_table.json     the aggregate (from --aggregate)
```

> **Model name / serving note.** The default `gemma3:27b` is an ollama tag —
> change it to whatever your server calls the vision-capable gemma. It **must**
> accept images; a text-only model will ignore the render and guess layouts from
> structure alone, which is exactly the quality you are trying to avoid. If your
> endpoint speaks ollama's *native* `/api/chat` (with an `images: [base64]`
> field) rather than the OpenAI schema, point `--base-url` at ollama's
> OpenAI-compatible `/v1` path as shown — ollama serves both.

### Method 2 — drive it with opencode

If you would rather have your agent handle the loop and file IO, run opencode in
the `catalog/` directory and give it this instruction (it reads
`render.png` because gemma4 is vision-capable):

> Read `../references/gemma-catalog-prompt.md`. For each `slide-NN/` folder here,
> look at its `render.png` and `structure.json` and produce the per-layout JSON
> object the prompt specifies. Write each to `entries/slide-NN.json`. Then run
> the aggregation prompt over all entries and write `decision_table.json`.

This is more forgiving of a nonstandard serving setup (opencode owns the model
connection), at the cost of being less repeatable than the script.

### What a good per-layout entry looks like

```json
{
  "slide": 5,
  "layout": 20,
  "master": "DEFAULT",
  "name": "Section divider",
  "purpose": "Announce the start of a new section.",
  "content_types": ["section-divider"],
  "capacity": {"headline": true, "body_slots": 0, "max_bullets_per_body": 0,
               "images": 1, "supports_chart": false},
  "use_when": "Between major parts of the deck, to reset the reader.",
  "avoid_when": "Anywhere you need to convey actual content.",
  "notes": "The full-bleed accent bar is fixed decoration — never delete it."
}
```

`slide`, `layout` and `master` are stamped in by the script from
`manifest.json`, not answered by the model. **`slide` is the identifier** — the
design's position in the template deck, i.e. what `clone_slide` takes. `layout`
is the master's autolayout id and is usually the same number on every slide of a
designed template (all nine slides of the Kite fixture report `20`/`DEFAULT`), so
it can't tell two designs apart. The decision table cites `use_slide` for the
same reason, and `--aggregate` warns if a rule names a slide the template lacks.

The model **overstates capacity** — this is its most common failure. You catch
it in Stage C, against the renders.

---

## Stage C — Assemble into the skill

1. **Write the catalog.** Fold `catalog/entries/*.json` and
   `catalog/decision_table.json` into `references/company-template.md`, using
   `references/company-template.example.md` as the skeleton. **Review every
   decision-table row and every capacity figure against the `render.png`s** —
   this human pass is where the catalog earns its keep.
2. **Place the asset.** Copy your template to
   `assets/company-template.pptx` (create the `assets/` folder).
3. **Activate the wiring — do this last.**
   ```bash
   mv SKILL-with-template.md SKILL.md
   ```
   Only after steps 1–2, or the skill will point at files that do not exist yet.

---

## Stage D — Calibrate

Build three or four real briefs with the now-defaulted skill. For each:

```bash
python3 scripts/qa.py out.pptx --template assets/company-template.pptx
python3 scripts/render.py out.pptx        # then look at the slides
```

Where the skill picked a wrong layout or overran a box, correct the offending
row in `references/company-template.md`. Two calibration passes and the default
is genuinely good.

---

## Privacy & git rules

- The template, the renders in `catalog/`, and `references/company-template.md`
  are brand-internal. Commit them to the **internal fork only**.
- `catalog/` is a working directory — you may `.gitignore` it once the catalog
  and asset are written; the committed artifacts are the asset and the reference.
- Nothing in this pipeline calls a hosted model. If a step ever wants network
  access to anything other than your own model endpoint, stop.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `catalog_extract.py` crashes on `AttributeError` for a shape property | A shape type your template has that the probe didn't guard. The known ones (`LineColor`, `FillColor`) are fixed in `probe.py`; if a new one appears, send the traceback line — the same `getattr(..., None)` guard applies. |
| `manifest.json not found` | Point `catalog_gemma.py` at the directory `catalog_extract.py` wrote, not a `slide-NN/` subfolder. |
| Model reply isn't valid JSON | `catalog_gemma.py` already strips ``` fences and extracts the first JSON value. If it still fails it prints the reply head — usually the model added prose; lower temperature is already 0, so shorten the structure by trimming empty slots, or switch to a larger gemma. |
| Layouts all catalogued as "bulleted content" | The model isn't seeing the images (text-only model, or endpoint dropping the `image_url` part). Confirm with `--dry-run` that the image part is present, and that `--model` names a vision model. |
| `no per-layout entries` on `--aggregate` | Run the per-layout pass first; aggregate reads `catalog/entries/*.json`. |
| Capacity figures look too generous | Expected. Fix them by hand in Stage C against the renders — don't trust the model's `max_bullets_per_body`. |

---

## File map

| Path | Role | Runs when |
|---|---|---|
| `scripts/catalog_extract.py` | Stage A — template → `catalog/` dossiers | build-time |
| `scripts/catalog_gemma.py` | Stage B — `catalog/` → entries + decision table via gemma | build-time |
| `references/gemma-catalog-prompt.md` | The prompt spec (human-readable source of truth) | reference |
| `references/company-template.example.md` | Skeleton for the finished catalog | reference |
| `references/company-template.md` | The finished catalog *(you create it)* | runtime |
| `assets/company-template.pptx` | The template asset *(you create it)* | runtime |
| `SKILL-with-template.md` → `SKILL.md` | The wiring that makes the skill use the default | runtime |
