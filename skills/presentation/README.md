# Template packs: the build-time manual

**This is the reference for making and maintaining a pack.** It is run once
per template, on your own infrastructure, and its output is a directory the
skill then uses for every deck. The runtime documentation, how a deck actually
gets built, is `SKILL.md`. The guided, interactive version of this manual is
the **`presentation-setup`** skill: point a coding agent at your template and
it walks through the same stages, asking the questions only a person can
answer. Read this file when you want to drive the scripts yourself, or to
understand what the setup skill wrote.

**Privacy boundary, first.** Every step here runs on your own hardware. The
extraction uses no model at all. The two steps that *can* call a model are
opt-in and say so; point them at a local server and nothing leaves the
machine. `templates/*` is gitignored, so a pack cannot be committed by
forgetting a step.

---

## What a pack is

```
templates/acme/
  template.pptx        the master itself; build_deck duplicates slides out of it
  profile.json         palette, type scale, authoring notes   (make_pack.py)
  catalog.md           which design carries which content    (generated)
  designs/             one file per design: keys, slots      (generated)
  overrides.md         the prose no script can measure        (hand-written)
  history/             the template before each add_designs.py append
  catalog/
    manifest.json      one row per design                     (catalog_extract.py)
    slide-NN/          render.png + structure.json per design (catalog_extract.py)
    entries/*.json     what each design is for                (judged: model or person)
    decision_table.json  content situation -> design         (judged)
    designs.md         the worksheet, if you filled one       (name_designs.py)
    added.json         designs appended from project decks     (add_designs.py)
```

Two halves feed the catalog, and keeping them apart is the point:

| Half | Written by | Contains |
|---|---|---|
| **measured** | `catalog_extract.py`, deterministically | retext keys, empty slots, group counts, canvas hint, title budgets, geometry |
| **judged** | a person or the vision-agent (worksheet) | name, purpose, use/avoid, capacity, imagery, the decision table |

The first version of this pipeline asked a model for the catalog wholesale.
Everything that made the result usable was added by hand afterwards, and none
of it survived to the next template. All of it is transcription or arithmetic,
so the extractor measures it and nobody is asked.

`templates/README.md` says how a pack is chosen at build time. `references/catalog.example.md`
shows the generated format with each field marked measured or judged.

---

## Stage A: palette, type scale, authoring notes

```bash
SETUP=../presentation-setup/scripts          # the setup skill's scripts
python3 scripts/thumbnail.py brand.pptx      # contact sheet: count the designs
python3 $SETUP/count_style.py brand.pptx     # every colour and size the file uses
python3 $SETUP/find_notes.py brand.pptx      # strings that look like builder notes
```

Then write what a person decided:

```bash
python3 $SETUP/make_pack.py --name acme --template brand.pptx \
    --colour accent=#RRGGBB --colour alert=#RRGGBB --colour ok=#RRGGBB \
    --colour muted=#RRGGBB --colour neutral=#RRGGBB \
    --series '#RRGGBB,#RRGGBB,#RRGGBB' \
    --size display=60 --size title=16 --size subtitle=14 --size body=14 \
    --size label=12 --size caption=10 --size footnote=9 \
    --note 'EXAMPLE' --note-prefix 'Instructions: 1.' \
    --default
```

That creates `templates/acme/` with the template copied in and `profile.json`
written. `make_pack.py` refuses a colour the template does not contain, and
re-running it merges: anything not on the command line is kept from the
profile on disk.

`profile.json` is plain JSON if you would rather write it yourself:

```json
{
  "name": "acme",
  "asset": "template.pptx",
  "annotations": {"exact": ["EXAMPLE"], "prefixes": ["Instructions: 1."]},
  "palette": {
    "_order": ["blue", "orange", "grey"],
    "blue": "#2A7DE1", "orange": "#F28C28", "grey": "#5A5A5A",
    "_roles": {"accent": "blue", "alert": "orange", "ok": "blue",
               "muted": "grey", "neutral": "grey",
               "series": ["blue", "orange", "grey"]}
  },
  "type_scale": {"display": 60, "title": 16, "subtitle": 14, "body": 14,
                 "label": 12, "caption": 10, "footnote": 9}
}
```

`D.color("accent")` resolves through `_roles`; `role="body"` through
`type_scale`; `strip_annotations(page)` deletes exactly the `annotations`.

---

## Stage B: extract

```bash
python3 scripts/catalog_extract.py brand.pptx -o templates/acme/catalog
```

Deterministic, no model, no network, about thirty seconds. One folder per
design with its render and a `structure.json`:

| Field | What it is |
|---|---|
| `slots` | every text-bearing shape: index, name, type, geometry, sample text, fonts |
| `retext_keys` | **every string on the slide with the address that carries it**, group children included (`6.2`), truncated the way a catalog quotes them |
| `empty_slots` | shapes with no text; `retext` can never match these |
| `text_counts` | how many text shapes sit inside groups versus at the top level |
| `title_budgets` | per display title slot: box width, point size, characters that fit on one line |
| `canvas` | the `keep=` tuple `draw.canvas` would use and the free `Box` it would return |
| `images` | picture shapes with geometry |
| `reading_order` | shapes grouped into columns, since z-order is not reading order |

Hidden slides are rendered too, from an all-visible copy, so image N is
always design N. Re-running the extraction into the same directory refreshes
the measured half and leaves `entries/`, `decision_table.json` and
`designs.md` alone.

**Sanity-check before judging:** open a couple of renders and confirm
`manifest.json` has one row per real design. A design missing or doubled here
is wrong in every later stage.

---

## Stage C: judge

The judged half is what each design is *for*. Two ways to write it.

### Method 1: a worksheet, no model

```bash
python3 $SETUP/name_designs.py templates/acme/catalog            # writes designs.md
# fill purpose / use_when / avoid_when per design, beside the renders
python3 $SETUP/name_designs.py templates/acme/catalog --import   # -> entries/, decision_table.json
```

Every measured field is pre-filled; only judgement is asked for. Twenty
designs is about five minutes with the contact sheet open, and a person does
this better than a 31B model did. A design imported before comes back with
what was approved, not with the vision-agent's proposal. `--only 32,33` writes
the worksheet for those designs alone, and the import then leaves every other
entry exactly as it was. The optional `table_row` field gives a design its own
row at the top of the decision table (stage D). The worksheet's `## Canvas` block lets you
name the canvas design and its `keep` tuple by hand when the measured pick is
wrong (stage E shows you).

### Method 2: the vision-agent

The same worksheet, prefilled by the vision-agent (the `vision` subagent),
one design per call, from a generated brief:

```bash
python3 scripts/vision_brief.py designs templates/acme/catalog
# per design:  Read the brief at <catalog>/slide-NN/brief.md and follow it exactly.
# save reply:  <catalog>/slide-NN/answer.md
python3 $SETUP/name_designs.py templates/acme/catalog     # worksheet, prefilled
```

The brief asks only for what needs eyes — name, purpose, content shape, item
limit, when to use and avoid it, what its pictures depict — never for retext
keys, indices, budgets or colours, which are measured. The vision-agent runs
on your own infrastructure and the renders never leave the machine. Review the
prefilled worksheet with the user — the vision-agent proposes, the user
approves — then the same import as Method 1:

```bash
python3 $SETUP/name_designs.py templates/acme/catalog --import
```

---

## Stage D: assemble

```bash
python3 scripts/catalog_build.py templates/acme/catalog
```

Inside a pack this needs no flags: it finds `template.pptx`, splices in
`overrides.md` if there is one, and writes `catalog.md`. Elsewhere pass
`--asset`, `--overrides` and `-o`.

The catalog is: header, designs by role (canvas, dividers, cover), the
decision table with the canvas row **last**, variety rule, coverage gaps,
display-title budgets, palette and type scale, and one entry per design. It
flags rather than hides what went wrong: a row citing a design the template
does not have gets a ⚠, a design with no judged entry reads *(not
catalogued)*, an override anchor that does not land is listed at the foot.

The decision table has two kinds of row. Designs with a `table_row` come
**first**, one row each, because first match wins and a specific row ("a live
demo, step by step with commands") must be read before the generic one that
would also match. Then one generic row per content situation picks one design
and lists the others as alternatives **with their content shape**, so a builder
can choose between them without opening every design file.

**Then review it against the renders.** Every capacity figure and every
decision-table row. Capacity is what a model overstates most.

### Review notes go in `overrides.md`, never in `catalog.md` or `designs/`

The catalog is regenerated from `catalog/` every time; edits to it are lost
on the next run. Everything you want to keep goes in `templates/acme/overrides.md`
as a `## anchor` section:

| Anchor | Lands |
|---|---|
| `header` | after the generated asset/identity paragraphs |
| `after-decision-table` | after the table, variety rule and coverage gaps |
| `before-design-entries` | after the display-title budgets |
| `design-entries-intro` | after the "Retext keys are…" preamble |
| `design-N` | at the end of that design's entry |
| `footer` | after the last entry, before the provenance note |

Blocks add to the generated text rather than replacing it. An anchor that no
longer lands prints a warning and appears as a ⚠ line in the catalog, so
prose cannot go quietly missing.

### Drift check

```bash
python3 scripts/catalog_extract.py templates/acme/template.pptx -o templates/acme/catalog
python3 scripts/catalog_build.py templates/acme/catalog --check
```

`--check` writes nothing: it rebuilds, diffs against `catalog.md`, and exits 1
with the diff if they differ. Run it whenever the template file changes. A
drifted catalog is not cosmetic: its retext keys are what a build script
copies, and a missing key is a template label left on a finished slide.

### Recovering a catalog that lost its judged half

If `entries/` was deleted, or a catalog was written by hand before this
pipeline existed, `catalog_import.py` reads the judged fields back out of the
markdown so the pack can be rebuilt:

```bash
python3 scripts/catalog_extract.py templates/acme/template.pptx -o templates/acme/catalog
python3 scripts/catalog_import.py templates/acme/catalog.md templates/acme/catalog
python3 scripts/catalog_build.py templates/acme/catalog --check
```

It recovers only the judged fields. The measured ones are deliberately
ignored: the hand-maintained file is exactly where those went stale.

---

## Stage E: probe, then calibrate

```bash
python3 $SETUP/probe_deck.py acme          # three real slides -> probe.pptx
python3 scripts/render.py probe.pptx       # -> render/slide-1.png ...
```

A divider, a content slide retexted through the catalog's own keys, and a
drawn slide on the canvas design in the pack's colours. Look at all three. A
wrong colour means the roles in stage A were mapped wrong; leftover wording
means a missed authoring note; furniture showing through the drawing means the
canvas design is wrong, which the worksheet's `## Canvas` block fixes.

Then build three or four real briefs with the now-defaulted skill:

```bash
python3 scripts/make_deck.py build.py out.pptx --template acme
```

and open every render it prints. Every wrong design choice gets corrected
**in the catalog** (through `overrides.md` or the worksheet), not in the build
script: the point is that the next deck picks better without being told. Two
passes and the default is genuinely good.

While the build is still failing, iterate with `--build-only` or `--skip-qa`
(no report is written, so a not-yet-checked deck never looks clean); the full
command above is the gate and is run once. `--require-verdicts` exits 1 unless
every slide has a verdict in the report's table.

Calibrating is the pass most worth having eyes on, because it judges looks
rather than correctness: build the deck, then **delegate every render to the
vision-agent** (the `vision` subagent) and correct what it flags in the
catalog. Nothing here calls a hosted model; the renders never leave the
machine.

---

## Stage F: add designs from project decks

A pack can grow after setup. Custom slides that people built on the company
master for one project become designs of the pack, numbered after the last
one, so no existing number moves. `.pptx`/`.potx` packs only.

```bash
python3 $SETUP/add_designs.py candidates acme project1.pptx project2.pptx
python3 $SETUP/add_designs.py texts acme project1.pptx --slides 21,8
# fill every "new" in add-designs/project1.texts.json with neutral sample text
python3 $SETUP/add_designs.py append acme add-designs/project1.texts.json
python3 scripts/catalog_extract.py templates/acme/template.pptx -o templates/acme/catalog
python3 scripts/vision_brief.py designs templates/acme/catalog --only 32,33
python3 $SETUP/name_designs.py templates/acme/catalog --only 32,33
python3 $SETUP/name_designs.py templates/acme/catalog --import
python3 scripts/catalog_build.py templates/acme/catalog
```

- **`candidates`** compares where each slide's top-level shapes sit with every
  design of the pack. It reads XML only, takes about a second, and calls no
  model. It writes `add-designs/candidates.md`: slides no design matches,
  grouped when they look alike; slides that are a design already; slides on
  another master, which cannot be added. A person picks one slide per
  candidate.
- **`texts`** lists every paragraph of the chosen slides with a blank `new`.
  The project's wording would become the retext keys and the budgets of every
  later deck, so it is replaced with sample text of about the same length.
  Footer, date and slide-number placeholders are left alone. Pictures, chart
  data and embedded objects travel as they are, and the command lists them.
- **`append`** copies the slides into `template.pptx` at the `.pptx` level,
  with their media, charts and embedded objects, and rebinds each to the
  template's layout of the same name. Notes and comments are dropped.
  LibreOffice never saves the template, because every `.pptx` save through it
  adds one more master. The template before is kept in `history/`, and
  `catalog/added.json` records each design's origin. `append` refuses:
  - an unfilled `new`;
  - sample text that makes two different originals equal or one a substring
    of the other, since a retext key is a substring;
  - a deck that changed since `texts`;
  - a slide that is a design already.

---

## Privacy and git

- `templates/*` is gitignored (`templates/README.md` excepted). The template,
  the renders, the catalog and the profile are brand-internal and stay on the
  machine that built them.
- If your company keeps packs in a private fork, un-ignore the pack directory
  there and commit **all of it**, `catalog/entries/*.json` and
  `catalog/decision_table.json` included: they cost a model call or a human
  review each, and nothing can recreate them from the template. The renders
  and `structure.json` files are regenerated in thirty seconds, so leaving
  those out is fine.
- Nothing here calls a hosted model by default. The only step that needs eyes
  — judging a design or reviewing a render — is the vision-agent, which runs on
  your own infrastructure; no script reads another tool's credential store.
  Per-deck visual QA and catalog judging keep the renders on the machine.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `catalog_extract.py` crashes on `AttributeError` for a shape property | A shape type the probe did not guard. The known ones (`LineColor`, `FillColor`) are handled in `probe.py`; a new one wants the same `getattr(..., None)` guard. |
| `no manifest.json in …` | Point the script at the directory `catalog_extract.py` wrote, not at a `slide-NN/` subfolder. |
| `--asset: no template found beside …` | The catalog directory is not inside a pack. Pass `--asset path/to/template.pptx`. |
| `make_pack.py` refuses a colour | It is not on any slide of the template. Take it from `count_style.py`; `--allow-unseen` for a colour that lives only in the master. |
| The vision-agent's judgement is generic, or praises a broken slide | It is not seeing the renders, or is a text-only model. Confirm it actually reads the `render.png` you delegate to it. |
| `add_designs.py` lists slides of a deck you know is on the master under *Not on this master* | Their layout name is not in the template: the deck was built on another copy of the master, with layouts renamed. Rename the layout in the deck, or add the slide to a pack built on that copy. |
| `add_designs.py append` refuses with "now collides" | Two paragraphs got sample text where one is inside the other. A retext key is a substring, so it would hit both. Give each its own wording. |
| `vision_brief.py designs` lists every design after adding two | A re-extraction re-renders everything, and LibreOffice renders do not always come out byte-identical. Pass `--only` with the new design numbers. |
| A vision finding looks wrong and the string it names is nowhere in the file | Check *how* you looked. A footer is a page property, not a shape, so a shape search cannot see one. `wireframe.py` and `probe.py` settle most of the rest; a finding nothing can settle stands. |
| No vision-agent is available for the visual half of QA | `wireframe.py` over `render/` is the fallback; say which slides nothing with eyes has seen. |
| `no per-layout entries` on `--aggregate` | Run `name_designs.py --import` first; aggregate reads `catalog/entries/*.json`. |
| Capacity figures look too generous | Expected. Fix them in `overrides.md`, or in `catalog/entries/slide-NN.json`, and rebuild. |
| My edits to `catalog.md` disappeared | It is generated. Put them in `overrides.md` under an anchor and rebuild. |
| `--check` reports drift after the template changed | Expected, and the point of it. Re-extract, rebuild, read the diff. |
| An override block "matched no anchor" | Misspelled, or a `design-N` for a design this template no longer has. Valid names are in stage D. |
| LibreOffice dies mid-build with `free(): invalid pointer` or `Binary URP bridge disposed` | A LibreOffice 25.2.x race on released text proxies. The workaround is on by default; check that `PRESENTATION_PIN_TEXT_PROXIES` is not set to `0`. |
| The first QA run after touching the template is a minute slower | The template baseline is recomputed and cached under `$XDG_CACHE_HOME/presentation-skill/`. Later runs are about 40 s faster. |
| The canvas row picked a design you would not have | It picks the largest free area whose canvas keeps the design's chrome. Name the right one in the worksheet's `## Canvas` block (or a `canvas` object in `decision_table.json`), re-import, rebuild. |

---

## File map

| Path | Role | Runs when |
|---|---|---|
| `../presentation-setup/scripts/count_style.py` | Stage A: the colours and sizes a template uses | build-time |
| `../presentation-setup/scripts/find_notes.py` | Stage A: candidate authoring notes | build-time |
| `../presentation-setup/scripts/make_pack.py` | Stage A: write the pack and its `profile.json` | build-time |
| `scripts/catalog_extract.py` | Stage B: template to `catalog/` dossiers | build-time |
| `../presentation-setup/scripts/name_designs.py` | Stage C without a hosted model: worksheet in, entries out | build-time |
| `scripts/catalog_gemma.py` — *removed* | Used to judge dossiers with a hosted vision model; that job is now the vision-agent filling the `name_designs.py` worksheet | build-time |
| `scripts/catalog_build.py` | Stage D: both halves plus overrides to `catalog.md` (the index) and `designs/design-N.md` (one per design); `--check` for drift | build-time |
| `scripts/catalog_import.py` | Recovery: a hand-written catalog to the entries it never kept | once, if needed |
| `../presentation-setup/scripts/harvest_icons.py` | Stage C: the template's own icons into the pack's `icons/` | build-time |
| `../presentation-setup/scripts/probe_deck.py` | Stage E: three real slides on the new pack | build-time |
| `../presentation-setup/scripts/add_designs.py` | Stage F: custom slides from project decks become designs (`candidates`, `texts`, `append`) | build-time |
| `scripts/review_renders.py` — *removed* | The per-deck visual reviewer used to send renders to a hosted model; that job is now the vision-agent, which the build agent delegates to. |
| `scripts/vision_brief.py` | Writes the tested vision-agent brief beside every render or design, and collects the saved answers into the report or the worksheet | runtime + build-time |
| `scripts/split_source.py` | A long source (a slide spec) as one file per slide, for big decks built from parts (`references/big-decks.md`) | runtime |
| `scripts/api_ref.py` | The build API from the code; `--write` regenerates `references/api.md` (a test fails when it is stale) | runtime + build-time |
| `scripts/wireframe.py` | A render as a text ink-map, no model; the fallback when the vision-agent is unavailable | runtime |
| `references/catalog.example.md` | The generated format, annotated | reference |
| `templates/README.md` | What a pack is and how one is chosen | reference |
| `SKILL.md` | The runtime documentation: how decks get built | runtime |
