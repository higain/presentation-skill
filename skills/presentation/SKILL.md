---
name: presentation
description: "Create, edit, read and QA presentations — .odp and .pptx alike, plus .otp/.potx templates and legacy .ppt. Use whenever a slide deck is involved: building a deck or pitch, filling in a template, editing or restyling an existing one, extracting text or speaker notes, converting between formats, or rendering slides to images for visual review. Triggers on \"PowerPoint\", \"Impress\", \"deck\", \"slides\", \"presentation\", \"Präsentation\", \"Folien\", or any .odp/.otp/.pptx/.potx/.ppt filename. The user's word for a file says nothing about its extension — a file someone calls \"the PowerPoint\" is often an .odp — so check the actual file, and save a deck back in the format it arrived in."
compatibility: "Needs LibreOffice Impress 25.2+ with python3-uno, plus python3-lxml, python3-defusedxml, python3-pil, poppler-utils, util-linux 2.38+ (unshare) or bubblewrap for running LibreOffice without network, and Python 3.11+. All dependencies come from apt — never run pip or npm. Works offline. Visual QA never sends a render anywhere: the build agent delegates the renders to a vision-agent, and company slides never leave the machine."
license: Apache-2.0 (see LICENSE at the repository root; third-party notices in THIRD-PARTY.md)
---

# Presentations: .odp and .pptx

The engine is LibreOffice Impress, driven over UNO: it opens the real file and
saves it in its own format. **Every path below is relative to this skill's
directory** (the one holding this file, as you were given it). Run a script by
its full path, `python3 <skill dir>/scripts/dump.py deck.pptx`, from the user's
folder, never `cd` into the skill and `scripts/...`: permission rules match the
command as typed, and a relative path can stop the run at a prompt nobody sees.

| Task | How |
|---|---|
| Read a deck | `scripts/dump.py deck.pptx` |
| Inspect shapes before editing | `scripts/probe.py deck.pptx --slides 4,9` |
| Edit an existing deck | `impress_document(path)` — saves in place, same format (`references/templates.md`) |
| Convert | `scripts/convert.py in.odp out.pptx` |
| **Build a deck** | the recipe below |
| Teach the skill a company template | the `presentation-setup` skill |

**Rules that override everything below**

1. Keep the format a deck arrived in; a new deck is `.pptx` unless asked otherwise.
2. Build on the company template pack — the user's own template if they bring
   one. Never invent colours, fonts or a look (`D.color()` raises without a pack).
3. Write in the user's language, titles included.
4. You cannot see images. Every look at a slide goes to the **vision-agent**
   (the `vision` subagent) through `scripts/vision_brief.py`. Never describe a
   slide nothing with eyes has seen.

## Step 0 — what to build

| You were handed | Start at |
|---|---|
| The content: an outline, a document, slide text, data to present | the recipe |
| Only a subject ("a deck about X", a goal, an audience) | `references/storyline.md` — intake, storyline, a ghost deck the user approves; then the recipe |

If you would have to invent what the slides say, it is the second row.

**A big deck** — more than about 20 slides, or a source over about 50 KB
(`wc -c`): do not read the source whole and do not write one `build.py`.
Follow `references/big-decks.md`: split the source, one file per slide,
written by builder subagents, assembled by `build_parts`.

## The recipe

`acme` is the pack name: the one in `templates/default`, or the only directory
in `templates/`. No pack installed → ask for the template and run
`presentation-setup`.

**1. Read the pack's `templates/acme/catalog.md`.** It is short: which design
for which content, a design index with each design's *content shape*, how each
design's headline is written, the palette. Read `references/language.md` too
if you write the slide text yourself.

**2. Plan the deck as a table** before any code — one row per slide:

| slot | design | why it fits (content shape) | headline (a full-sentence claim) |
|---|---|---|---|

Rules for the plan:
- **Count the pieces before you pick the design.** Three risks, two options,
  five steps: write the count in the row. Match it to a design's *content
  shape* and item limit; two items do not go into a three-column design, and
  seven do not go into a four-card one — split the slide instead.
- **Vary the designs.** QA fails a deck of six or more slides where one
  design — the canvas included — carries more than half of them.
- Content no design fits (a diagram, chart, comparison, timeline you draw)
  goes on the **canvas design** (catalog, *Designs by role*), stripped back.
- Every content slide needs one thing to look at — a diagram, chart, big
  number, table, image — not only text (`references/drawing.md`, Composition).
- A headline states the slide's conclusion as a sentence, not a topic label.
- **Read the finished plan once more.** A row that would look the same in any
  other deck on this template (title, bullets, nothing to look at) gets a
  different design or a drawing before any code is written.

**3. Read `templates/acme/designs/design-N.md` for every design in the plan.**
It lists the design's **retext keys** (the exact template wording to replace),
its empty slots and its budgets. Keys differ per design: copy them from that
file, never from memory or another design.

**4. Write one `build.py`.** One `fill` call per slot does all retexting and
drawing. Branch on `slot` (position in your deck), never on `design`, and end
with an `else` that raises:

```python
from office.deckbuild import build_deck
from office.template import retext, set_title
from office.impress import mm
from office import draw as D

def fill(doc, page, slot, design):
    if slot == 0:                                   # a designed slide
        set_title(page, "Billing is blocked by one vendor key")   # if the index says set_title
        retext(page, {"<key from design-N.md>": "New text",
                      "<another key>": None})       # None deletes that shape
    elif slot == 1:                                 # the canvas design
        set_title(page, "Three levers secure the March launch")
        area = D.canvas(page, keep=(0, 1, 2))       # keep tuple from its design file
        D.cards(doc, page, area, [("Vendor keys", "Escalate by week 42."),
                                  ("Interface", "Freeze the API contract."),
                                  ("Test window", "Book the QA slot.")],
                numbered=True)
    else:
        raise ValueError(f"no branch for slot {slot}")

build_deck("acme", "out.pptx", [4, 2], fill=fill,
           footer="Aurora · Steering committee · 2026")   # None hides it
```

- **Headlines:** write each one the way the index's *Headline* column says —
  `set_title(page, ...)` when the title placeholder is empty (`retext` cannot
  reach an empty shape), otherwise the listed retext key.
- **Every shape that still carries template text must be retexted or deleted**
  — including labels the canvas keeps (the catalog warns about them).
- **The footer** is a page property: set it with `footer=` in `build_deck`.
- **Draw with the composition helpers**, laid out from the `Box` that
  `D.canvas` returns: `D.cards`, `D.steps`, `D.kpis`, `D.statement`, `D.table`,
  `D.venn`, `D.bullets`, `D.chart_with_comment`, `D.source`, and the
  consulting components `D.tracker`, `D.pairs`, `D.levels`, `D.matrix2x2`,
  `D.quote`, `D.kicker`, `D.sticker` (`references/drawing.md`
  §Consulting components). A slide built from these comes out finished;
  one built from `D.rect` and `D.text` comes out however the arithmetic
  went. Colours by name (`D.color("accent")`, `D.series(3)`),
  sizes by role (`role="body"`). Never hardcode millimetre offsets or hex.
  Details: `references/drawing.md`.
- **Never draw a chart.** Reserve its place with `D.chart_placeholder` or
  `D.chart_with_comment`: header line, the brief for the author's Excel
  chart, comment column. Every slide with figures gets a source line,
  `D.source` (`references/drawing.md` §Charts, tables and sources).
- **Icons with `D.icon(doc, page, box, "name")`** — the pack's own (listed in
  `catalog.md`, *Icons*) first; `scripts/icons.py search <word>` finds more
  in the shipped Element Plus set (`"ep:name"`).
  Never build an icon from rectangles and lines or use a Unicode symbol.
- Never reopen the finished file to add more; everything happens in `fill`.

**5. Build — fast loop first, then the gate.** Give each run a timeout of at
least five minutes; a large deck's full gate (build + QA + render) takes
about ten minutes cold, so budget the whole run, not one command.

```bash
scripts/make_deck.py build.py out.pptx --build-only               # fix build errors (instant)
scripts/make_deck.py build.py out.pptx --template acme            # the gate: QA, schema, renders
```

Never pipe the output into `grep`, `head` or `tail`: its last lines say
whether the build worked, and a pipe reports exit 0 when it did not. Too long
to read? `... > build.log 2>&1; echo "exit $?"`, then read the end of the
log. A failed build leaves the previous deck in place, so a file on disk
proves nothing.

If a build dies part-way and the only error is `Binary URP bridge disposed
during call` (or `... already disposed`, a `DisposedException`), that is
soffice itself crashing — the bridge error is what Python sees on the *next*
call after `soffice.bin` died, so the traceback shows neither the slide nor
the real cause. The engine keeps soffice's stderr, and after a crash the
build prints `[impress] soffice exited abnormally (rc); its stderr was
captured in soffice-stderr-<profile>.log`. Read that file first: it holds
the actual signature (an abort, `free(): invalid pointer`, a stack). If it
is `free(): invalid pointer`, add `--pin-text-proxies` to the re-run (see
below). Otherwise simply re-run the same command — these crashes are probabilistic and one
retry often goes through. If it dies again *at a different slide*, build in
halves (`--parts`, never get the user to merge separate decks) instead of
re-running the whole deck. After three full-deck failures at different
slides, tell the user (a LibreOffice version pin or a split session may be
needed) before a fourth retry. Never `try/except: pass` over it in
`build.py`: the error is the information.

The gate writes `out-report.md` and `render/` beside the deck and exits 1 on
any finding. **Fix every finding in `build.py` and run the gate again** until
it is clean. Leftover template text, a missing title, text off the slide,
low contrast, top-heavy slides and drawn shapes that almost line up are
defects, not noise.

**Assume the first build has problems.** It always does. A deck is not done
after one gate run, however clean it reports: it is done after at least one
round of fix, rebuild and look again (steps 5 to 7).

Re-running the gate is cheap: it keeps each slide's findings in
`out-qa-cache.json` and re-measures only the slides whose content changed, so
after fixing two slides of sixty only those two are checked again. Fix several
findings per round rather than one. `--slides 4,7-9` narrows the checks
further, but skips the deck-wide ones (variety, leftover wording across
slides) — the last run before step 6 is always the full gate.

**6. The vision-agent looks at every slide — one slide per call.**

```bash
scripts/vision_brief.py slides out.pptx --language English
```

It lists the slides to send. For **each** listed slide N send the
vision-agent exactly
`Read the brief at <render>/slide-N.brief.md and follow it exactly.` and save
its whole reply, unedited, as `<render>/slide-N.answer.md`. Never put several
slides in one call and never write the brief yourself. If it answers
`CANNOT SEE IMAGE`, send the same brief again (up to three tries — it is often
transient); only then save it as the answer and tell the user.

The image in each brief has a **review code** printed under the slide, and an
answer counts only when it gives that code back. You cannot read it, so
**never write, repair or "restore" an answer yourself** — the gate rejects it.
Slides whose render did not change keep their answer across rebuilds; after a
fix, only the changed slides are listed again. Then:

```bash
scripts/vision_brief.py collect out.pptx       # fills the report's verdict table
```

**7. Every `fix` verdict is work.** Change `build.py`, re-run step 5, brief the
changed slides (step 6). Rewriting the verdict cell clears nothing: the gate
reads the vision-agent's saved answers. A finding a measurement proves wrong
(`probe.py` for geometry, `qa.py` for wording and colour, `wireframe.py` for
emptiness) goes in the allowlist below as kind `vision`, `content` the
question (`"Q7"`), the measurement as the reason — and **counts only once the
user has approved it**: show the user the entry and the measurement, and let
the user add `"approved_by"`. Never fill that field yourself. That text exists
in the file is no measurement that it shows: a title under a picture is in the
file and invisible. "Hollow", "empty", "no focal point" and "unfinished" are
not disproved by measurement; change the slide.

**8. Finish.**

```bash
scripts/make_deck.py out.pptx --require-verdicts
```

It only reads the report and the saved answers, and exits 1 while any slide
lacks an answer with its review code, an answer still holds a medium or high
finding, an overrule awaits the user's approval, or the report belongs to an
earlier build. The deck is done when this passes, and not before: never tell
the user it is built, done or finished until it has printed
`All N slide(s) seen by the vision-agent` for this build. Tell the user every
approved overrule it prints.
No vision-agent available? Run `scripts/wireframe.py render/`, tell the user
which slides nothing with eyes has seen, and ask the user to look at them: a
slide the user has checked is a `vision` entry with `"content": "seen"` and
the user's `"approved_by"`.

## When QA flags something intentional

Correct content can match the template (a phase name that *is* the content, a
divider's background graphic). Record it in `<deck>-allowlist.json` beside the
deck, never by distorting the content:

```json
{"acknowledgements": [
  {"slide": 16, "kind": "leftover-wording", "content": "Testing",
   "reason": "phase label on the lifecycle circle"}
]}
```

`kind` is `leftover-wording`, `stray-image`, `footer`, `balance`, `no-title`,
`variety`, `alignment` or `vision`; each entry must match a real finding and carry a
reason. Tell the user about every entry you add. A `vision` entry also needs
`"approved_by"`, which only the user writes.

## Scripts

| Script | What it does |
|---|---|
| `scripts/make_deck.py build.py deck.pptx --template acme` | Build, QA, schema-check, render; writes `<deck>-report.md` and `render/`. `--build-only` for the fast loop, `--require-verdicts` for the final gate |
| `scripts/vision_brief.py slides\|collect deck.pptx` | One tested vision-agent brief per render; `collect` reads the answers into the report |
| `scripts/make_deck.py build.py deck.pptx --parts 7-9 --template acme` | A `build_parts` deck: build and check only those slides, into `<deck>-parts-7_9.pptx` |
| `scripts/split_source.py source.md --out spec/` | A long source as one file per slide, plus `_shared.md` and `index.md` |
| `scripts/api_ref.py [draw]` | Every build function with its signature (`references/api.md`) |
| `scripts/probe.py deck.pptx [--slides 4,7-9]` | Every shape's index, position, size, colours and fonts |
| `scripts/dump.py deck.pptx` | The deck as markdown: text, tables, notes, empty slots |
| `scripts/thumbnail.py deck.pptx` | A contact sheet of all slides (for the vision-agent) |
| `scripts/render.py deck.pptx` | One PNG per slide |
| `scripts/qa.py deck.pptx --template acme` | The checks alone, without rendering |
| `scripts/icons.py search shield` | Icon names for `D.icon`: the pack's own, then the installed set |
| `scripts/wireframe.py render/slide-3.png` | A render as a text ink grid — the fallback without a vision-agent |
| `scripts/convert.py in out` | Between `.odp`/`.pptx`/`.otp`/`.potx`/`.ppt`/`.pdf` |
| `scripts/office/validate.py deck.pptx` | OOXML schema and relationship checks |

`make_deck.py` puts `scripts/` on the build script's path. Running a build
script yourself: `sys.path.insert(0, "<skill>/scripts")`.

| Module | Use it for |
|---|---|
| `office.deckbuild` | `build_deck` — the whole deck from a pack; `build_parts` — the same from one file per slide |
| `office.template` | `set_title`, `retext`, `set_text`, `set_char_color`, `find_shapes`, `set_footer`, `shrink_to_fit` — changing what a design already has |
| `office.draw` | `canvas`, `Box`, `text`, `para` (one paragraph styled apart), `cards`, `steps`, `kpis`, `statement`, `table`, `venn`, `icon`, `color`, `series`, `size_for` — drawing new things |
| `office.impress` | `impress_document`, `save_as`, `slides`, `shapes`, `mm()`, `pt()`, `slide_size` |

## References

| File | Read it when |
|---|---|
| the pack's `catalog.md`, `designs/design-N.md` | every house-style deck |
| `references/api.md` | the signature of any build function — instead of reading the source |
| `references/drawing.md` | drawing anything: `Box`, the helpers, composition |
| `references/big-decks.md` | over ~20 slides or a long source: parts, builder subagents |
| `references/language.md` | writing slide text: headlines, banned words |
| `references/storyline.md` | the deck is a subject, not supplied content |
| `references/accents.md` | choosing a look and a hero visual |
| `references/templates.md` | the user's own template or an existing deck to edit |
| `references/qa.md` | what a finding means, the vision workflow in detail |
| `references/authoring.md`, `references/design.md` | no template at all |
| `references/pptx-xml.md` | hand-editing OOXML |

## Environment

- apt only — never `pip` or `npm`. If a package is missing, name it and stop.
- Geometry is 1/100 mm: use `mm()`/`pt()`. Values read back (`area.y`,
  `probe.py` output) are already 1/100 mm — never wrap them in `mm()`.
- `from com.sun.star... import` works only after `import uno`.
- Each script starts its own headless LibreOffice; parallel runs are safe,
  but only a few run at once per machine (`PRESENTATION_MAX_OFFICE`); the
  rest print `[office] waiting` and wait. Set `PRESENTATION_TMPDIR` if `/tmp`
  is not writable. When soffice dies mid-run, its stderr is kept as
  `soffice-stderr-<profile>.log` beside your deck (the crash signature);
  clean runs leave none.
- **No network for LibreOffice or `build.py`.** Both run in their own network
  namespace (`unshare`, else `bwrap`), so a deck that links an image by URL
  cannot make anything fetch it, and a build script cannot send anything out.
  `build.py` runs with `python3 -P`: import only `office.*` and the standard
  library, not helper files beside it (a `build_parts` deck shares code
  through `slides/_common.py`, loaded by path). Without either tool a warning says so
  and decks that link remote resources are refused;
  `PRESENTATION_NETWORK=allow` is the explicit way past, for trusted decks.
- If LibreOffice dies mid-build with `free(): invalid pointer` or
  `Binary URP bridge disposed`, rerun with `make_deck.py --pin-text-proxies`
  (or `PRESENTATION_PIN_TEXT_PROXIES=1` for any script). Unpatched
  LibreOffice 25.2.x frees a text paragraph twice when a client drops it
  while re-reading the same shape; with the flag no text proxy is released
  until LibreOffice has exited. Costs a few MB per deck; off by default.
- On a slow or shared mount, build into a local directory and copy the
  finished deck across.
