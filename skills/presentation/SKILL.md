---
name: maxi-presentation
description: "Create, edit, read and QA presentations — .odp and .pptx alike, plus .otp/.potx templates and legacy .ppt. Use whenever a slide deck is involved: building a deck or pitch, filling in a template, editing or restyling an existing one, extracting text or speaker notes, converting between formats, or rendering slides to images for visual review. Triggers on \"PowerPoint\", \"Impress\", \"deck\", \"slides\", \"presentation\", \"Präsentation\", \"Folien\", or any .odp/.otp/.pptx/.potx/.ppt filename. The user's word for a file says nothing about its extension — a file someone calls \"the PowerPoint\" is often an .odp — so check the actual file, and save a deck back in the format it arrived in."
compatibility: "Needs LibreOffice Impress 25.2+ with python3-uno, plus python3-lxml, python3-defusedxml, python3-pil and poppler-utils. All dependencies come from apt — never run pip or npm. Works offline; no network access required."
license: Proprietary. LICENSE.txt has complete terms
---

# Presentations: .odp and .pptx

**The engine is LibreOffice Impress, driven over UNO.** It opens the real
document and saves it in its own format, so an `.odp` stays ODF and a `.pptx`
stays OOXML. Nothing here rewrites a file format by hand.

| Task | How |
|---|---|
| **Read** a deck | `scripts/dump.py deck.odp` |
| **Survey** a template's layouts | `scripts/thumbnail.py template.odp` |
| **Inspect** shapes to edit them | `scripts/probe.py template.odp` |
| **Build in the house style** | the recipe below — this is the common job |
| **Adapt** a user's own template | `references/templates.md` |
| **Create** a deck from nothing | `references/authoring.md` |
| **Edit** an existing deck | `impress_document(path)` — saves in place, same format |
| **Convert** formats | `scripts/convert.py in.odp out.pptx` |
| **Check and render** a deck | `scripts/make_deck.py build.py deck.pptx --template brand.pptx` |

**Three rules that override anything below:**

1. **Keep the format the deck arrived in.** Editing an `.odp` produces an
   `.odp`. Convert only when the user asks for the other format.
2. **A new deck defaults to `.pptx`** unless the user asks for `.odp` or names
   an output path that says otherwise.
3. **A deck with no template of the user's own is built from our house
   template** — see [Company template](#company-template-default) — not from a
   blank document and not from an invented look. Their template always wins if
   they supply one.

## Step 0 — which job is this?

**Answer this before running a single command.** The recipe below builds slides.
It does not decide what the slides should say, and there is no check anywhere in
this skill that notices when nobody did.

| What you were handed | Start at |
|---|---|
| An outline, a document, a deck to edit, or the slide text itself | **The recipe**, below |
| A **subject** — "a deck about X", a sentence or a paragraph of intent | **`references/storyline.md`**, first, and do not open the recipe until stage 3 of it is approved |

A one-line request for a deck on a topic is the second row, not the first. So is
a paragraph. So is a brief that names an audience and a goal but no content:
the giveaway is that **you would have to invent what the slides say**, and that
is the work `references/storyline.md` governs — six intake questions, one
scaffold, a ghost deck the user approves, an accent pack the user picks, and
only then the recipe.

**Do not substitute your own outline for that file.** Writing a plan of your own
called `storyline.md`, or picking designs while sketching the sections, skips
the two user gates and produces a deck that lists what is known about a topic
instead of arguing something. That failure is invisible to `qa.py`, to
`validate.py` and to the renders — a deck can pass every check in this skill and
still have never made a point.

## The recipe

**This is the commonest job — a `.pptx` on the house template — and this is the
whole of it.** Six steps in order, and **only once step 0 sent you here**.

```bash
# 1. See what the template offers.
scripts/thumbnail.py assets/company-template.pptx
scripts/probe.py assets/company-template.pptx --slides 8,13
```

**2. Pick a design per section** from `references/company-template.md` — a
decision table, top to bottom, first match wins. Its numbers are the slide
numbers you pass to `build_deck`, and each entry lists the **retext keys** that
design actually carries. Content the table has no design for — a diagram, a
chart, a comparison you want to draw — goes on **design 8 stripped to a
canvas**; see `references/drawing.md`.

**3. Write one build script.** One `fill` per page, retexting *and* drawing.
Never reopen the finished file to add more: that is a second `.pptx` save, it
grows the masters, and it is how shapes end up layered over template text
nobody replaced.

**Branch on `slot`, the position in your deck — not on `design`, and never on
both.** `if design == 8 and slot == 1` is false whenever slot 1 holds a
different design, and a branch that never fires leaves that slide exactly as the
template drew it: German body text under whatever you drew on top. **Every slot
needs a branch**; end `fill` with an `else` that raises if you want to be sure.

```python
from office.deckbuild import build_deck
from office.template import retext
from office.impress import mm
from office import draw as D

def fill(doc, page, slot, design):
    if slot == 0:
        retext(page, {"Potentielle Navigatoren": "How we got here"})
    if slot == 1:                                   # a drawn slide
        area = D.canvas(page, keep=(0, 1, 2, 3))    # title, topic, source
        body, foot = area.cut_bottom(mm(14))        # a PAIR -- unpack it
        D.venn(doc, page, body,                     # sizes itself to the box
               D.series(3),                         # the template's own colours
               labels=["Private data", "Untrusted content", "External comms"],
               centre=D.color("alert"), centre_text="Exfiltration")
        D.text(doc, page, foot, "Any two is safe by design",
               role="caption", align="center")      # colour picked for the bg

build_deck("assets/company-template.pptx", "out.pptx", [7, 8, 13], fill=fill,
           footer="Lethal trifecta · Board briefing · 2026")
```

**Vary the designs, and set the footer.** Both are one argument of that last
call, and four separate runs got both wrong: every one built `[8] * N`, drawing
every slide from scratch on a template whose twenty other designs it never
touched, and every one left the template's German `Gegenstand | Titel | Ort |
Datum` along the bottom of an English deck. The footer is a page property, so
no amount of `retext` reaches it; `footer=None` hides it instead.

```bash
# 4-6. Build, check, render — one command, and it prints the PNG paths.
scripts/make_deck.py build.py out.pptx --template assets/company-template.pptx
```

**Give this command a timeout of at least five minutes.** It opens LibreOffice
several times and a fifteen-slide deck runs to a minute or two. A shell that
gives up at the default kills it *after* the build and *before* it prints where
the renders are — which is how a deck gets called finished having never been
looked at. It writes `<deck>-report.md` and `render/` beside the deck either
way, and names both in its first line, so a run that did get cut off is
recovered by reading those rather than by building again.

**5. Fix every finding it reports, then run it again.** Exit 1 means findings,
and they are also left on disk in `<deck>-report.md` beside the deck. Leftover
template wording, a stray template picture, text running off the slide and
black-on-navy text are defects, not noise. **Editing the build script is not
fixing them** — nothing has changed until the build has re-run and the report
comes back clean.

**6. Open the PNGs, look at them, and write a verdict per slide.** The report
ends in a table with one row per slide and the verdict column blank. **The deck
is not finished while a cell is empty** — that is the visual half of the check,
and it has the same standing as the automated half. If you cannot see images,
run `scripts/wireframe.py render/` — it prints each slide as an ink grid, which
shows emptiness, banding and content that is not centred — say plainly that you
worked from that, and name the slides you could not verify.

Where it prints **slides to look at**, start there: those are the ones measured
as painting almost nothing, and they are either deliberate statement slides or
slides you left unfinished. Deciding which is a thing only eyes can do.

**Every slide needs something to look at.** A deck of twelve bullet lists is a
failed deck even when every check passes — reach for a diagram, a comparison, a
timeline, a table of real numbers. See Composition in `references/drawing.md`.

## Scripts

**Every path on this page is relative to this skill's own directory, not to
where you are working.** Find it once — the directory holding the `SKILL.md`
you are reading — and prefix it, or `cd` there. `assets/` and `references/` are
inside it. (`build_deck` resolves a template against it for you, so
`build_deck("assets/company-template.pptx", ...)` works from anywhere.)

**Run each script directly** — every one carries a shebang and the executable
bit, so `scripts/dump.py deck.odp` is the whole invocation. `python3
scripts/dump.py deck.odp` does the same thing if the executable bit was lost.

| Script | What it does |
|---|---|
| `scripts/make_deck.py [build.py] deck.pptx --template brand.pptx` | **Run this last.** Runs the build script, QAs, schema-checks, renders every slide, and writes `render/` and `<deck>-report.md` beside the deck. Exit 1 on any finding |
| `scripts/probe.py deck.odp [--slides 4,7-9]` | Every shape's **index**, position, size, fill, line and character formatting, plus a per-slide reading-order block. `dump.py` says what a slide *says*; this says what it *is*. Run it before editing any template slide |
| `scripts/thumbnail.py deck.odp [prefix]` | One labelled contact sheet of all slides, for picking template layouts |
| `scripts/dump.py deck.odp [-o out.md]` | Deck to markdown: per slide the layout, master, text, tables, speaker notes, empty placeholder slots and a hidden flag |
| `scripts/render.py deck.odp [-o dir] [--dpi N] [--slides 3,7-9]` | One PNG per slide; prints their absolute paths. **This is visual QA** |
| `scripts/qa.py deck.pptx [--template brand.pptx]` | The checks alone, without the renders: overflow, mid-word breaks, off-slide shapes, leftover template wording and pictures, placeholders, markdown, contrast, occlusion — plus `validate.py` on a `.pptx` |
| `scripts/review_renders.py render/ --model <vision model>` | **The only check with eyes.** Sends each render to a vision model and reports what it saw. Advisory: always exit 0. `make_deck.py` runs it for you when `$PRESENTATION_REVIEW_MODEL` is set |
| `scripts/wireframe.py render/slide-03.png` | A render as text: an ink grid, the centroid, the widest blank band, the margins. **What to run when you cannot see images** — it shows emptiness, banding and things that are not centred without a model. `make_deck.py` prints one for every slide it flags |
| `scripts/convert.py in.odp out.pptx` | Convert between `.odp`/`.pptx`/`.otp`/`.potx`/`.ppt`/`.pdf` |
| `scripts/office/validate.py deck.pptx [--original src.pptx]` | Schema, relationship, content-type, chart and slide checks. **`.pptx` only** |
| `scripts/catalog_extract.py` · `scripts/catalog_build.py` | **Build-time only, not per deck:** teach the skill a new company template. See `README.md` |
| `tests/test_geometry.py` · `tests/roundtrip_report.py` | Fast geometry checks; and the fidelity re-proof after a LibreOffice upgrade |

The `office` modules a build script imports:

| Module | What it gives you |
|---|---|
| `office.deckbuild` | `build_deck` — a whole deck from a template, in any order |
| `office.draw` | `rect`, `ellipse`, `line`, `arrow`, `polygon`, `chevron`, `text`, `chip`, `bullets`, `table`, `venn`, `canvas`, `Box`, `page_box`, `restyle`, `contrast_ratio`, `readable_on`, `fit_size`, `fit_lines` — **putting new things on a slide**, and `Box` to lay it out |
| `office.draw` — composition | `cards`, `steps`, `kpis`, `statement` — the four layouts every build draws by hand. One call each, with the padding, the equal heights, the arrows and the line breaks already right. See Composition in `references/drawing.md` |
| `office.draw` — house style | `color`, `series`, `palette`, `size_for`, and `role=` on `text` — **the template's own colours and sizes, by name.** Four runs invented hex the template does not contain and picked a footnote size per slide; these are counted off the asset. `references/company-template.md` prints both tables |
| `office.template` | `retext`, `set_text`, `find_shapes`, `shape_by_text`, `clone_slide`, `strip_annotations`, `fit_pill`, `shrink_to_fit`, `scale_bars`, `text_overflow`, `overflowing`, `colliding`, `set_footer` — editing what a template slide already has |
| `office.impress` | open, save, slides, shapes, placeholders, notes, `slide_size`, `mm()`/`pt()` |
| `office.rasterize` | `rasterize` — deck to images, with the hidden-slide policy as an argument |

**Importing `office.*`:** `make_deck.py` puts this skill's `scripts/` directory
on the build script's path for you, so `from office.draw import ...` just works
under it. Running a build script yourself, add the path — the directory that
*contains* `office/`, never `office/` itself:

```python
import sys; sys.path.insert(0, "/absolute/path/to/presentation/scripts")
```

## References — read the one the job calls for

| File | When |
|---|---|
| `references/company-template.md` | **Always, for a house-style deck** — which design for which content, the retext keys, the capacities |
| `references/drawing.md` | Drawing anything: `Box` layout, every primitive, the canvas pattern, composition |
| `references/templates.md` | The user brought their own template: probe, `clone_slide`, `retext`, charts, pills, long titles |
| `references/qa.md` | What each check means, visual QA, `validate.py`, what a round trip preserves |
| `references/authoring.md` | No template at all: new documents, placeholders, typography, environment |
| `references/design.md` | Inventing a look — palette and motif, for a deck with no template |
| `references/language.md` | **Writing the words** — action titles, the banned list, the four rewrite tests |
| `references/storyline.md` | **A deck from a brief, not from supplied text** — intake, scaffold, ghost deck, the two user gates |
| `references/accents.md` | The look: accent packs, the hero-visual catalogue that `plan.md`'s `visual:` field names |
| `references/pptx-xml.md` | Hand-editing OOXML |

**Writing slide text rather than placing text the user supplied?** Read
`references/language.md` first. Every title is a sentence stating the
conclusion, not a label for the topic — it is the difference between a deck that
argues and a deck that lists, and `qa.py` checks it.

**Asked for a deck on a subject, rather than handed the content for one?** That
is step 0 above, and `references/storyline.md` runs before anything on this
page. If you have already written a build script, you are past the point where
reading it helps.

## Company template (default)

This skill ships our house template as the default. **When a deck is requested
and the user names no template of their own, build from
`assets/company-template.pptx`.**

1. **Choose designs from the catalog.** `references/company-template.md` maps
   content to **design numbers** — the slide numbers in the asset — as a
   decision table applied top to bottom, first match wins. Read it first, pick a
   design per section, and vary them per its variety rule. Do not eyeball the
   thumbnail and guess.

   **"Design 8" always means the catalog's design 8**, which is the number you
   put in `build_deck`'s list. It is never an Impress autolayout number: the
   autolayout `probe.py` prints is `20` for every slide in this template and
   identifies nothing.
2. **Respect the catalog's capacities.** Each entry states how many bullets a
   body holds, how many columns, how many images. When content exceeds that,
   split the slide rather than overflow it.

   **Take `retext` keys from the entry's *Retext keys* line, never from
   memory.** They differ per design, and reusing design 8's
   `"Potentielle Navigatoren"` on design 5 or 12 is the single commonest way a
   build script dies. A slot that is **absent from a design's key list** — or
   listed under *Empty slots* — holds no text at all, and `retext` cannot match
   an empty shape: write it with `set_text(shapes(page)[i], ...)` or draw your
   own.
3. **Anything the table has no design for goes on design 8 stripped to a
   canvas.** Do not force it into the nearest design's bullet list, and do not
   start from a blank slide.
4. **Build with `build_deck`** against the asset, passing the design numbers in
   final order and doing all retexting *and* drawing in its `fill` callback. It
   duplicates rather than rebuilds, so this template's bullets, inline colour
   and think-cell objects survive; `clone_slide` loses all three here.
5. **Check and look:**
   `scripts/make_deck.py build.py out.pptx --template assets/company-template.pptx`,
   then open every PNG it prints.

**Maintaining the default** — extracting the catalog from a new company
template — is `README.md`. It runs once, on internal infra; never send company
slides to a hosted model.

## QA (required)

**A deck is not finished until `make_deck.py` has run clean and every slide has
a verdict written in the report's table.** Every deck that shipped broken had
skipped one of the two — and the ones that skipped the second had all "looked
at the renders" first.
`make_deck.py` reports text that overruns its box or the slide, blocks of text
printed over each other, leftover template wording and pictures, markdown,
unreadable colour contrast and occluded shapes, then schema-checks the `.pptx`.
It writes `<deck>-report.md` and `render/` beside the deck.

It also prints **slides to look at**: the ones painting under a quarter of
their own area, or leaving a third of their height in one bare band. Those are
measurements, not findings — they do not fail the run, because a one-sentence
statement slide measures the same as a slide nobody finished. Open them first.

**If `$PRESENTATION_REVIEW_MODEL` names a vision model, `make_deck.py` also
sends every render to it and prints what it saw.** That is advisory too and
never fails the build. It does not excuse you from looking — but if you cannot
see images, it is the difference between a deck nothing has looked at and a
deck one thing has.

**A visual finding is dismissed only by a measurement you show.** Reasoning
about why the reviewer must be wrong is not a measurement, and it is how a deck
shipped with a German footer on all ten slides: the run searched the shapes for
the string, found nothing, and wrote the finding off as a hallucination — the
footer is a page property and no shape search could ever have reached it.
`probe.py` settles geometry, `qa.py` settles wording and colour,
`wireframe.py` settles emptiness and placement. If none of them can settle it,
the finding stands, and fixing it is cheaper than arguing with it.

What it **cannot** see: alignment, crowding, a diagram off-centre, a chart
whose bars are not proportional, a slide still in the template's language. That
is what the renders are for. Details, and what each finding means:
`references/qa.md`.

## Notes on this environment

- **apt only.** Never run `pip` or `npm`. Everything needed is installed; if
  something is missing, say so and stop.
- **Geometry is 1/100 mm.** Use `mm()`/`pt()`; read `slide_size(doc)` rather
  than assuming — the house template is 32001 × 18000, not LibreOffice's
  28000 × 15750.
- **`from com.sun.star... import ...` only works after `import uno`.**
- The locale is German (`de_DE.UTF-8`); German content is normal.
- Each script starts its own headless LibreOffice with a private profile and
  port, so parallel runs never collide. If `/tmp` is not writable, set
  `PRESENTATION_TMPDIR` — see `references/authoring.md`.
