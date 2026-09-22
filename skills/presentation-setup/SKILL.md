---
name: presentation-setup
description: "Teach the presentation skill a company template. Use when a deck should be built in a company's own design and no template pack is installed yet, when someone brings a .pptx/.potx/.odp/.otp master or corporate design to 'set up', 'onboard' or 'teach' the tooling, when D.color() raises because a deck has no palette, or when an existing pack needs updating after the template changed. Triggers on 'company template', 'corporate design', 'Firmentemplate', 'Folienmaster', 'brand template', 'set up the presentation skill', 'no pack installed'."
license: Apache-2.0
compatibility: "Needs the `presentation` skill installed beside this one (same skills/ directory or ~/.claude/skills/), and its dependencies: LibreOffice Impress 25.2+ with python3-uno, python3-lxml, python3-defusedxml, python3-pil, poppler-utils, Python 3.11+. All from apt, never pip or npm. Works offline; a vision model is optional and never required."
---

# Teaching the presentation skill your template

**The `presentation` skill ships with no template.** This skill makes it a
**pack**: `templates/<name>/` inside that skill, holding the master, its
palette, its type scale and a catalog of what each design is for. One pack,
and every deck after it comes out in the company's own design without anyone
choosing a colour.

Five stages. **Four of them end in a question only a person can answer**:
which colour is the accent, which strings are the template talking to its
builder, what each design is for, and whether three built slides look right.
The rest is measurement, and none of it asks you anything.

```
template.pptx ─▶ 1 SURVEY ─▶ 2 PALETTE ─▶ 3 NOTES ─▶ 4 DESIGNS ─▶ 5 PROBE ─▶ pack
                  contact     counted,     ticked ▲   named ▲      three real
                  sheet       named ▲                              slides ▲
                                                          ▲ = ask the user
```

About twenty minutes for a twenty-design master. **Nothing leaves the
machine**: the extraction uses no model at all, and the one step that can
call one is opt-in and says so.

## Before you start

Work from this skill's own directory, so the commands below are exactly as
written. `$SKILL` is the presentation skill; `presentation-setup` finds it
beside itself or in `~/.claude/skills/`, and every script here stops with a
message if it cannot.

```bash
cd <path to>/presentation-setup
SKILL=../presentation          # or ~/.claude/skills/presentation
```

Ask for **the real master**, not a finished deck someone liked: a pattern
collection with one slide per design is what the catalog is built from. A
finished customer deck gives you twelve variations of one layout and a
customer's name on every slide.

Ask three things, and write the answers down:

1. **What is the pack called?** Lowercase, one word: `acme`. It goes in build
   scripts as `build_deck("acme", ...)` and becomes the directory name.
2. **What language are decks written in?** The template's own footer and
   labels are in *some* language, and a deck in another one has to override
   them. This is the fact stage 5 checks.
3. **Is there a local vision model?** If yes, stage 4 can use it. If no, stage
   4 still works: it is a worksheet instead of a model pass, and the result is
   usually better.

---

## Stage 1 — Survey

```bash
python3 $SKILL/scripts/thumbnail.py brand.pptx      # writes brand.jpg, a contact sheet
python3 scripts/count_style.py brand.pptx           # the colours and sizes it uses
```

Open the contact sheet and **count**. Does the design count match what you
see? Is it one slide per design, or does the file repeat a layout with
different content? If it repeats, say so now: the catalog will have twelve
rows meaning the same thing, and no later stage recovers from that.

Note the masters the sheet labels. A template with three masters usually
means dividers and content are built differently, which stage 4 will confirm.

## Stage 2 — Palette and type scale  ·  **ask the user**

`count_style.py` printed a table. Every colour in it is one the template
**already** fills a shape or sets text in; every size is one it types at. It
excludes what is provably not design (LibreOffice's default fill, the yellow
a designer highlights fill-in fields with) and says so.

**Show the table and ask, in these words:**

> Which of these is your accent, the colour that marks the one thing on a
> slide that matters? And which is the alert colour, for something wrong?

Then map the rest: `ok` (a positive state), `muted` (the quiet fill behind
neutral content), `neutral` (a dark grey for captions and quiet type), and
`series`, the order colours should be used in on a chart, most important
first.

**Do not choose these yourself, and do not take the most-used colour as the
accent.** The most-used colour on a corporate master is nearly always a grey.
The accent is often used twice in the whole file.

For the type scale, map the sizes to `display, title, subtitle, body, label,
caption, footnote`. The counter shortlists seven and keeps the two largest
whatever their weight: a display size is used once a deck and would otherwise
never make the list.

Then write it:

```bash
python3 scripts/make_pack.py --name acme --template brand.pptx \
    --colour accent=#RRGGBB --colour alert=#RRGGBB --colour ok=#RRGGBB \
    --colour muted=#RRGGBB --colour neutral=#RRGGBB \
    --series '#RRGGBB,#RRGGBB,#RRGGBB' \
    --size display=60 --size title=16 --size subtitle=14 --size body=14 \
    --size label=12 --size caption=10 --size footnote=9 \
    --default
```

**Every `#RRGGBB` above is a placeholder.** `make_pack.py` refuses a colour the
template's slides do not contain and names the nearest one they do. When it
refuses, do not retype the colour from memory: show the user the refusal and
the nearest colour it names, and ask which they meant. If the colour really
is the brand's and lives only in the master, `--allow-unseen` accepts it.

It warns rather than fails on the rest: an accent under 4.5:1 on white is
headline-only, under 3:1 it can never be a word at all. **Both are normal for
a brand colour**, and knowing which one you have decides how every deck on
this template looks.

This creates `$SKILL/templates/acme/` with the template copied in and a
`profile.json`. `--default` makes it the pack a deck is built on when nobody
names one.

## Stage 3 — Authoring notes  ·  **ask the user**

```bash
python3 scripts/find_notes.py brand.pptx
```

A pattern collection talks to whoever builds from it: `EXAMPLE`, `Customer
logo here`, a paragraph of instructions on slide four. They are ordinary
shapes, nothing flags them, and they ship.

**Show the candidates and have them ticked.** Getting this wrong in one
direction leaves instructions in a customer deck; in the other it silently
deletes a column header. A person reading twelve short strings settles it in
a minute.

```bash
python3 scripts/make_pack.py --name acme --template brand.pptx \
    --note 'EXAMPLE' --note 'Customer logo here' \
    --note-prefix 'Instructions: 1.'
```

Re-running `make_pack.py` **merges**: what the command line does not mention
(here the palette and the type scale) is kept from the profile already on
disk, and the catalog is never touched.

## Stage 4 — What each design is for  ·  **ask the user**

Extract the measured half straight into the pack:

```bash
python3 $SKILL/scripts/catalog_extract.py brand.pptx -o $SKILL/templates/acme/catalog
```

Thirty seconds, no model. One folder per design with its render and a
`structure.json` of everything measurable: retext keys, empty slots, title
budgets, the canvas hint. Then take one of two paths for the judged half.

**Without a vision model (the default, and usually better):**

```bash
python3 scripts/name_designs.py $SKILL/templates/acme/catalog
```

That writes `designs.md` beside the renders, with every measured field
already filled and three blanks per design: `purpose`, `use_when`,
`avoid_when`. **Fill them with the user, not for them.** Open the contact
sheet, go design by design, propose one sentence for each blank from what
you see, and let the user correct it. Twenty designs is about five minutes.
Then:

```bash
python3 scripts/name_designs.py $SKILL/templates/acme/catalog --import
```

A design left blank still builds; the decision table just never picks it for
the right reason, and the import says which ones.

**With a local vision model**, the `presentation` skill's own pipeline does
the same job from the renders:

```bash
python3 $SKILL/scripts/catalog_gemma.py $SKILL/templates/acme/catalog \
    --base-url http://localhost:11434/v1 --model gemma3:27b
python3 $SKILL/scripts/catalog_gemma.py $SKILL/templates/acme/catalog --aggregate \
    --base-url http://localhost:11434/v1 --model gemma3:27b
```

**It must be a vision model.** A text-only one answers confidently about
images it never received, and every design comes back "bulleted content".
Check that `--dry-run` shows an image part in the request.

Either way, assemble:

```bash
python3 $SKILL/scripts/catalog_build.py $SKILL/templates/acme/catalog
```

Inside a pack that needs no flags: it writes the pack's `catalog.md` and
splices in the pack's `overrides.md` if there is one.

## Stage 5 — Probe  ·  **ask the user**

```bash
python3 scripts/probe_deck.py acme            # writes probe.pptx here
python3 $SKILL/scripts/render.py probe.pptx   # writes render/slide-1.png ...
```

Three real slides: a divider, a content slide retexted through the catalog's
own keys, and a drawn slide on the pack's canvas design in the pack's own
colours. **Open the three PNGs and ask, in order:**

1. Does the divider title fit on one line, clear of the master's decoration?
2. Is any word on slide 2 still the template's own?
3. Are the colours on slide 3 the ones the brand actually uses?
4. On slide 3, is anything of the template's own furniture showing through
   the drawing, or is the drawing sitting where the design's own title was?

A *no* to 1, 2 or 3, or a *yes* to 4, is a five-minute fix now:

- **Title wraps:** the budget is in the catalog; note it in `overrides.md`
  under `## design-N` as a hard limit.
- **Leftover wording:** a missed authoring note. Re-run stage 3 with it added.
- **Wrong colours:** stage 2's roles were mapped wrong. Re-run `make_pack.py`
  with corrected roles; it does not touch the catalog.
- **Wrong canvas:** the catalog picked the design with the largest free area,
  and on this template that is the wrong one. Fill the `## Canvas` block at
  the top of `designs.md` (the design number, and the shape indices to keep,
  as `probe.py` numbers them), re-run `name_designs.py --import`, and rebuild
  with `catalog_build.py`. The hand-set canvas survives every rebuild.

A finding on a shape the template itself drew past the slide edge (some
generated masters do this) is the template's, not the pack's; note it in
`overrides.md` and move on.

Then run the probe again and look again.

**If you cannot see images, say so and stop here.** Do not report a pack as
finished on the strength of a clean QA run. QA cannot see a title crossing a
decoration, and that is precisely what this stage is for.

## Afterwards

The pack is live. `build_deck("acme", ...)` and `--template acme` work from
anywhere, and a deck requested with no template named is built on it.

Put the review notes nobody can measure (a design that is not what it looks
like, a slot that is empty in the template, a divider budget that is a hard
limit) into `templates/acme/overrides.md` under a `## header` or
`## design-N` anchor, and rebuild:

```bash
python3 $SKILL/scripts/catalog_build.py $SKILL/templates/acme/catalog
```

**Never edit `catalog.md`**: it is generated, and the next rebuild overwrites
it. `catalog_build.py --check` fails when it has drifted, which is how that
rule stays true.

When the template changes, re-run the extraction into the same `catalog/`
directory (it keeps `entries/` and `decision_table.json`), rebuild, and read
the diff. The retext keys are what changed.

## Rules

- **Never invent a colour, a size or a design number.** Everything this skill
  writes is counted off the template or answered by a person. Four separate
  runs against one corporate template drew decks in a hex the company does
  not own, and no check in the pipeline could see it.
- **Nothing leaves the machine by default.** Extraction uses no model.
  `catalog_gemma.py` calls the endpoint you give it; point it at localhost.
  A hosted endpoint means slide images going to a third party: say so out
  loud before doing it, and never for a confidential template.
- **apt only.** Never run `pip` or `npm`.
- **The template file is brand-internal.** `templates/*` is gitignored in the
  presentation skill. Keep it that way.
