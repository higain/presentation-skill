---
name: presentation-setup
description: "Teach the presentation skill a company template. Use when a deck should be built in a company's own design and no template pack is installed yet, when someone brings a .pptx/.potx/.odp/.otp master or corporate design to 'set up', 'onboard' or 'teach' the tooling, when D.color() raises because a deck has no palette, or when an existing pack needs updating after the template changed. Triggers on 'company template', 'corporate design', 'Firmentemplate', 'Folienmaster', 'brand template', 'set up the presentation skill', 'no pack installed'."
compatibility: "Needs the `presentation` skill installed beside this one, and its dependencies: LibreOffice Impress 25.2+ with python3-uno, python3-lxml, python3-defusedxml, python3-pil, poppler-utils. All from apt — never run pip or npm. Works offline; a vision model is optional and never required."
---

# Teaching the presentation skill your template

**The `presentation` skill ships with no template.** This one makes it a
**pack** — `templates/<name>/` in that skill, holding the master, its palette,
its type scale and a catalog of what each design is for. One pack, and every
deck after it comes out in the company's own design without anyone choosing a
colour.

Six stages. **Two of them are questions for a person and cannot be answered
any other way** — which colour is the accent, and what each design is for. The
other four are measurement, and none of them asks you anything.

```
template.pptx ─▶ 1 SURVEY ─▶ 2 PALETTE ─▶ 3 NOTES ─▶ 4 DESIGNS ─▶ 5 PROBE ─▶ pack
                  contact     counted,     ticked ▲   named ▲      three real
                  sheet       named ▲                              slides ▲
                                                          ▲ = ask the user
```

Twenty minutes for a twenty-design master. **Nothing leaves the machine**: the
extraction uses no model at all, and the one step that can call one is opt-in
and says so.

## Before you start

```bash
SKILL=../presentation        # or ~/.claude/skills/presentation
```

Ask for **the real master**, not a finished deck someone liked: a pattern
collection with one slide per design is what the catalog is built from. A
finished customer deck gives you twelve variations of one layout and a
customer's name on every slide.

Ask three things, and write the answers down:

1. **What is the pack called?** Lowercase, one word — `acme`. It goes in build
   scripts: `build_deck("acme", ...)`.
2. **What language are decks written in?** The template's own footer and
   labels are in *some* language, and a deck in another one has to override
   them. This is the fact stage 5 checks.
3. **Is there a local vision model?** If yes, stage 4 is faster. If no, stage 4
   still works — it is a worksheet instead of a model pass, and the result is
   usually better.

---

## Stage 1 — Survey

```bash
python3 $SKILL/scripts/thumbnail.py brand.pptx          # a contact sheet
python3 $SKILL/scripts/catalog_extract.py brand.pptx -o /tmp/catalog
```

Open the contact sheet and **count**. Does the design count match what you see?
Is it one slide per design, or does the file repeat a layout with different
content? If it repeats, say so now — the catalog will have twelve rows meaning
the same thing, and no later stage recovers from that.

Note the masters it prints. A template with three masters usually means
dividers and content are built differently, which stage 4 will confirm.

## Stage 2 — Palette and type scale  ·  **ask the user**

```bash
python3 scripts/count_style.py brand.pptx
```

Every colour it prints is one the template **already** fills a shape or sets
text in; every size is one it types at. It excludes what is provably not
design — LibreOffice's default fill, the yellow a designer highlights fill-in
fields with — and says so.

**Show the table and ask, in these words:**

> Which of these is your accent — the colour that marks the one thing on a
> slide that matters? And which is the alert colour, for something wrong?

Then map the rest: `ok` (a positive state), `muted` (the quiet fill behind
neutral content), `neutral` (a dark grey for captions and quiet type), and
`series` — the order colours should be used in on a chart, most important
first.

**Do not choose these yourself, and do not take the most-used colour as the
accent.** The most-used colour on a corporate master is nearly always a grey.
The accent is often used twice in the whole file.

For the type scale, map the sizes to `display, title, subtitle, body, label,
caption, footnote`. The counter shortlists seven and keeps the two largest
whatever their weight — a display size is used once a deck and would otherwise
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
template's slides do not contain and names the nearest one they do — that check
exists because the first pack built with this script took three of its colours
from an example command written for a different template.

It warns rather than fails on the rest: an accent under 4.5:1 on white is
headline-only, under 3:1 it can never be a word at all. **Both are normal for a brand colour**
and knowing which one you have decides how every deck on this template looks.

## Stage 3 — Authoring notes  ·  **ask the user**

```bash
python3 scripts/find_notes.py brand.pptx
```

A pattern collection talks to whoever builds from it — `EXAMPLE`, `Customer
logo here`, a paragraph of instructions on slide four. They are ordinary
shapes, nothing flags them, and they ship.

**Show the candidates and have them ticked.** Getting this wrong in one
direction leaves instructions in a customer deck; in the other it silently
deletes a column header. A person reading twelve short strings settles it in a
minute.

```bash
python3 scripts/make_pack.py --name acme --template brand.pptx \
    --note 'EXAMPLE' --note 'Customer logo here' \
    --note-prefix 'Instructions: 1.'
```

Re-running `make_pack.py` updates the profile and leaves the catalog alone.

## Stage 4 — What each design is for  ·  **ask the user**

Move the extraction into the pack, then take one of two paths.

```bash
mv /tmp/catalog $SKILL/templates/acme/catalog
```

**Without a vision model — the default, and usually better:**

```bash
python3 scripts/name_designs.py $SKILL/templates/acme/catalog
# fill in designs.md beside the renders, then:
python3 scripts/name_designs.py $SKILL/templates/acme/catalog --import
```

The worksheet arrives with every measured field already filled — slot counts,
capacities, retext keys, title budgets — and three blanks per design: what it
is for, when to use it, when not to. Twenty designs is about five minutes with
the contact sheet open. A design left blank still builds; the decision table
just never picks it for the right reason, and the import says which.

**With a local vision model**, the `presentation` skill's own pipeline does the
same job from the renders:

```bash
python3 $SKILL/scripts/catalog_gemma.py $SKILL/templates/acme/catalog \
    --base-url http://localhost:11434/v1 --model gemma3:27b
```

**It must be a vision model.** A text-only one answers confidently about images
it never received, and every design comes back "bulleted content". Check the
`--dry-run` output shows an image part.

Then assemble:

```bash
python3 $SKILL/scripts/catalog_build.py $SKILL/templates/acme/catalog \
    -o $SKILL/templates/acme/catalog.md
```

## Stage 5 — Probe  ·  **ask the user**

```bash
python3 scripts/probe_deck.py acme
python3 $SKILL/scripts/render.py probe.pptx
```

Three real slides: a divider, a content slide retexted through the catalog's
own keys, and a drawn slide in the pack's own colours. **Open the three PNGs
and ask three questions in order:**

1. Does the divider title fit on one line, clear of the master's decoration?
2. Is any word on slide 2 still the template's own?
3. Are the colours on slide 3 the ones the brand actually uses?

A *no* to any of them is a five-minute fix now: the title budget is in the
catalog, leftover wording means a missing retext key or a missed authoring
note, and a wrong colour means stage 2's roles were mapped wrong. Re-run
`make_pack.py` with corrected roles; it does not touch the catalog.

**If you cannot see images, say so and stop here.** Do not report a pack as
finished on the strength of a clean QA run — QA cannot see a title crossing a
decoration, and that is precisely what this stage is for.

## Afterwards

The pack is live. `build_deck("acme", ...)` and `--template acme` work from
anywhere, and a deck requested with no template named is built on it.

Put the review notes nobody can measure — a design that is not what it looks
like, a slot that is empty in the template, a divider budget that is a hard
limit — into `templates/acme/overrides.md` under a `## header` or
`## design-N` anchor, and rebuild. **Never edit `catalog.md`**: it is
generated, and the next rebuild overwrites it. `catalog_build.py --check` fails
when it has drifted, which is how that rule stays true.

When the template changes, re-run stages 1 and 4 and read the diff. The retext
keys are what changed.

## Rules

- **Never invent a colour, a size or a design number.** Everything this skill
  writes is counted off the template or answered by a person. Four separate
  runs against a corporate template drew decks in a hex the company does not
  own, and no check in the pipeline could see it.
- **Nothing leaves the machine by default.** Extraction uses no model.
  `catalog_gemma.py` calls the endpoint you give it — point it at localhost.
  A hosted endpoint means slide images going to a third party; say so out loud
  before doing it, and never for a confidential template.
- **apt only.** Never run `pip` or `npm`.
- **The template file is brand-internal.** `templates/*` is gitignored in the
  presentation skill. Keep it that way.
