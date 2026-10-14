---
name: presentation-setup
description: "Teach the presentation skill a company template: turn a .pptx/.potx/.odp/.otp master into a template pack. Use when someone brings a master or corporate design to 'set up', 'onboard', 'install' or 'teach' the tooling, when no template pack is installed yet, when D.color() raises because a deck has no palette, or when an existing pack must be rebuilt after the template changed. Triggers on 'set up the presentation skill', 'Firmentemplate einrichten', 'Folienmaster einlesen', 'install a brand template', 'no pack installed'. Not for building or editing a deck on a template that is already installed -- that is the presentation skill alone, even when the request says 'im Firmentemplate' or 'in our corporate design'."
license: Apache-2.0
compatibility: "Needs the `presentation` skill installed beside this one (same skills/ directory or ~/.claude/skills/), and its dependencies: LibreOffice Impress 25.2+ with python3-uno, python3-lxml, python3-defusedxml, python3-pil, poppler-utils, Python 3.11+. All from apt, never pip or npm. Works offline. Judging each design's render is done by the vision-agent (or a person) and never sends a render anywhere."
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
machine**: the extraction uses no model at all, and the one step that needs
eyes — judging what each design is for — is done by the vision-agent, which
runs on your own infrastructure.

## Before you start

**Work in the user's folder, where the template file is** — never inside a
skill directory. Contact sheets, the probe deck and its renders are written to
the working directory, and they are the company's slides: they must not land in
a skill that may be committed or shared. In the commands below, `$SETUP` and
`$SKILL` stand for two directories:

```bash
$SETUP = <path to>/presentation-setup     # this skill
$SKILL = <path to>/presentation           # the presentation skill, usually beside it
```

Write both out in full in every command, as you were given them, rather than
setting shell variables: each command may run in a fresh shell, and permission
rules match the command as typed. A `$SKILL/...` or a relative path can stop the
run at a prompt nobody sees.

Every script here finds the presentation skill by itself and stops with a
message if it cannot.

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
3. **Is the vision-agent available?** If yes, stage 4 delegates judging each
   design's render to it. If no, stage 4 still works: fill the worksheet
   yourself with the user, and the result is usually better.

---

## Stage 1 — Survey

```bash
python3 $SKILL/scripts/thumbnail.py brand.pptx      # writes brand.jpg, a contact sheet
python3 $SETUP/scripts/count_style.py brand.pptx           # the colours and sizes it uses
```

Open the contact sheet — delegate it to the vision-agent if you cannot see
images, asking **only** how many distinct designs it shows and whether any
repeat; never ask it about colours, which `count_style.py` counts and a vision
model guesses wrong — and **count**. Does the design count match what you
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
python3 $SETUP/scripts/make_pack.py --name acme --template brand.pptx \
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
python3 $SETUP/scripts/find_notes.py brand.pptx
```

A pattern collection talks to whoever builds from it: `EXAMPLE`, `Customer
logo here`, a paragraph of instructions on slide four. They are ordinary
shapes, nothing flags them, and they ship.

**Show the candidates and have them ticked.** Getting this wrong in one
direction leaves instructions in a customer deck; in the other it silently
deletes a column header. A person reading twelve short strings settles it in
a minute.

```bash
python3 $SETUP/scripts/make_pack.py --name acme --template brand.pptx \
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
budgets, the canvas hint. Then have the vision-agent judge what each design is
for — **one design per call, with the generated brief**:

```bash
python3 $SKILL/scripts/vision_brief.py designs $SKILL/templates/acme/catalog
```

It writes `brief.md` beside each `render.png` and lists the designs to send.
For **each** listed design, send the vision-agent exactly `Read the brief at
<catalog>/slide-NN/brief.md and follow it exactly.` and save its whole reply,
unedited, as `slide-NN/answer.md`. Never batch designs: fifteen renders in one
call come back as fifteen generic answers, and the decision table built from
them picks designs by accident. If it answers `CANNOT SEE IMAGE`, send the
same brief again (up to three tries); only then save it and tell the user.

The brief's image is `slide-NN/review.png`, the render with a **review code**
printed under it, and an answer counts only when it gives that code back.
**Never write an answer yourself**: the worksheet ignores an answer without
the code, and that design stays blank for the user to fill. After a
re-extraction only the designs whose render changed are listed again.
No vision-agent? Skip this; you will fill the blanks by hand.

Then write the worksheet:

```bash
python3 $SETUP/scripts/name_designs.py $SKILL/templates/acme/catalog
```

That writes `designs.md` beside the renders: every measured field filled, and
every judged field — `name`, `content_types`, `purpose`, `use_when`,
`avoid_when`, `content_shape`, `max_items`, `looks_best` — prefilled from the
vision-agent's answers. **Review it with the user, design by design**; the
vision-agent proposes, the user approves. Without answers the judged fields
are blank: go through the contact sheet with the user and propose one sentence
each. `content_shape` and `use_when` matter most — they are what a text-only
builder matches content against.

The measured lines are already right and must not be edited. Then import:

```bash
python3 $SETUP/scripts/name_designs.py $SKILL/templates/acme/catalog --import
```

A design left blank still builds; the decision table just never picks it for
the right reason, and the import says which ones.

Collect the template's own icons — the pictograms its designs use, usually
from Office's icon library — into the pack, so decks can place them with
`D.icon` instead of drawing their own:

```bash
python3 $SETUP/scripts/harvest_icons.py acme                     # from the template
python3 $SETUP/scripts/harvest_icons.py acme --from deck.pptx    # and a finished deck
```

It keeps small, square SVG pictures only and names them from Office's icon
id or the alt text. **Show the user the list it prints**: an icon it could not
name is `icon-<slide>-<n>` and gets a real name (and keywords) in
`templates/acme/icons/index.json`; re-running keeps those edits. Beyond the
brand's own, builds fall back to the Element Plus set the skill ships.
Assemble:

```bash
python3 $SKILL/scripts/catalog_build.py $SKILL/templates/acme/catalog
```

Inside a pack that needs no flags: it writes the pack's `catalog.md` (a
compact index) and `designs/design-N.md` (one file per design), and
splices in the pack's `overrides.md` if there is one.

## Stage 5 — Probe  ·  **ask the user**

```bash
python3 $SETUP/scripts/probe_deck.py acme     # writes probe.pptx here
python3 $SKILL/scripts/render.py probe.pptx   # writes render/slide-1.png ...
```

Three real slides: a divider, a content slide retexted through the catalog's
own keys, and a drawn slide on the pack's canvas design in the pack's own
colours. **Delegate the three PNGs to the vision-agent, one per call** (you
cannot see images yourself), and ask it questions 1, 2 and 4. **Question 3 is
for the user**, never the vision-agent: it cannot know which colours a brand
owns, and given the hex codes it simply agrees.

1. Does the divider title fit on one line, clear of the master's decoration?
2. Is any word on slide 2 still the template's own sample text?
3. Are the colours on slide 3 the ones the brand actually uses? *(the user)*
4. On slide 3, is anything of the template's own furniture showing through
   the drawing, or is the drawing sitting where the design's own title was?
   And does slide 3 look like the template's *content* slides -- the same
   title position, footer and page number -- or like a cover?

A *no* to 1, 2 or 3, or a *yes* to 4, is a five-minute fix now:

- **Title wraps:** the budget is in the catalog; note it in `overrides.md`
  under `## design-N` as a hard limit.
- **Leftover wording:** a missed authoring note. Re-run stage 3 with it added.
- **Wrong colours:** stage 2's roles were mapped wrong. Re-run `make_pack.py`
  with corrected roles; it does not touch the catalog.
- **Wrong canvas:** the catalog picked the plainest roomy design on the
  master most designs use, and on this template that is the wrong one. Fill
  the `## Canvas` block at the top of `designs.md` (the design number, and
  the shape indices to keep, as `probe.py` numbers them), re-run
  `name_designs.py --import`, and rebuild with `catalog_build.py`. The
  hand-set canvas survives every rebuild.

A finding on a shape the template itself drew past the slide edge (some
generated masters do this) is the template's, not the pack's; note it in
`overrides.md` and move on.

Then run the probe again and look again — through the vision-agent.

**You cannot see the images yourself, so the vision-agent is how this stage
is done — do not skip it.** Do not report a pack as finished on the strength
of a clean QA run. QA cannot see a title crossing a decoration, and that is
precisely what this stage is for.

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

**Never edit `catalog.md` or `designs/`**: they are generated, and the next rebuild overwrites
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
- **Nothing leaves the machine by default.** Extraction uses no model, and
  judging is done by the vision-agent on your own infrastructure. A template
  and its renders are brand-internal: never send them to a hosted service.
- **apt only.** Never run `pip` or `npm`.
- **The template file is brand-internal.** `templates/*` is gitignored in the
  presentation skill. Keep it that way.
