# presentation-skill

Two [Agent Skills](https://agentskills.io) that let a coding agent build,
edit and check slide decks in **your company's own template**, driven by a
headless LibreOffice Impress over UNO. `.pptx` and `.odp` alike, saved back in
the format they arrived in, with nothing rewritten by hand.

| Skill | What it does |
|---|---|
| [`skills/presentation`](skills/presentation/SKILL.md) | Build a deck from a brief or an outline on the installed template pack, edit an existing deck, extract text and notes, convert formats, render slides to PNG, and run the QA that finds overflow, leftover template wording, bad contrast, occlusion and drawn shapes that almost line up. |
| [`skills/presentation-setup`](skills/presentation-setup/SKILL.md) | Teach the first skill your template: an interactive, five-stage walkthrough that counts the palette and type scale off the file, asks you the four questions only a person can answer, and writes a **template pack**. Twenty minutes, once. Later, it adds the custom slides from your finished project decks to the pack as new designs. |

The skills ship **no template**. A deck is only ever built on a pack you
installed, in colours and sizes counted off your own master, and the drawing
API refuses to invent a colour when no pack is in force.

## Requirements

Linux with LibreOffice 25.2 or newer and Python 3.11 or newer. Everything
comes from apt; nothing is installed with pip or npm.

```bash
sudo apt install libreoffice-impress python3-uno python3-lxml python3-defusedxml \
                 python3-pil poppler-utils fonts-crosextra-carlito fonts-crosextra-caladea
soffice --version        # 25.2 or newer
```

The two font packages are optional but make the overflow check trustworthy
for Calibri and Cambria. A company template set in its own typeface needs that
font installed too, or LibreOffice substitutes another and every text
measurement is off.

LibreOffice and the generated build scripts run without network, inside a
network namespace made by `unshare` (util-linux 2.38 or newer, installed on
every Debian 12+) or, where `unshare` cannot be used, `bubblewrap`:

```bash
sudo apt install bubblewrap   # only if unshare is unavailable, e.g. in some containers
```

An untrusted deck cannot make anything fetch the URLs it links to. That needs
unprivileged user namespaces, which Debian enables by default; where they are
off, a warning says so and decks that link remote resources are refused. Run
the scripts with the system `python3`: a virtualenv or conda interpreter
cannot see `python3-uno`.

Icons need nothing extra: `D.icon` uses the template pack's own icons and the
293 Element Plus icons shipped in `skills/presentation/icons/ep/` (MIT).
A company icon set can be added through `$PRESENTATION_ICONS`.

An optional vision-agent adds the visual half of QA: the build agent delegates
the rendered slides to it and writes its per-slide verdicts. It is never
required, and no render is ever sent anywhere off the machine.

## Install

Both skills must sit side by side, either in one `skills/` directory or both
under `~/.claude/skills/`.

**Claude Code**, as a plugin marketplace:

```
/plugin marketplace add <owner>/presentation-skill
/plugin install presentation@presentation-skill
```

**Claude Code or any agent that reads `SKILL.md` files**, by copying:

```bash
git clone https://github.com/<owner>/presentation-skill.git
ln -s "$PWD/presentation-skill/skills/presentation"       ~/.claude/skills/presentation
ln -s "$PWD/presentation-skill/skills/presentation-setup" ~/.claude/skills/presentation-setup
```

Then check the install:

```bash
cd presentation-skill/skills/presentation
python3 tests/test_geometry.py && python3 tests/test_theme_repair.py && python3 tests/test_generic_docs.py
python3 scripts/render.py tests/fixtures/kite-brand-master.pptx     # nine PNGs in render/
```

## First run: teach it a template

Tell your agent *"set up the presentation skill with our template"* and hand
it the `.pptx`, `.potx`, `.odp` or `.otp` master. The `presentation-setup`
skill runs five stages and stops four times to ask you something:

1. **Survey**: a contact sheet of every design, so you can confirm the file
   is one slide per design.
2. **Palette and type scale**: the colours and sizes the template actually
   uses, counted. *You* say which one is the accent, which is the alert, and
   which sizes are body, caption and so on. The script refuses a colour the
   template does not contain.
3. **Authoring notes**: the strings the template says to its builder
   ("EXAMPLE", "Customer logo here"). You tick the real ones so they are
   deleted from every deck.
4. **Designs**: a worksheet with everything measurable pre-filled, and three
   blanks per design for what it is for. The vision-agent (or you) fills them.
5. **Probe**: three real slides built on the new pack, rendered, for you to
   look at — through the vision-agent. Anything wrong is a five-minute fix at
   this point.

**Later: add your project slides.** The best designs are often not in the
master but in finished project decks built on it. *"Add these slides to the
catalog"* runs `add_designs.py`. It finds the slides no design of the pack
matches, groups the ones that look alike, and has you pick one per group.
Their project wording is replaced with neutral sample text you approve, and
they are appended to the pack's template as new designs. Each one can get its
own row at the top of the decision table ("risks, each paired with their
countermeasure"). See *Adding designs from project decks* in
`skills/presentation-setup/SKILL.md`.

To try it without a template of your own, use the synthetic demo master in
`skills/presentation/tests/fixtures/kite-brand-master.pptx`. It was generated
for testing and belongs to nobody.

To drive the scripts by hand instead, `skills/presentation/README.md` is the
same pipeline as a manual.

## Then: build decks

*"Build a deck about X"* runs the storyline layer first (six intake questions,
a scaffold and a ghost deck you approve), then the recipe: one build script,
one `build_deck` call against the pack, and a single `make_deck.py` run that
QAs, schema-checks and renders every slide. The report it writes has a row per
slide for a verdict, and the deck is not finished until every row has one.

```python
from office.deckbuild import build_deck
from office.template import retext
from office import draw as D

def fill(doc, page, slot, design):
    if slot == 0:
        retext(page, {"Chapter title": "How we got here"})
    if slot == 1:
        area = D.canvas(page, keep=(0, 1))
        D.kpis(doc, page, area, [("3", "regions"), ("12%", "growth")],
               color=D.color("accent"))

build_deck("acme", "out.pptx", [4, 5], fill=fill, footer="Q3 review · 2026")
```

```bash
python3 scripts/make_deck.py build.py out.pptx --template acme
```

## Privacy

Every step runs on your machine. The extraction uses no model, and nothing
here calls a hosted one: per-deck visual QA and catalog judging both delegate
to a vision-agent, so company slides never leave the machine.
`skills/*/templates/*` is
gitignored, so a company template cannot be committed by forgetting a step.

## Layout

```
skills/presentation/          the skill: SKILL.md, scripts/, references/, tests/
  scripts/office/             the engine: UNO document handling, drawing, template editing, QA
  templates/                  packs live here (gitignored except its README)
skills/presentation-setup/    the setup skill: SKILL.md and its scripts
```

## Licence

Apache License 2.0, see [LICENSE](LICENSE). Parts of the writing guidance
were adapted from other projects under their own licences; the notices are in
[`skills/presentation/THIRD-PARTY.md`](skills/presentation/THIRD-PARTY.md).
