# Company template — layout catalog (EXAMPLE / skeleton)

> Rename to `company-template.md` once filled. This is the reference the skill
> consults to choose layouts. The per-layout entries and the decision table are
> produced by the gemma prompt (`gemma-catalog-prompt.md`) from real renders,
> then reviewed by a human. The rows below are illustrative placeholders — replace
> every one. Ship this alongside `assets/company-template.pptx` in the internal fork.

**Asset:** `assets/company-template.pptx` — one slide per design, master «REPLACE».
Build with `clone_slide` + `retext` (see "Adapting a template deck" in SKILL.md).
QA every build with `qa.py --template assets/company-template.pptx`.

**Designs are named by their slide number in that asset.** Slide `N` is
`slides(doc)[N - 1]` — what you hand to `clone_slide`. Do *not* key this catalog
on the autolayout number that `probe.py` reports as `layout`: a designed template
reports the same one on every slide, so it distinguishes nothing.

## Decision table — content → template slide

Apply top to bottom; first match wins.

| When the content is… | Use slide | Design | Because |
|---|---|---|---|
| The deck's opening slide | `1` | Title cover | Title + subtitle, brand cover |
| An agenda / table of contents | `?` | … | … |
| The start of a new section | `?` | … | Section divider, full-bleed |
| A single headline idea, no bullets | `?` | … | One-big-idea, large type |
| Standard supporting points | `?` | … | Title + bulleted body |
| Comparing two things | `?` | … | Two-column |
| A process or sequence of steps | `?` | … | Process-steps |
| Numbers / KPIs to highlight | `?` | … | Metrics tiles |
| A chart or data graphic | `?` | … | Chart placeholder |
| A pull quote / testimonial | `?` | … | Quote |
| Team or people | `?` | … | People grid |
| The closing / contact slide | `?` | … | Closing |

**Variety rule:** _(from aggregation prompt — e.g. "never run more than two
bulleted-content slides in a row; open sections with the divider and break up
long sections with a metrics or chart slide.")_

## Design entries

One block per template slide, from the per-layout prompt. Example shape:

### Slide 1 — Title cover
- **Purpose:** the deck's front page.
- **Content types:** `title`
- **Capacity:** headline ✓ · body slots 1 · max bullets 1 · images 1 (logo) · chart ✗
- **Use when:** first slide of any deck.
- **Avoid when:** anything mid-deck.
- **Imagery:** the company logo (logo, fixed); a coral abstract shape top-right (decorative, fixed).
- **Notes:** logo shape is fixed decoration — never delete it; retext the title slot only.

### Slide ? — …
- **Purpose:** …
- **Content types:** …
- **Capacity:** …
- **Use when:** …
- **Avoid when:** …
- **Imagery:** … (what each picture depicts + decorative/photo-placeholder/icon/logo/background; "none" if text-only)
- **Notes:** …
