# Gemma prompt — author the company layout catalog

Stage B of building the baked-in company template (see `README.md`). Runs on
your internal infra. `catalog_extract.py` has already produced one `slide-NN/`
per template layout, each with a `render.png` and a `structure.json`.

**Ask the model only for what needs eyes.** The retext keys, the empty slots,
the group counts, the canvas indices and the title budgets are all measured in
stage A and assembled by `catalog_build.py`; a model asked to produce them
invents plausible wrong ones. What is left here is genuine judgement: what a
layout is *for*, what it holds, what its pictures depict.

Run it **once per layout** (best quality), then run the aggregation prompt once
over all the per-layout answers. Keep a human in the loop on the output — the
model proposes, you approve.

---

## Per-layout prompt

Paste this, then attach that layout's `render.png` and the contents of its
`structure.json`.

> You are cataloguing one slide layout from our company's PowerPoint template so
> an automated deck-builder knows when to use it. You are given a rendered image
> of the layout and a JSON description of its shapes (indices, positions in
> 1/100 mm, sizes, fonts, sample text).
>
> Look at the image first — the JSON is there to confirm capacities and names,
> not to override what you can see.
>
> Return **only** a JSON object, no prose, with exactly these fields:
>
> ```json
> {
>   "name": "<short human name you'd say out loud, e.g. 'Section divider'>",
>   "purpose": "<one sentence: what this slide is FOR>",
>   "content_types": ["<tags from the fixed list below>"],
>   "capacity": {
>     "headline": <true|false>,
>     "body_slots": <int: how many distinct text bodies/columns>,
>     "max_bullets_per_body": <int, your read of what fits without crowding>,
>     "images": <int: image/logo placeholders>,
>     "supports_chart": <true|false>
>   },
>   "use_when": "<one sentence: the trigger to reach for this layout>",
>   "avoid_when": "<one sentence: content that does NOT belong here>",
>   "imagery": [
>     {"depicts": "<concretely WHAT the picture shows — 'a bird in flight', 'a
>                   city skyline at dusk', 'two hands shaking'; never just
>                   'an image'>",
>      "role": "<decorative | photo-placeholder | icon | logo | background>"}
>   ],
>   "notes": "<optional: quirks a builder must respect — fixed decoration that
>             must not be deleted, a slot that only holds one line, etc.>"
> }
> ```
>
> `content_types` must be drawn from exactly this fixed vocabulary (pick all that
> apply): `title`, `agenda`, `section-divider`, `one-big-idea`, `bulleted-content`,
> `two-column`, `comparison`, `process-steps`, `timeline`, `metrics-kpi`, `quote`,
> `team-people`, `image-showcase`, `data-chart`, `table`, `closing`, `contact`.
>
> **Describe every picture you can see.** For each photograph, illustration,
> icon or logo on the slide, name what it actually depicts — "a bird", "a bar
> chart sketch", "the company logo" — not "an image". Put one entry per picture
> in `imagery`, and set `role`: `photo-placeholder` if it looks like sample
> art meant to be swapped, `decorative`/`background` if it is fixed brand
> styling, `icon`/`logo` otherwise. Empty list `[]` if the slide has no pictures.
> The JSON's shape list will not tell you this — only the rendered image does.
>
> Be honest about capacity — if the body box is sized for one line, say
> `max_bullets_per_body: 1`. Overstating capacity is the mistake that makes decks
> overflow. Do not invent slots the layout does not have.

The script stamps `"slide"`, `"layout"` and `"master"` onto each entry from the
manifest, so the model is never asked for them.

### Why the identifier is the *slide* number

`structure.json` carries a `layout` field, and it is tempting to key the catalog
on it. Don't. It is the master's autolayout id, and on a designed template it is
the same number on every slide — the demo template in `tests/fixtures/` reports `layout: 20, master:
DEFAULT` for all nine of its designs. A decision table keyed on it points every
rule at the same place and is worthless.

The design's identity is its **position in the template deck**: slide `N` is
`slides(doc)[N - 1]`, which is exactly what `clone_slide` takes. That is the
number the catalog carries and the number the decision table cites.

---

## Aggregation prompt

After collecting every per-layout JSON, paste them all and this instruction:

> Here are the catalogue entries for every design in our template. Each entry's
> `slide` field is that design's number in the template deck — the only
> identifier a deck-builder can act on. Ignore `layout`: it is the master's
> autolayout id and is usually identical on every slide, so it names nothing.
> Produce two things.
>
> 1. A **decision table** as JSON: an ordered list of rules mapping content to a
>    template slide, most specific first. Each rule is
>    `{"when": "<content situation>", "use_slide": <int>, "because": "<short>"}`.
>    Every `use_slide` **must** be a slide number that appears in the entries —
>    never invent one, never emit a layout number. Cover at least: deck opener,
>    agenda, each section start, a single headline idea, standard bulleted
>    content, comparing two things, showing a process or timeline, presenting
>    numbers/KPIs, a chart, a quote, people, and the closing slide. If two slides
>    serve the same purpose, pick one as default and note the other's number as
>    the alternative in `because`.
>
> 2. A **variety rule**: one sentence on how to avoid every slide being
>    title-and-bullets, referencing which template slides to interleave.
>
> Return JSON: `{"decision_table": [...], "variety_rule": "..."}`. No prose.

The script checks every `use_slide` against the entries it fed in and warns on
any rule citing a slide the template does not have — a model that fell back to
layout ids shows up there rather than in the shipped catalog.

---

## What you do with the output

Nothing by hand:

```bash
scripts/catalog_build.py templates/acme/catalog     # writes templates/acme/catalog.md
```

`catalog_build.py` merges these judged fields with the measured ones from stage
A and writes the catalog in its shipped format. Then **review every capacity
figure and every decision-table row against the renders** -- that human pass is
where the catalog earns its keep -- and commit the catalog and the `.pptx` asset
to your **internal** fork only. `references/catalog.example.md` shows
the finished shape, annotated with which half comes from where.
