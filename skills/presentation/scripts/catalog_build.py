#!/usr/bin/env python3
"""
Assemble a layout catalog from the extraction and the model pass -- stage C.

    python3 scripts/catalog_build.py templates/acme/catalog/ \
        -o templates/acme/catalog.md --asset templates/acme/template.pptx

Reads what ``catalog_extract.py`` measured (``manifest.json`` and each
``slide-NN/structure.json``) and what the vision model judged
(``catalog/entries/slide-NN.json`` and ``catalog/decision_table.json``), and
writes the catalog the skill reads at build time.

**Why this exists.**  The instruction used to be "fold the model's JSON into
the catalog by hand".  What that produced, once, was a good
catalog -- and the fields that made it good were the ones added by hand
afterwards: the ``retext`` keys, the empty slots, the canvas indices, the title
budgets.  None of them survived into the skeleton, so the next template would
have started from the model's answer alone and lost them.  Now the mechanical
fields come from stage A, the judgement comes from stage B, and this puts them
together the same way every time.

Review the result against the renders before shipping it: the model overstates
capacity, and this script cannot tell that it did.
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

# The palette block below reads the same profile the engine reads, through the
# same accessors, so the catalog cannot document a palette the build would not
# hand out.  (These pull in ``uno``; this script runs on the infra that ran
# ``catalog_extract.py``, which needs LibreOffice anyway.)
from office.draw import contrast_ratio  # noqa: E402
from office.pack import SKILL, TEMPLATE_EXTS, pack_for  # noqa: E402
from office.template import palette, palette_roles, type_scale  # noqa: E402

#: A title box at or above this size is a display slot -- a divider headline
#: set into master decoration -- where a second line is a defect no check can
#: see.  Below it, wrapping is normal and a budget would be noise.
DISPLAY_PT = 32.0

#: The content situations a decision table should cover, each with the words
#: that would appear in a rule covering it.  Anything missing gets a warning in
#: the catalog rather than silence: "there is no cover design" is exactly the
#: kind of thing a deck-builder must be told once, loudly.
#:
#: Each entry is a *word stem*, matched at a word boundary -- ``compar`` covers
#: "comparison" and "comparing", ``end`` covers "end" and "ending".  Matching
#: these as bare substrings, which is what this did first, is silently wrong in
#: the expensive direction: ``end`` is inside **agenda**, and the aggregation
#: prompt demands an agenda row, so "there is no closing design" could never be
#: reported for any template.  ``recommendations`` and ``trend`` hid it too.
EXPECTED_RULES = [
    ("a cover or deck opener", ("opener", "cover", "title slide", "front")),
    ("an agenda", ("agenda", "contents", "overview")),
    ("a section divider", ("section", "divider", "chapter")),
    ("a single headline idea", ("headline", "one big", "single idea", "statement")),
    ("bulleted content", ("bullet", "prose", "body text", "standard content")),
    ("a comparison", ("compar", "versus", "vs", "two things", "before/after")),
    ("a process or timeline", ("process", "timeline", "roadmap", "phase", "step")),
    ("numbers or KPIs", ("kpi", "number", "metric", "financial", "cost")),
    # No bare "data": a row about "customer data" is not a chart design, and
    # matching it hides the gap.  "chart" and "graph" name the thing itself.
    ("a chart", ("chart", "graph", "plot")),
    ("a quote", ("quote", "testimonial")),
    # "person" would match "personal"; the words a real rule uses are these.
    ("people", ("people", "team", "org chart", "personnel", "headshot")),
    ("a closing slide", ("closing", "close", "final", "thank", "end")),
]


def _load(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return default
    except json.JSONDecodeError as e:
        raise SystemExit(f"{path}: not valid JSON ({e})")


def _addr(a):
    return f"`[{a}]`"


def _capacity(cap):
    if not cap:
        return "_(not catalogued)_"
    tick = "✓" if cap.get("headline") else "✗"
    chart = "✓" if cap.get("supports_chart") else "✗"
    return (f"headline {tick} · body slots {cap.get('body_slots', '?')} · "
            f"max bullets {cap.get('max_bullets_per_body', '?')} · "
            f"images {cap.get('images', '?')} · chart {chart}")


def _imagery(entry, structure):
    """What the model saw, cross-checked against how many pictures there are."""
    seen = entry.get("imagery") or []
    described = "; ".join(
        f"{i.get('depicts', '?')} ({i.get('role', '?')})" for i in seen)
    count = len(structure.get("images", []))
    if not seen:
        return "none" if not count else f"_{count} picture(s), not described_"
    if count and len(seen) != count:
        described += f"  _[{count} picture shape(s) in the file]_"
    return described


#: The places an overrides file may splice prose into the generated catalog.
#: ``design-N`` is accepted for any design as well.
ANCHORS = ("header", "after-decision-table", "before-design-entries",
           "design-entries-intro", "footer")


def _load_overrides(path):
    """``{anchor: markdown}`` from a file of ``## anchor`` sections.

    **The reason this exists.**  Everything below the model's judgement is
    generated, so a reviewer who improves the catalog by editing it improves a
    file the next run overwrites.  That is not hypothetical: the house
    catalog's hardest-won paragraphs -- that the footer lives on the master and
    no per-slide check can see it, that design 20 is a dense slide rather than
    a blank one, that a title slide has to be built on a canvas because the
    template has no cover -- existed only in the shipped file, while the same
    file had silently fallen twelve retext keys behind its own template.
    Prose goes here; measurements come from the extractor; neither overwrites
    the other.
    """
    if not path:
        return {}
    text = Path(path).read_text(encoding="utf-8")
    blocks = {}
    # An anchor is a level-2 heading holding exactly one word, and the
    # lookahead must be that same rule -- not a laxer version of it.  When it
    # was laxer, the first ordinary ``## Multi word heading`` *inside* a block
    # ended the block, and everything after it vanished from the catalog with
    # no warning: the failure this file exists to prevent, reintroduced in the
    # file that prevents it.  Write headings inside a block as ``###`` or as
    # several words.
    for m in re.finditer(
            r"^##[ \t]+(\S+)[ \t]*$(.*?)(?=^##[ \t]+\S+[ \t]*$|\Z)",
            text, re.M | re.S):
        body = m.group(2).strip("\n")
        if body.strip():
            blocks[m.group(1).strip().lower()] = body
    return blocks


def _entry_block(number, entry, structure, extra=None):
    name = entry.get("name") or f"Design {number}"
    out = [f"### Slide {number} — {name}"]
    out.append(f"- **Purpose:** {entry.get('purpose', '_(not catalogued)_')}")
    types = entry.get("content_types") or []
    out.append("- **Content types:** "
               + (", ".join(f"`{t}`" for t in types) or "_(none)_"))
    out.append(f"- **Capacity:** {_capacity(entry.get('capacity'))}")

    keys = structure.get("retext_keys") or []
    if keys:
        out.append("- **Retext keys:** " + " · ".join(
            f'{_addr(k["address"])} "{k["key"]}"' for k in keys))
    else:
        out.append("- **Retext keys:** _none — every slot is empty; write with "
                   "`set_text(shapes(page)[i], ...)` or draw your own_")

    empty = structure.get("empty_slots") or []
    if empty:
        out.append("- **Empty slots** (no text, so `retext` cannot match): "
                   + " · ".join(
                       f'{_addr(s["address"])} {s.get("placeholder") or s["name"]}'
                       for s in empty))

    counts = structure.get("text_counts") or {}
    if counts.get("grouped"):
        out.append(f"- **Text in groups:** {counts['grouped']} shape(s) inside "
                   f"groups against {counts.get('top_level', 0)} at the top "
                   f"level — `shapes(page)[i]` reaches none of them; retext by "
                   f"wording")

    budgets = [b for b in structure.get("title_budgets") or []
               if b["pt"] >= DISPLAY_PT]
    for b in budgets:
        out.append(f"- **Title budget:** {_addr(b['address'])} holds about "
                   f"{b['chars_per_line']} characters on one line at "
                   f"{b['pt']:g}pt (box {b['width']} wide). Over that it wraps, "
                   f"and a display title wraps into the master's own decoration")

    out.append(f"- **Use when:** {entry.get('use_when', '_(not catalogued)_')}")
    out.append(f"- **Avoid when:** {entry.get('avoid_when', '_(not catalogued)_')}")
    out.append(f"- **Imagery:** {_imagery(entry, structure)}")
    if entry.get("notes"):
        out.append(f"- **Notes:** {entry['notes']}")
    if extra:
        out.append(extra)
    return "\n".join(out)


#: A "free area" larger than this share of the slide means the design has no
#: chrome worth keeping -- a full-bleed cover, not a content slide.  Drawing on
#: one of those is drawing on a blank page, which is what the canvas pattern
#: exists to avoid.
MAX_CANVAS_SHARE = 0.85


def _canvas_design(structures, slide_area):
    """``(number, qualified)`` -- the design a canvas should be cut from.

    The biggest free area that still leaves the design's own chrome standing:
    a title, a topic label, a source line.  That is the whole point -- you keep
    the master, the margins and the fonts, and you get an empty middle.  A
    full-bleed design whose "free area" is the entire slide gives you none of
    that, so it is excluded however large it measures.

    A catalog with no canvas row is a catalog that pushes every diagram into
    the nearest bullet list, which is the defect this row exists to prevent.

    ``qualified`` is False when nothing cleared the bar and the answer is the
    least-bad candidate rather than a good one.  **The fallback obeys the
    full-bleed guard too.**  It used to be tracked before the guards ran, so on
    a template where no design cleared ``single_body`` -- and on the
    template exactly one does -- it returned the design with the largest free
    area of all, which is the full-bleed cover this function exists to reject.
    The catalog would then have told every deck to draw on a blank page, in the
    confident voice of a measured field.
    """
    best, best_area, fallback, fallback_area = None, 0, None, 0
    for number, s in structures.items():
        canvas = s.get("canvas") or {}
        box = canvas.get("free_box")
        if not box:
            continue
        area = box["w"] * box["h"]
        if area > slide_area * MAX_CANVAS_SHARE:
            continue                      # full bleed: no chrome to inherit
        if area > fallback_area:
            fallback, fallback_area = number, area
        # Two conditions, and the second is the one that matters: the free area
        # has to come from *one* body placeholder, not from clearing a slide
        # full of sample content.  A design whose canvas drops nine shapes is
        # an org chart being demolished; a design whose canvas drops one is the
        # content slide the designer left room on.
        keeps_chrome = len(canvas.get("keep", [])) >= 2
        single_body = len(canvas.get("drops", [])) <= 2
        if keeps_chrome and single_body and area > best_area:
            best, best_area = number, area
    return (best, True) if best else (fallback, False)


def _decision_table(rules, structures, canvas_number, entries, qualified=True):
    known = set(structures)
    lines = ["| When the content is… | Use slide | Design | Because |",
             "|---|---|---|---|"]
    for rule in rules:
        slide = rule.get("use_slide")
        if slide not in known:
            because = (f"⚠ the model cited slide {slide}, which this template "
                       f"does not have — fix or delete this row")
            lines.append(f"| {rule.get('when', '?')} | `?` | — | {because} |")
            continue
        name = (entries.get(slide, {}) or {}).get("name", "")
        lines.append(f"| {rule.get('when', '?')} | `{slide}` | {name} | "
                     f"{rule.get('because', '')} |")

    # The canvas row goes LAST, and this is not a cosmetic decision.  It sat at
    # the top, where "a diagram, chart, matrix, comparison" is the first match
    # for nearly any slide worth building -- and four separate runs against the
    # test template duly built [8] * N, every slide drawn from scratch on a
    # deck whose twenty other designs they never opened.  Last, and worded as
    # the fallback it is, the rows above it get read first.
    if canvas_number:
        canvas = structures[canvas_number]["canvas"]
        box = canvas["free_box"]
        keep = ", ".join(str(k) for k in canvas["keep"])
        warn = "" if qualified else (
            " ⚠ **Unverified:** no design in this template strips back cleanly "
            "— every one of them either keeps too little of its own chrome or "
            "drops a slideful of sample content. This is the least-bad "
            "candidate, not a good one: run `probe.py` on it, check the `keep` "
            "tuple by hand, and correct this row.")
        lines.append(
            "| **a diagram, chart, matrix, comparison — anything no row above "
            f"fits** | `{canvas_number}` | **Canvas** | Strip design "
            f"{canvas_number} back with `draw.canvas(page, keep=({keep}))` and "
            f"draw in the area the body vacates: **{box['x']}, {box['y']} to "
            f"{box['x'] + box['w']}, {box['y'] + box['h']}**. You keep the "
            "master, background, chrome, fonts and margins. Better than forcing "
            "the content into the nearest bullet list — and worse than a design "
            f"that already fits it.{warn} |")
    return "\n".join(lines)


def _palette_block(asset):
    """The palette and type scale from the profile beside the asset, as prose.

    Emitted here rather than left to the overrides file because they are
    measured facts about one template, like the retext keys -- and because the
    runs that needed them did not have them: four decks went out in an orange
    and a purple this template does not contain, and a footnote came out 9pt on
    one slide and 12pt on the next.
    """
    profile = pack_for(asset).profile() if pack_for(asset) else {}
    if not profile:
        # A template profiled in place rather than installed as a pack.
        try:
            profile = json.loads(Path(asset).with_suffix(".profile.json")
                                 .read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return []

    colours = palette(profile)
    roles = palette_roles(profile)
    scale = type_scale(profile)
    if not colours and not scale:
        return []

    out = ["## Palette and type scale", "",
           "**Counted off the asset, not chosen.** Every colour here is one "
           "this template fills a shape or sets text in, and every size is one "
           "it types at. Take them from `office.draw` rather than writing hex "
           "in a build script:", "",
           "```python",
           "D.color(\"cyan\")        D.color(\"accent\")     # by name, or by role",
           "D.series(3)            # the first three, in order, for n things",
           "D.text(..., role=\"footnote\")                 # the size, by name",
           "```", ""]

    if colours:
        named = {v: k for k, v in colours.items()}
        out += ["| Colour | Hex | Reads as text on white |", "|---|---|---|"]
        for name, value in colours.items():
            ratio = contrast_ratio(value, 0xFFFFFF)
            verdict = ("yes" if ratio >= 4.5 else
                       "headlines only" if ratio >= 3.0 else
                       "**no — fill colour**")
            out.append(f"| `{name}` | `#{value:06X}` | {ratio:.1f}:1 — {verdict} |")
        out.append("")
        pairs = [(role, named.get(v, f"#{v:06X}"))
                 for role, v in roles.items() if not isinstance(v, list)]
        if pairs:
            out.append("**Roles:** " + ", ".join(
                f"`{role}` → `{name}`" for role, name in pairs) + ".")
        if isinstance(roles.get("series"), list):
            out.append("**Series order:** " + ", ".join(
                f"`{named.get(v, f'#{v:06X}')}`" for v in roles["series"])
                + " — the nth colour for the nth thing on a figure.")
        out += ["", "Most of these are **fill** colours: a brand picks them "
                "against white paper, not against the 3:1 a word has to clear "
                "to be read from a seat. Set text in them only where the table "
                "above says so — `qa.py` reports the rest.", ""]

    if scale:
        out += ["| Role | pt |", "|---|---|"]
        out += [f"| `{role}` | {size:g} |" for role, size in
                sorted(scale.items(), key=lambda kv: -kv[1])]
        out += ["", f"`D.text(..., role=…)` sets these; `D.size_for(role)` is "
                f"the number. Anything under {min(scale.values()):g}pt is "
                f"smaller than this template ever types and `qa.py` reports "
                f"it.", ""]
    return out


def _gaps(rules):
    """Content situations no rule covers, by keyword.

    Crude on purpose: a false positive costs a line the reviewer deletes, a
    false negative costs a deck-builder an hour discovering that the template
    has no cover design.  So the match is anchored at a **word boundary** and
    nowhere else -- see ``EXPECTED_RULES``.  A stem that matched mid-word
    reported silence, which is the costly answer.
    """
    covered = " ".join((r.get("when") or "").lower() for r in rules)
    return [label for label, words in EXPECTED_RULES
            if not any(re.search(rf"\b{re.escape(word)}", covered)
                       for word in words)]


#: Content types a design carries that make it a divider rather than a body
#: slide.  Written by the judging pass; matched loosely because models phrase
#: it "section-divider", "section divider" and "divider" interchangeably.
DIVIDER_TYPES = re.compile(r"section[-_ ]?divider|^divider$", re.I)


def _roles_block(canvas, rules, entries):
    """The three designs the generic documentation needs by role, not number.

    `references/` may not name a design number -- a number is a fact about one
    template and wrong for every other reader -- so it says "the canvas design"
    and "a divider design" and this block is where those resolve. Without it
    the pointer is a 500-line file and a search.

    Every value is derived, never judged: the canvas is the design the extractor
    picked, the dividers are the designs whose content types say so, and the
    cover is whichever decision-table row is about opening a deck. A template
    with no cover design says so out loud, because the commonest thing a build
    does with a missing cover is press a divider into service and overrun its
    title budget.
    """
    dividers = sorted(
        number for number, entry in entries.items()
        if any(DIVIDER_TYPES.search(str(kind))
               for kind in entry.get("content_types", ()) or ())
    )
    cover = None
    for rule in rules:
        when = str(rule.get("when") or rule.get("content") or "")
        if re.search(r"\b(deck )?opener|title slide|cover\b", when, re.I):
            try:
                cover = int(str(rule.get("design")).strip("`"))
            except (TypeError, ValueError):
                cover = None
            break

    rows = [("canvas", f"`{canvas}`" if canvas else "—",
             "anything the decision table has no row for: a diagram, a chart, "
             "a comparison you draw yourself"),
            ("divider", ", ".join(f"`{n}`" for n in dividers) or "—",
             "a section break; one line of display type"),
            ("cover", f"`{cover}`" if cover else "**none — build one**",
             "the deck's first slide, with a subtitle and a date")]
    out = ["## Designs by role", "",
           "What the generic documentation means when it says *the canvas "
           "design* or *a divider design*. These three are the only designs "
           "anything outside this pack refers to.", "",
           "| Role | Design | What it carries |", "|---|---|---|"]
    out += [f"| **{role}** | {value} | {what} |" for role, value, what in rows]
    if not cover:
        out += ["", "**There is no cover design.** The dividers are one-line "
                "slides and a real title with a subtitle and a date does not "
                "fit their budget. Build the opener on the canvas design — "
                "`references/storyline.md` stage 3 says what goes on it."]
    return out + [""]


def build(catalog: Path, asset: str, title: str, overrides=None) -> str:
    blocks = _load_overrides(overrides)
    manifest = _load(catalog / "manifest.json")
    if manifest is None:
        raise SystemExit(
            f"no manifest.json in {catalog} — point this at the directory "
            f"catalog_extract.py wrote, not at a slide-NN/ subfolder")

    structures, entries = {}, {}
    for row in manifest["slides"]:
        n = row["number"]
        structures[n] = _load(catalog / f"slide-{n:02d}" / "structure.json", {})
        entries[n] = _load(catalog / "entries" / f"slide-{n:02d}.json", {}) or {}

    aggregate = _load(catalog / "decision_table.json", {}) or {}
    rules = aggregate.get("decision_table") or []
    variety = aggregate.get("variety_rule")

    missing_entries = [n for n, e in entries.items() if not e]
    canvas_number, canvas_ok = _canvas_design(
        structures, manifest["slide_width"] * manifest["slide_height"])
    masters = sorted({row["master"] for row in manifest["slides"]})
    grouped = [row["number"] for row in manifest["slides"]
               if row.get("grouped_text")]

    out = [f"# {title}", "",
           "The pack this skill builds from when the user names no template of",
           "their own. Generated by `catalog_extract.py` + `catalog_build.py`",
           "from the real deck, with the judgement fields written by a local"
           " vision model,",
           "then reviewed against the renders.", "",
           f"**Asset:** `{asset}` — {len(manifest['slides'])} designs, "
           f"**{manifest['slide_width']} × {manifest['slide_height']}** in "
           f"1/100 mm (read `slide_size(doc)`; do not assume the LibreOffice "
           f"default 28000 × 15750). Masters "
           + ", ".join(f"`{m}`" for m in masters) + ".", "",
           "**Build with `build_deck`** — it duplicates rather than rebuilds, so "
           "bullets, inline colour and any OLE objects survive; `clone_slide` "
           "loses all three here. Check every build with `make_deck.py build.py "
           f"out.pptx --template {asset}`, and open the renders it prints.", "",
           "**Designs are named by their slide number in that asset.** Design "
           "`N` is `slides(doc)[N - 1]`, and `N` is what you put in "
           "`build_deck`'s design list. Do *not* key on the autolayout number "
           "`probe.py` reports as `layout` — on a designed template it is the "
           "same on every slide and distinguishes nothing.", ""]

    out += _roles_block(canvas_number, rules, entries)

    if "header" in blocks:
        out += [blocks.pop("header"), ""]

    if grouped:
        out += [f"**Text inside groups.** Designs {', '.join(str(g) for g in grouped)} "
                f"carry text in grouped shapes, which `shapes(page)` cannot "
                f"reach. `retext`, `find_shapes` and `qa.py` all walk groups, "
                f"and `probe.py` lists them indented as `[6.2]`. Each entry "
                f"below notes its count. **Retext by wording, never by index.**",
                ""]

    out += ["## Decision table — content → template slide", "",
            "Apply top to bottom; first match wins. **Read to the bottom "
            "before you pick** — the canvas row is last, and it means what it "
            "says: nothing above fits.", ""]
    out.append(_decision_table(rules, structures, canvas_number, entries,
                               qualified=canvas_ok)
               if rules else
               "_No decision table: run the aggregation pass "
               "(`catalog_gemma.py catalog/ --aggregate`) and rebuild._")
    out.append("")
    if variety:
        out += [f"**Variety rule:** {variety}", ""]

    gaps = _gaps(rules) if rules else []
    if gaps:
        out += ["**No design for:** " + ", ".join(gaps)
                + ". Build these on the canvas row rather than forcing "
                "them into the nearest design — and say so here once you have "
                "checked the renders.", ""]

    if "after-decision-table" in blocks:
        out += [blocks.pop("after-decision-table"), ""]

    budgets = [(n, b) for n in sorted(structures)
               for b in (structures[n].get("title_budgets") or [])
               if b["pt"] >= DISPLAY_PT]
    if budgets:
        out += ["**Display titles hold one line.** These slots set their title "
                "large, in a gap the master's decoration leaves. A second line "
                "crosses that decoration, and nothing on the page collides — "
                "`qa.py` cannot see it. Measured budgets:", "",
                "| Design | Title box width | pt | One line holds |",
                "|---|---|---|---|"]
        out += [f"| `{n}` | {b['width']} | {b['pt']:g} | "
                f"~{b['chars_per_line']} characters |" for n, b in budgets]
        out += ["", "Over budget, in order of preference: **shorten the "
                "title**; or `shrink_to_fit(doc, page, shape, max_lines=1)`. Do "
                "not add a second text box beside the placeholder.", ""]

    out += _palette_block(asset)

    if "before-design-entries" in blocks:
        out += [blocks.pop("before-design-entries"), ""]

    out += ["## Design entries", "",
            "Each entry's **Retext keys** are the strings the design actually "
            "carries, in `retext`/`find_shapes` form, transcribed from the file "
            "— copy one from here rather than guessing. Any distinctive "
            "substring of a listed string works. `[3.1]` addresses a shape "
            "*inside* the group at `[3]`.", ""]
    if "design-entries-intro" in blocks:
        out += [blocks.pop("design-entries-intro"), ""]
    for row in manifest["slides"]:
        n = row["number"]
        out += [_entry_block(n, entries[n], structures[n],
                             blocks.pop(f"design-{n}", None)), ""]

    if "footer" in blocks:
        out += [blocks.pop("footer"), ""]

    out += ["---", "",
            "**Provenance.** Structure, retext keys, canvas indices and title "
            "budgets are measured by `catalog_extract.py`; purpose, capacity, "
            "imagery and the decision table are written by a local vision model "
            "from each slide's render and reviewed by hand; the prose blocks "
            "come from the overrides file. The asset, the renders and this file "
            "are brand-internal: commit them to the internal fork only."]

    if missing_entries:
        out += ["", f"> **Incomplete:** no model entry for design(s) "
                f"{', '.join(str(n) for n in missing_entries)} — their "
                f"judgement fields read _(not catalogued)_. Run the per-layout "
                f"pass for those slides and rebuild."]

    # An override that no longer lands is prose someone wrote and nobody will
    # ever read again -- the same silent loss this mechanism exists to stop.
    if blocks:
        for anchor in sorted(blocks):
            print(f"WARNING: override block '{anchor}' matched no anchor in "
                  f"this catalog", file=sys.stderr)
        out += ["", f"> ⚠ **Unplaced overrides:** "
                + ", ".join(f"`{a}`" for a in sorted(blocks))
                + f". Valid anchors are {', '.join(f'`{a}`' for a in ANCHORS)} "
                + "and `design-N` for a design this template has."]
    return "\n".join(out) + "\n"


def _asset_beside(catalog: Path):
    """The template of the pack a ``catalog/`` directory sits in, or ``None``.

    ``catalog_build.py templates/acme/catalog/`` should not also have to be
    told where ``templates/acme/template.pptx`` is.  The old default was one
    hardcoded path, which worked for exactly one company.
    """
    pack = pack_for(catalog.resolve())
    if pack is not None and pack.template is not None:
        # Skill-relative if it lives in the skill: the catalog is read by
        # someone writing a build script, and an absolute path from the
        # machine that generated it is noise to them and drift to `--check`.
        try:
            return str(pack.template.relative_to(SKILL))
        except ValueError:
            return str(pack.template)
    for ext in TEMPLATE_EXTS:
        candidate = catalog.resolve().parent / f"template{ext}"
        if candidate.is_file():
            return str(candidate)
    return None


def main():
    ap = argparse.ArgumentParser(
        description="Assemble a pack's catalog.md from its catalog/")
    ap.add_argument("catalog", help="The directory catalog_extract.py wrote")
    ap.add_argument("-o", "--output", help="Write here (default: stdout)")
    ap.add_argument("--asset", default=None,
                    help="Path the skill will load the template from")
    ap.add_argument("--title", default="Company template — layout catalog")
    ap.add_argument("--overrides",
                    help="Markdown file of '## anchor' prose blocks spliced "
                         "into the generated catalog. Anchors: "
                         + ", ".join(ANCHORS) + ", design-N")
    ap.add_argument("--check", action="store_true",
                    help="Do not write: rebuild and diff against --output, "
                         "exit 1 if it has drifted. Re-extract first, so this "
                         "compares the shipped catalog with the template as it "
                         "is now.")
    args = ap.parse_args()

    asset = args.asset or _asset_beside(Path(args.catalog))
    if asset is None:
        ap.error("--asset: no template found beside %s. Pass the path to the "
                 "template this catalog was extracted from." % args.catalog)
    text = build(Path(args.catalog), asset, args.title, args.overrides)

    if args.check:
        if not args.output:
            ap.error("--check needs --output: the catalog to compare against")
        current = Path(args.output).read_text(encoding="utf-8")
        if current == text:
            print(f"{args.output} is up to date with {args.catalog}")
            return
        import difflib
        sys.stdout.writelines(difflib.unified_diff(
            current.splitlines(keepends=True), text.splitlines(keepends=True),
            fromfile=f"{args.output} (shipped)", tofile="rebuilt", n=1))
        print(f"\n{args.output} has drifted from the template. Rebuild it, or "
              f"move whatever is only in the shipped file into --overrides.",
              file=sys.stderr)
        sys.exit(1)

    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Wrote {args.output} ({len(text.splitlines())} lines)")
        print("Review every capacity figure and decision-table row against the "
              "renders before shipping it.")
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
