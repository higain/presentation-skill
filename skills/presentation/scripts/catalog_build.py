#!/usr/bin/env python3
"""
Assemble a layout catalog from the extraction and the model pass -- stage C.

    python3 scripts/catalog_build.py catalog/ -o references/company-template.md

Reads what ``catalog_extract.py`` measured (``manifest.json`` and each
``slide-NN/structure.json``) and what the vision model judged
(``catalog/entries/slide-NN.json`` and ``catalog/decision_table.json``), and
writes the catalog the skill reads at build time.

**Why this exists.**  The instruction used to be "fold the model's JSON into
references/company-template.md by hand".  What that produced, once, was a good
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
import sys
from pathlib import Path

#: A title box at or above this size is a display slot -- a divider headline
#: set into master decoration -- where a second line is a defect no check can
#: see.  Below it, wrapping is normal and a budget would be noise.
DISPLAY_PT = 32.0

#: The content situations a decision table should cover, each with the words
#: that would appear in a rule covering it.  Anything missing gets a warning in
#: the catalog rather than silence: "there is no cover design" is exactly the
#: kind of thing a deck-builder must be told once, loudly.
EXPECTED_RULES = [
    ("a cover or deck opener", ("opener", "cover", "title slide", "front")),
    ("an agenda", ("agenda", "contents", "overview")),
    ("a section divider", ("section", "divider", "chapter")),
    ("a single headline idea", ("headline", "one big", "single idea", "statement")),
    ("bulleted content", ("bullet", "prose", "body text", "standard content")),
    ("a comparison", ("compar", "versus", " vs", "two things", "before/after")),
    ("a process or timeline", ("process", "timeline", "roadmap", "phase", "step")),
    ("numbers or KPIs", ("kpi", "number", "metric", "financial", "cost")),
    ("a chart", ("chart", "graph", "data")),
    ("a quote", ("quote", "testimonial")),
    ("people", ("people", "team", "org", "person")),
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
    described = " ; ".join(
        f"{i.get('depicts', '?')} ({i.get('role', '?')})" for i in seen)
    count = len(structure.get("images", []))
    if not seen:
        return "none" if not count else f"_{count} picture(s), not described_"
    if count and len(seen) != count:
        described += f"  _[{count} picture shape(s) in the file]_"
    return described


def _entry_block(number, entry, structure):
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
    return "\n".join(out)


#: A "free area" larger than this share of the slide means the design has no
#: chrome worth keeping -- a full-bleed cover, not a content slide.  Drawing on
#: one of those is drawing on a blank page, which is what the canvas pattern
#: exists to avoid.
MAX_CANVAS_SHARE = 0.85


def _canvas_design(structures, slide_area):
    """The design a canvas should be cut from.

    The biggest free area that still leaves the design's own chrome standing:
    a title, a topic label, a source line.  That is the whole point -- you keep
    the master, the margins and the fonts, and you get an empty middle.  A
    full-bleed design whose "free area" is the entire slide gives you none of
    that, so it is excluded however large it measures.

    A catalog with no canvas row is a catalog that pushes every diagram into
    the nearest bullet list, which is the defect this row exists to prevent.
    """
    best, best_area, fallback, fallback_area = None, 0, None, 0
    for number, s in structures.items():
        canvas = s.get("canvas") or {}
        box = canvas.get("free_box")
        if not box:
            continue
        area = box["w"] * box["h"]
        if area > fallback_area:
            fallback, fallback_area = number, area
        # Two conditions, and the second is the one that matters: the free area
        # has to come from *one* body placeholder, not from clearing a slide
        # full of sample content.  A design whose canvas drops nine shapes is
        # an org chart being demolished; a design whose canvas drops one is the
        # content slide the designer left room on.
        keeps_chrome = len(canvas.get("keep", [])) >= 2
        single_body = len(canvas.get("drops", [])) <= 2
        if (keeps_chrome and single_body
                and area <= slide_area * MAX_CANVAS_SHARE and area > best_area):
            best, best_area = number, area
    return best or fallback


def _decision_table(rules, structures, canvas_number, entries):
    known = set(structures)
    lines = ["| When the content is… | Use slide | Design | Because |",
             "|---|---|---|---|"]
    if canvas_number:
        canvas = structures[canvas_number]["canvas"]
        box = canvas["free_box"]
        keep = ", ".join(str(k) for k in canvas["keep"])
        lines.append(
            "| **a diagram, chart, matrix, comparison — anything no row below "
            f"fits** | `{canvas_number}` | **Canvas** | Strip design "
            f"{canvas_number} back with `draw.canvas(page, keep=({keep}))` and "
            f"draw in the area the body vacates: **{box['x']}, {box['y']} to "
            f"{box['x'] + box['w']}, {box['y'] + box['h']}**. You keep the "
            "master, background, chrome, fonts and margins. **Reach for this "
            "rather than forcing the content into the nearest bullet list.** |")
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
    return "\n".join(lines)


def _gaps(rules):
    """Content situations no rule covers, by keyword.

    Crude on purpose: a false positive costs a line the reviewer deletes, a
    false negative costs a deck-builder an hour discovering that the template
    has no cover design.
    """
    covered = " ".join((r.get("when") or "").lower() for r in rules)
    return [label for label, words in EXPECTED_RULES
            if not any(word in covered for word in words)]


def build(catalog: Path, asset: str, title: str) -> str:
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
    canvas_number = _canvas_design(
        structures, manifest["slide_width"] * manifest["slide_height"])
    masters = sorted({row["master"] for row in manifest["slides"]})
    grouped = [row["number"] for row in manifest["slides"]
               if row.get("grouped_text")]

    out = [f"# {title}", "",
           "The house template this skill builds from when the user names no",
           "template of their own. Generated by `catalog_extract.py` +"
           " `catalog_build.py`",
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

    if grouped:
        out += [f"**Text inside groups.** Designs {', '.join(str(g) for g in grouped)} "
                f"carry text in grouped shapes, which `shapes(page)` cannot "
                f"reach. `retext`, `find_shapes` and `qa.py` all walk groups, "
                f"and `probe.py` lists them indented as `[6.2]`. Each entry "
                f"below notes its count. **Retext by wording, never by index.**",
                ""]

    out += ["## Decision table — content → template slide", "",
            "Apply top to bottom; first match wins.", ""]
    out.append(_decision_table(rules, structures, canvas_number, entries)
               if rules else
               "_No decision table: run the aggregation pass "
               "(`catalog_gemma.py catalog/ --aggregate`) and rebuild._")
    out.append("")
    if variety:
        out += [f"**Variety rule:** {variety}", ""]

    gaps = _gaps(rules) if rules else []
    if gaps:
        out += ["**No design for:** " + ", ".join(gaps)
                + ". Build these on the canvas row above rather than forcing "
                "them into the nearest design — and say so here once you have "
                "checked the renders.", ""]

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

    out += ["## Design entries", "",
            "Each entry's **Retext keys** are the strings the design actually "
            "carries, in `retext`/`find_shapes` form, transcribed from the file "
            "— copy one from here rather than guessing. Any distinctive "
            "substring of a listed string works. `[3.1]` addresses a shape "
            "*inside* the group at `[3]`.", ""]
    for row in manifest["slides"]:
        n = row["number"]
        out += [_entry_block(n, entries[n], structures[n]), ""]

    out += ["---", "",
            "**Provenance.** Structure, retext keys, canvas indices and title "
            "budgets are measured by `catalog_extract.py`; purpose, capacity, "
            "imagery and the decision table are written by a local vision model "
            "from each slide's render and reviewed by hand. The asset, the "
            "renders and this file are brand-internal: commit them to the "
            "internal fork only."]

    if missing_entries:
        out += ["", f"> **Incomplete:** no model entry for design(s) "
                f"{', '.join(str(n) for n in missing_entries)} — their "
                f"judgement fields read _(not catalogued)_. Run the per-layout "
                f"pass for those slides and rebuild."]
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(
        description="Assemble references/company-template.md from a catalog/")
    ap.add_argument("catalog", help="The directory catalog_extract.py wrote")
    ap.add_argument("-o", "--output", help="Write here (default: stdout)")
    ap.add_argument("--asset", default="assets/company-template.pptx",
                    help="Path the skill will load the template from")
    ap.add_argument("--title", default="Company template — layout catalog")
    args = ap.parse_args()

    text = build(Path(args.catalog), args.asset, args.title)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Wrote {args.output} ({len(text.splitlines())} lines)")
        print("Review every capacity figure and decision-table row against the "
              "renders before shipping it.")
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
