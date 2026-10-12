#!/usr/bin/env python3
"""
Assemble a layout catalog from the extraction and the model pass -- stage C.

    python3 scripts/catalog_build.py templates/acme/catalog/ \
        -o templates/acme/catalog.md --asset templates/acme/template.pptx

Reads what ``catalog_extract.py`` measured (``manifest.json`` and each
``slide-NN/structure.json``) and what was judged
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
from office.draw import contrast_ratio, free_area  # noqa: E402
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
    ("a single headline idea", ("headline", "one big", "one idea", "single idea", "statement")),
    ("bulleted content", ("bullet", "prose", "body text", "standard content")),
    ("a comparison", ("compar", "versus", "vs", "two things", "before/after")),
    ("a process or timeline", ("process", "timeline", "roadmap", "phase", "step")),
    ("numbers or KPIs", ("kpi", "number", "metric", "financial", "cost")),
    # No bare "data": a row about "customer data" is not a chart design, and
    # matching it hides the gap.  "chart" and "graph" name the thing itself.
    ("a chart", ("chart", "graph", "plot")),
    ("a quote", ("quote", "testimonial")),
    # "person" would match "personal"; the words a real rule uses are these.
    ("people", ("people", "team", "org chart", "personnel", "headshot",
                "who is involved")),
    ("a closing slide", ("closing", "close", "final", "thank", "end", "last slide")),
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
    items = (f" · holds {cap['max_items']} item(s)"
             if cap.get("max_items") else "")
    return (f"headline {tick} · body slots {cap.get('body_slots', '?')}{items} · "
            + (f"max bullets {cap['max_bullets_per_body']} · "
               if cap.get("max_bullets_per_body") else "")
            + f"images {cap.get('images', '?')} · chart {chart}")


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
    out = [f"# Design {number} — {name}"]
    out.append(f"- **Purpose:** {entry.get('purpose', '_(not catalogued)_')}")
    types = entry.get("content_types") or []
    out.append("- **Content types:** "
               + (", ".join(f"`{t}`" for t in types) or "_(none)_"))
    if entry.get("content_shape"):
        out.append(f"- **Content shape:** {entry['content_shape']}")
    out.append(f"- **Capacity:** {_capacity(entry.get('capacity'))}")
    out.append(f"- **Headline:** {_headline(structure)}")

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
    if entry.get("looks_best"):
        out.append(f"- **Looks best:** {entry['looks_best']}")
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

#: The least free area, as a share of the slide, a canvas has to offer.  Below
#: it a design does not qualify, and among the unqualified ones plainness
#: only outranks size above it.
MIN_CANVAS_SHARE = 0.35


def _canvas_design(structures, slide_area, exclude=()):
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
    # The canvas is a content slide, so it comes from the master most designs
    # are built on.  The largest free area alone picked a cover once: a photo
    # under the title bar is a big rectangle to drop, and every content slide
    # of the deck after it was a title slide with no footer and no page
    # number.  Only when the content master offers no candidate at all is
    # the field widened to every master.
    # Nor is it a design judged to open a deck or a section: a divider that
    # strips back to its big number is roomy and plain, and still a divider.
    if exclude:
        pick = _canvas_design({n: s for n, s in structures.items()
                               if n not in set(exclude)}, slide_area)
        if pick[0]:
            return pick
    counts = {}
    for s in structures.values():
        counts[s.get("master")] = counts.get(s.get("master"), 0) + 1
    content = max(counts, key=counts.get) if counts else None
    pick = _canvas_pick({n: s for n, s in structures.items()
                         if s.get("master") == content}, slide_area)
    # Even unqualified, the content master's least-bad design beats a clean
    # strip-back of a cover: it is printed with a warning, the cover is not.
    return pick if pick[0] else _canvas_pick(structures, slide_area)


def _canvas_pick(structures, slide_area):
    """``(number, qualified)`` over ``structures``: see :func:`_canvas_design`."""
    best, best_area, fallback, fallback_key = None, 0, None, None
    for number, s in structures.items():
        canvas = s.get("canvas") or {}
        box = canvas.get("free_box")
        if not box:
            continue
        area = box["w"] * box["h"]
        if area > slide_area * MAX_CANVAS_SHARE:
            continue                      # full bleed: no chrome to inherit
        keep = set(canvas.get("keep") or [])
        if any(_is_title(slot) and slot.get("covered_by") is not None
               and slot.get("index") in keep for slot in s.get("slots") or []):
            continue                      # its title is under a picture
        # Two conditions, and the second is the one that matters: the free area
        # has to come from *one* body placeholder, not from clearing a slide
        # full of sample content.  A design whose canvas drops nine shapes is
        # an org chart being demolished; a design whose canvas drops one is the
        # content slide the designer left room on.  Counted with group
        # contents, where the extractor recorded them.
        # And it has to leave room: a question slide that drops its picture
        # frees the right half and nothing else.
        demolished = canvas.get("dropped_shapes", len(canvas.get("drops", [])))
        keeps_chrome = len(keep) >= 2
        roomy = area >= slide_area * MIN_CANVAS_SHARE
        if keeps_chrome and roomy and demolished <= 2 and area > best_area:
            best, best_area = number, area
        # The fallback is the plainest design with room to draw, not the
        # largest: largest-first handed one template a slide of thirty sample
        # shapes to tear down.  One that keeps nothing of its own chrome
        # comes after one that does; below the room threshold, area decides.
        key = (roomy, keeps_chrome, -demolished if roomy else 0, area)
        if fallback_key is None or key > fallback_key:
            fallback, fallback_key = number, key
    return (best, True) if best else (fallback, False)


def _canvas_override(aggregate, structures, slide_width, slide_height):
    """``(number, qualified)`` from a ``canvas`` block a person wrote.

    ``decision_table.json`` may carry ``{"canvas": {"design": N, "keep": [..]}}``
    -- written by hand, or by ``name_designs.py --import`` from the worksheet.
    It is the one judged fact about the canvas: the heuristic in
    :func:`_canvas_design` picks the largest free area, and on a template
    whose designs all overhang or all keep too little it picks badly and says
    so with a ⚠.  The probe deck shows that in ten seconds; this is how the
    answer survives a rebuild.  With ``keep`` given, the design's own
    ``canvas`` hint is rewritten around it: the free box is the union of the
    text-bearing top-level shapes that are not kept, clamped to the slide.
    """
    block = aggregate.get("canvas") if isinstance(aggregate, dict) else None
    if not isinstance(block, dict):
        return None, False
    try:
        number = int(block.get("design"))
    except (TypeError, ValueError):
        return None, False
    if number not in structures:
        print(f"warning: decision_table.json names canvas design {number}, "
              f"which this template does not have -- ignored", file=sys.stderr)
        return None, False
    keep = block.get("keep")
    if isinstance(keep, (list, tuple)) and all(isinstance(k, int) for k in keep):
        keep = sorted(set(keep))
        slots = [s for s in structures[number].get("slots") or []
                 if isinstance(s.get("index"), int)]
        dropped = [s for s in slots if s["index"] not in keep]
        kept = [s for s in slots if s["index"] in keep]

        def rect(s):
            return (s["x"], s["y"], s["x"] + s["width"], s["y"] + s["height"])

        old = structures[number].get("canvas") or {}
        free = old.get("free_box")
        box = free_area([rect(s) for s in dropped], [rect(s) for s in kept],
                        slide_width, slide_height)
        if box:
            x0, y0, x1, y1 = box
            free = {"x": x0, "y": y0, "w": x1 - x0, "h": y1 - y0}
        structures[number]["canvas"] = {
            "keep": keep,
            "drops": [s["index"] for s in dropped],
            "free_box": free,
            "by_hand": True,
        }
    if not (structures[number].get("canvas") or {}).get("free_box"):
        print(f"warning: canvas design {number} has no measurable free area; "
              f"give it a keep= list in decision_table.json", file=sys.stderr)
        return None, False
    return number, True


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
           f"D.color(\"{next(iter(colours), 'accent')}\")   "
           "D.color(\"accent\")     # by name, or by role",
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

#: A decision-table row about the deck's first slide.  Judged rows are worded
#: freely: "Opening slide of the presentation" missed the old pattern, and the
#: catalog told every build there was no cover while two designs were covers.
COVER_ROW = re.compile(
    r"\bopener\b|\bopening (?:slide|of the)\b|\bfirst slide\b"
    r"|\btitle slide\b|\bcover\b|\bstart of the (?:deck|presentation)\b"
    r"|\bdeckblatt\b"
    r"|\beröffnung", re.I)


def _role_numbers(rules, entries):
    """``(dividers, cover)``: the designs judged to open a deck or a section."""
    dividers = sorted(
        number for number, entry in entries.items()
        if any(DIVIDER_TYPES.search(str(kind))
               for kind in entry.get("content_types", ()) or ())
    )
    cover = None
    for rule in rules:
        when = str(rule.get("when") or rule.get("content") or "")
        if COVER_ROW.search(when):
            # Rows are keyed ``use_slide`` by every writer -- the judging
            # pass, the worksheet import and the recovery script.  Reading
            # ``design`` here, which none of them writes, reported "no cover
            # design" on every template that had one.
            try:
                cover = int(str(rule.get("use_slide", rule.get("design")))
                            .strip("`"))
            except (TypeError, ValueError):
                cover = None
            break
    return dividers, cover


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
    dividers, cover = _role_numbers(rules, entries)

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


#: Where the per-design files go, beside ``catalog.md``.
DESIGNS_DIR = "designs"


def _key_at(structure, address):
    """The retext key the shape at ``address`` carries, or ``None``."""
    for key in structure.get("retext_keys") or []:
        if str(key.get("address")) == str(address):
            return key.get("key")
    return None


def _headline(structure):
    """How a build writes this design's headline, in one cell.

    Measured, never judged.  An empty title placeholder is the trap: ``retext``
    cannot reach an empty shape, so a build writes its headline into whatever
    small label does carry text, and the slide ships with no title.
    """
    all_slots = structure.get("slots") or []
    # A slot under a picture is not on the screen, whatever it holds: one
    # divider's title placeholder sat under a full-slide image, and "the
    # placeholder is empty -- set_title" sent every section name into it.
    slots = [s for s in all_slots if s.get("covered_by") is None]
    hidden = next((s for s in all_slots if s.get("covered_by") is not None
                   and _is_title(s)), None)
    # The title is the title shape, filled or not.  ``placeholder`` is only
    # set for an empty one, so a filled title used to be passed over for the
    # largest-type budget -- a 32pt subtitle in the corner, on one cover.
    title = next((s for s in slots if _is_title(s)), None)
    note = (f" — the title placeholder `[{hidden['index']}]` is hidden under "
            f"picture `[{hidden['covered_by']}]`: never `set_title` here"
            if hidden is not None and title is None else "")
    if title is not None and title.get("empty"):
        return "`set_title(page, …)` — the placeholder is empty"
    if title is not None:
        key = _key_at(structure, title.get("index"))
        if key:
            return f'retext `"{key[:30]}"`'
    visible = {s.get("index") for s in slots}
    budgets = [b for b in structure.get("title_budgets") or []
               if isinstance(b, dict) and b.get("pt", 0) >= DISPLAY_PT
               and b.get("address") in visible]
    if budgets:
        b = budgets[0]
        key = _key_at(structure, b.get("address"))
        where = f'retext `"{key[:30]}"`' if key else f"shape `[{b['address']}]`"
        return f"{where}, one line of ≤{b['chars_per_line']} characters{note}"
    # No title placeholder: the headline is the largest text in the top third,
    # which is what a designer's eye -- and the audience's -- reads first.
    # A section number ("01") is large too, and is never the headline.
    words = [s for s in slots
             if len((s.get("sample_text") or "").strip()) >= 4
             and re.search(r"[^\W\d_]", s.get("sample_text") or "")]
    third = (structure.get("slide_height") or 0) / 3
    # A divider centres its title, so the whole slide is the fallback.
    tops = [s for s in words if s.get("y", 1 << 30) < third] or words
    if tops:
        biggest = max(tops, key=lambda s: max(
            (f.get("size") or 0 for f in s.get("fonts") or [{}]), default=0))
        key = _key_at(structure, biggest.get("index"))
        if key:
            return f'retext `"{key[:30]}"` (the largest text){note}'
    return "—" + note


#: How many unnamed designs ``catalog.md`` spells out before it points at the
#: index: on a sixty-design template the whole list buries the table.
UNNAMED_SHOWN = 15


def _unnamed_designs(rules, entries, canvas_number):
    """``[(number, name)]`` of judged designs no decision-table row uses."""
    used = set()
    for rule in rules:
        try:
            used.add(int(str(rule.get("use_slide")).strip("`")))
        except (TypeError, ValueError):
            continue
    return [(n, entry.get("name") or f"Design {n}")
            for n, entry in sorted(entries.items())
            if entry and n not in used and n != canvas_number]


def _is_title(slot):
    """A title slot, whether the template left it empty or filled it."""
    return (slot.get("placeholder") == "title"
            or slot.get("type") == "TitleTextShape")


def _canvas_leftovers(structure):
    """Kept canvas shapes that still carry the template's sample text.

    The canvas keeps the title and a topic label.  A topic label that carries
    sample wording ships it on every drawn slide unless the build retexts it --
    one pack's canvas put its own design name on nine slides of nine.
    """
    keep = set((structure.get("canvas") or {}).get("keep") or [])
    out = []
    for slot in structure.get("slots") or []:
        if (slot.get("index") in keep and slot.get("sample_text")
                and slot.get("placeholder") != "title"):
            out.append((slot["index"], slot["sample_text"].strip()))
    return out


def _index_table(manifest, structures, entries, canvas_number):
    """One row per design: what it holds and how its headline is written."""
    rows = ["| Design | Name | Content shape | Headline |", "|---|---|---|---|"]
    for row in manifest["slides"]:
        n = row["number"]
        entry, structure = entries[n], structures[n]
        name = entry.get("name") or f"Design {n}"
        if n == canvas_number:
            name += " — **canvas**"
        shape = entry.get("content_shape") or entry.get("use_when") or "—"
        items = (entry.get("capacity") or {}).get("max_items")
        if items and "item" not in shape:
            shape += f" (holds {items})"
        rows.append(f"| `{n}` | {name} | {shape} | {_headline(structure)} |")
    return rows


def build(catalog: Path, asset: str, title: str, overrides=None):
    """``(catalog_md, {"design-N.md": text})`` for a pack.

    ``catalog.md`` is what a build reads first and in full, so it carries
    only what picks a design: the roles, the decision table, one index row per
    design and the palette.  Everything needed to *fill* a design -- its retext
    keys, empty slots, budgets -- is in ``designs/design-N.md``, read for the
    designs the build actually uses.  One file of every key for sixty designs
    ran to sixteen thousand tokens, and a small model reading it lost the
    decision table it came for.
    """
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
        structures[n].setdefault("slide_height", manifest["slide_height"])
        entries[n] = _load(catalog / "entries" / f"slide-{n:02d}.json", {}) or {}

    aggregate = _load(catalog / "decision_table.json", {}) or {}
    rules = aggregate.get("decision_table") or []
    variety = aggregate.get("variety_rule")

    missing_entries = [n for n, e in entries.items() if not e]
    canvas_number, canvas_ok = _canvas_override(
        aggregate, structures, manifest["slide_width"], manifest["slide_height"])
    if canvas_number is None:
        dividers, cover = _role_numbers(rules, entries)
        canvas_number, canvas_ok = _canvas_design(
            structures, manifest["slide_width"] * manifest["slide_height"],
            exclude=set(dividers) | {cover})
    masters = sorted({row["master"] for row in manifest["slides"]})
    judged_by = next((entry.get("written_by") for entry in entries.values()
                      if entry.get("written_by")), None)

    out = [f"# {title}", "",
           f"**Asset:** `{asset}` — {len(manifest['slides'])} designs, "
           f"{manifest['slide_width']} × {manifest['slide_height']} (1/100 mm), "
           "masters " + ", ".join(f"`{m}`" for m in masters) + ".", "",
           "**How to use this file.**", "",
           "1. Pick a design per slide: the decision table first, then the "
           "index below by *content shape*. Vary them — a deck on one design "
           "is a form someone filled in.",
           "2. A design number is a slide number in the asset; it is what "
           "`build_deck`'s list takes.",
           f"3. **Before writing `fill` for design N, read "
           f"`{DESIGNS_DIR}/design-N.md`** — its retext keys, its empty "
           "slots, its budgets. Never reuse one design's keys on another.",
           "4. Write every headline the way the index's *Headline* column says.",
           "5. Anything no design fits goes on the canvas design, stripped "
           "back.", ""]

    out += _roles_block(canvas_number, rules, entries)
    if "header" in blocks:
        out += [blocks.pop("header"), ""]

    out += ["## Decision table", "",
            "Top to bottom, first match wins. The canvas row is last on "
            "purpose: it is for content nothing above fits.", ""]
    out.append(_decision_table(rules, structures, canvas_number, entries,
                               qualified=canvas_ok)
               if rules else
               "_No decision table: fill the worksheet and import it "
               "(`name_designs.py catalog/ --import`) and rebuild._")
    out.append("")
    leftovers = _canvas_leftovers(structures[canvas_number]) if canvas_number else []
    if leftovers:
        out += ["**The canvas keeps sample text.** Retext it on every canvas "
                "slide, or it ships: "
                + ", ".join(f'`[{i}]` "{t[:40]}"' for i, t in leftovers), ""]
    unnamed = _unnamed_designs(rules, entries, canvas_number) if rules else []
    if unnamed:
        # A small builder reads the table and stops there.  One pack's table
        # had eleven rows for thirty designs, and a workshop deck went 83% to
        # the canvas while a two-column design, a status table and a six-icon
        # grid that fitted its slides were never named anywhere it looked.
        shown = ", ".join(f"`{n}` {name}" for n, name in unnamed[:UNNAMED_SHOWN])
        if len(unnamed) > UNNAMED_SHOWN:
            shown += f", and {len(unnamed) - UNNAMED_SHOWN} more in the index"
        out += ["**No row names these designs** — check their *Content "
                "shape* in the index before you draw on the canvas: "
                + shown, ""]
    if variety:
        out += [f"**Variety:** {variety}", ""]
    gaps = _gaps(rules) if rules else []
    if gaps:
        out += ["**No design for:** " + ", ".join(gaps)
                + " — use the canvas design.", ""]
    if "after-decision-table" in blocks:
        out += [blocks.pop("after-decision-table"), ""]

    out += ["## Design index", ""]
    if "before-design-entries" in blocks:
        out += [blocks.pop("before-design-entries"), ""]
    out += _index_table(manifest, structures, entries, canvas_number) + [""]

    out += _palette_block(asset)
    if "footer" in blocks:
        out += [blocks.pop("footer"), ""]
    out += ["---", "",
            "Generated by `catalog_build.py`: structure, keys, budgets and the "
            "canvas are measured; names, content shapes and the decision table "
            "were judged by " + (judged_by or "the vision-agent")
            + " and reviewed. Edit `overrides.md`, never this file."]
    if missing_entries:
        out += ["", f"> **Incomplete:** no entry for design(s) "
                f"{', '.join(str(n) for n in missing_entries)}. Judge them in "
                f"the worksheet and rebuild."]

    details = {}
    intro = blocks.pop("design-entries-intro", None)
    for row in manifest["slides"]:
        n = row["number"]
        structure = structures[n]
        lines = [_entry_block(n, entries[n], structure,
                              blocks.pop(f"design-{n}", None)), ""]
        if n == canvas_number:
            box = structure["canvas"]["free_box"]
            keep = ", ".join(str(k) for k in structure["canvas"]["keep"])
            lines += ["**This is the canvas design.** "
                      f"`area = D.canvas(page, keep=({keep}))` strips it back "
                      f"and returns the free area ({box['x']}, {box['y']} to "
                      f"{box['x'] + box['w']}, {box['y'] + box['h']}); draw "
                      "inside it.", ""]
            for i, t in _canvas_leftovers(structure):
                lines.append(f'- Kept shape `[{i}]` still says "{t[:50]}" — '
                             f"retext it.")
            lines.append("")
        lines.append(("`[3.1]` is a shape inside the group at `[3]`. "
                      if (structure.get("text_counts") or {}).get("grouped")
                      else "")
                     + "Any distinctive substring of a key works; copy it "
                       "from here, not from memory. Shapes with identical "
                       "text: use the address as the key, e.g. `\"[3.1]\"`.")
        if intro:
            lines += ["", intro]
        details[f"design-{n}.md"] = "\n".join(lines) + "\n"

    if blocks:
        for anchor in sorted(blocks):
            print(f"WARNING: override block '{anchor}' matched no anchor in "
                  f"this catalog", file=sys.stderr)
        out += ["", "> ⚠ **Unplaced overrides:** "
                + ", ".join(f"`{a}`" for a in sorted(blocks))
                + f". Valid anchors are {', '.join(f'`{a}`' for a in ANCHORS)} "
                + "and `design-N` for a design this template has."]
    return "\n".join(out) + "\n", details


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
    ap.add_argument("-o", "--output",
                    help="Write here. Default: the pack's catalog.md when the "
                         "catalog directory sits inside a pack, else stdout")
    ap.add_argument("--asset", default=None,
                    help="Path the skill will load the template from")
    ap.add_argument("--title", default="Company template — layout catalog")
    ap.add_argument("--overrides",
                    help="Markdown file of '## anchor' prose blocks spliced "
                         "into the generated catalog. Default: the pack's "
                         "overrides.md if there is one. Anchors: "
                         + ", ".join(ANCHORS) + ", design-N")
    ap.add_argument("--check", action="store_true",
                    help="Do not write: rebuild and diff against --output, "
                         "exit 1 if it has drifted. Re-extract first, so this "
                         "compares the shipped catalog with the template as it "
                         "is now.")
    args = ap.parse_args()

    catalog = Path(args.catalog)
    if not (catalog / "manifest.json").is_file():
        ap.error(f"no manifest.json in {catalog} -- point this at the "
                 f"directory catalog_extract.py wrote")
    asset = args.asset or _asset_beside(catalog)
    if asset is None:
        ap.error("--asset: no template found beside %s. Pass the path to the "
                 "template this catalog was extracted from." % args.catalog)

    # Inside a pack the other two files have fixed names, so a bare
    # ``catalog_build.py templates/acme/catalog`` rebuilds the pack's own
    # catalog.md with its own overrides.md and needs no flags at all.
    in_pack = catalog.resolve().parent
    if args.overrides is None and (in_pack / "overrides.md").is_file():
        args.overrides = str(in_pack / "overrides.md")
    if args.overrides and not Path(args.overrides).is_file():
        ap.error(f"--overrides: {args.overrides} does not exist")
    if args.output is None and _asset_beside(catalog) is not None:
        args.output = str(in_pack / "catalog.md")

    text, details = build(catalog, asset, args.title, args.overrides)
    files = {}
    if args.output:
        out = Path(args.output)
        files[out] = text
        files.update({out.parent / DESIGNS_DIR / name: body
                      for name, body in details.items()})

    if args.check:
        if not args.output:
            ap.error("--check needs --output: the catalog to compare against")
        import difflib
        drifted = False
        for path, body in files.items():
            current = path.read_text(encoding="utf-8") if path.is_file() else ""
            if current != body:
                drifted = True
                sys.stdout.writelines(difflib.unified_diff(
                    current.splitlines(keepends=True),
                    body.splitlines(keepends=True),
                    fromfile=f"{path} (shipped)", tofile="rebuilt", n=1))
        if drifted:
            print(f"\n{args.output} has drifted from the template. Rebuild it, "
                  f"or move whatever is only in the shipped files into "
                  f"--overrides.", file=sys.stderr)
            return 1
        print(f"{args.output} is up to date with {args.catalog}")
        return 0

    if not args.output:
        sys.stdout.write(text)
        for name, body in details.items():
            sys.stdout.write(f"\n<!-- {DESIGNS_DIR}/{name} -->\n{body}")
        return 0
    stale = Path(args.output).parent / DESIGNS_DIR
    if stale.is_dir():
        for old in stale.glob("design-*.md"):
            if old not in files:
                old.unlink()            # a design the template no longer has
    for path, body in files.items():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(body, encoding="utf-8")
    print(f"Wrote {args.output} ({len(text.split())} words) and "
          f"{len(details)} design file(s) in {stale}/")
    print("Review every content shape and decision-table row against the "
          "renders before shipping it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
