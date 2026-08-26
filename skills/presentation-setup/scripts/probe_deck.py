#!/usr/bin/env python3
"""Build three real slides on a new pack, so somebody can look at them.

    python3 scripts/probe_deck.py acme

The last question of a setup run is never "is the catalog correct". It is
**"will this build slides that look like ours"**, and a 500-line catalog cannot
answer it. Three rendered slides can, in about ten seconds of looking:

1. a divider, which proves the title budget and the display size;
2. a content slide retexted through the catalog's own keys, which proves the
   keys are real and that `strip_annotations` took the authoring notes out;
3. a drawn slide on the canvas design -- cards and a KPI row in the pack's own
   colours -- which proves the palette, the type scale and the canvas box.

If the pack is wrong, this is where it shows: a title that wraps into the
master's decoration, a label the retext never replaced, a card in a colour
nobody recognises. Every one of those is a five-minute fix here and a shipped
deck later.

The deck is written to the working directory, not into the pack. It is
evidence, not an artifact.
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine  # noqa: E402

skill = engine()

from office.deckbuild import build_deck  # noqa: E402
from office.impress import mm  # noqa: E402
from office.pack import pack_named, packs  # noqa: E402
from office.template import retext, strip_annotations  # noqa: E402
from office import draw as D  # noqa: E402


def _catalog_roles(pack):
    """``(canvas, keep, dividers)`` from the pack's own catalog.

    Read out of `decision_table.json` and the entries rather than guessed:
    this script must not be the second place in the skill that decides what a
    canvas design is.
    """
    canvas, keep, dividers = None, None, []
    entries = sorted((pack.catalog_dir / "entries").glob("slide-*.json")) \
        if (pack.catalog_dir / "entries").is_dir() else []
    for path in entries:
        try:
            entry = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        if any("divider" in str(kind) for kind in entry.get("content_types", ())):
            dividers.append(entry["slide"])
    text = pack.catalog.read_text(encoding="utf-8") if pack.catalog.is_file() else ""
    for line in text.splitlines():
        if line.startswith("| **canvas**"):
            digits = "".join(ch for ch in line.split("|")[2] if ch.isdigit())
            canvas = int(digits) if digits else None
    found = re.search(r"keep=\(([0-9,\s]*)\)", text)
    if found:
        keep = tuple(int(part) for part in found.group(1).split(",")
                     if part.strip())
    return canvas, keep, dividers


def _keys_for(pack, design):
    """The retext keys the catalog records for one design."""
    path = pack.catalog_dir / f"slide-{design:02d}" / "structure.json"
    try:
        structure = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return [entry["key"] for entry in structure.get("retext_keys") or []]


def _content_design(pack, canvas, divider):
    """The design to probe the retext keys on: the plainest content slide.

    The catalog's own row for body text if it has one, else the design with the
    **fewest** retext keys. Taking the first design with any keys picks a
    process diagram or an org chart on most templates -- 20 grouped labels the
    probe was never going to retext -- and every one of them reports as
    leftover wording, which reads as the pack being broken when it is the
    probe that chose badly.
    """
    table = pack.catalog_dir / "decision_table.json"
    try:
        rules = json.loads(table.read_text(encoding="utf-8"))["decision_table"]
    except (OSError, ValueError, KeyError, TypeError):
        rules = []
    for rule in rules:
        when = str(rule.get("when", "")).lower()
        if "bullet" in when or "body text" in when:
            try:
                number = int(rule.get("use_slide") or rule.get("design"))
            except (TypeError, ValueError):
                continue
            if number not in (canvas, divider) and _keys_for(pack, number):
                return number
    scored = [(len(_keys_for(pack, n)), n) for n in range(1, 60)
              if n not in (canvas, divider) and _keys_for(pack, n)]
    return min(scored)[1] if scored else canvas


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("pack", help="The pack name, e.g. acme")
    ap.add_argument("-o", "--out", default="probe.pptx")
    args = ap.parse_args()

    pack = pack_named(args.pack)
    if pack is None:
        raise SystemExit(
            f"no pack called {args.pack!r}. Installed: "
            + (", ".join(p.name for p in packs()) or "none"))

    canvas, keep, dividers = _catalog_roles(pack)
    if canvas is None:
        raise SystemExit(
            f"{pack.catalog} has no canvas design in its Designs by role "
            f"table — build the catalog first (catalog_build.py), or this "
            f"probe would be inventing the one thing it is meant to check.")
    divider = dividers[0] if dividers else canvas
    content = _content_design(pack, canvas, divider)

    divider_keys = _keys_for(pack, divider)
    content_keys = _keys_for(pack, content)

    def fill(doc, page, slot, design):
        strip_annotations(page)
        if slot == 0 and divider_keys:
            retext(page, {divider_keys[0]: "The look, checked"})
        elif slot == 1 and content_keys:
            retext(page, {key: value for key, value in zip(
                content_keys,
                ["Every key in this slide came from the catalog.",
                 "If a word below is still the template's, the keys are wrong.",
                 "Third slot, if this design has one."])})
        elif slot == 2:
            # Retext what the canvas keeps before drawing in what it drops:
            # otherwise the probe reports the template's own title as leftover
            # wording on every pack, which reads as the pack being broken.
            canvas_keys = _keys_for(pack, canvas)
            if canvas_keys:
                retext(page, {canvas_keys[0]: "Your colours, your type scale."})
                for key in canvas_keys[1:]:
                    retext(page, {key: "Counted off your template — "
                                       "nothing here was invented."})
            area = D.canvas(page, keep=keep)
            top, bottom = area.cut_top(mm(40), gap=mm(8))
            D.kpis(doc, page, top,
                   [("3", "slides"), ("1", "pack"), ("0", "invented colours")],
                   color=D.color("accent"))
            D.cards(doc, page, bottom,
                    [("Accent", "One element per slide, and this is it."),
                     ("Muted", "Everything that is not the point."),
                     ("Alert", "Kept for the thing that is wrong.")],
                    fill=D.color("muted"),
                    accent=[D.color("accent"), D.color("neutral"),
                            D.color("alert")],
                    height="content")

    if keep is None:
        raise SystemExit(
            f"{pack.catalog} records no canvas keep tuple. Rebuild it with "
            f"catalog_build.py — the probe reads the tuple the same way a "
            f"build script does, and inventing one here would test nothing.")

    print(f"pack {pack.name}: divider {divider}, content {content}, "
          f"canvas {canvas} keep={keep}")
    build_deck(pack.name, args.out, [divider, content, canvas], fill=fill,
               footer=f"{pack.name} · pack probe")
    print(f"\nWrote {args.out}. Render it and look:\n"
          f"  python3 {skill}/scripts/render.py {args.out}\n\n"
          f"Three questions, in order: does the divider title fit on one line; "
          f"is any word on slide 2 still the template's; are the colours on "
          f"slide 3 the ones your brand actually uses?")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
