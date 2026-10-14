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

sys.path.insert(0, str(Path(__file__).absolute().parent))
from _skill import engine  # noqa: E402

skill = engine()

from office.deckbuild import build_deck  # noqa: E402
from office.impress import mm  # noqa: E402
from office.pack import pack_named, packs  # noqa: E402
from office.template import retext, set_title  # noqa: E402
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


def _is_annotation(text, profile):
    """Whether ``text`` is one of the pack's authoring notes.

    The same test ``strip_annotations`` applies -- whole string, case folded,
    whitespace collapsed, or one of the prefixes -- so a key this says no to is
    a key that survives the strip.
    """
    notes = profile.get("annotations") or {}
    norm = " ".join(str(text).split()).casefold()
    if norm in {" ".join(str(n).split()).casefold()
                for n in notes.get("exact") or ()}:
        return True
    return any(norm.startswith(" ".join(str(p).split()).casefold())
               for p in notes.get("prefixes") or () if str(p).strip())


def _keys_for(pack, design):
    """The retext keys the catalog records for one design, largest type first.

    Authoring notes are left out: ``fill`` strips them before it retexts, so a
    key that names one would match nothing.  The first probe of a template
    whose logo line repeats on every design took that line as the divider's
    title and died on it.  The rest are ordered by the point size of the shape
    that carries them, so the first key is the design's headline rather than
    whichever box happens to come first in z-order.
    """
    path = pack.catalog_dir / f"slide-{design:02d}" / "structure.json"
    try:
        structure = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    profile = pack.profile()
    sizes = {}
    for slot in structure.get("slots") or []:
        biggest = max((font.get("size") or 0
                       for font in slot.get("fonts") or ()), default=0)
        sizes[str(slot.get("index"))] = biggest
    keys = [entry["key"] for entry in structure.get("retext_keys") or []
            if not _is_annotation(entry.get("key", ""), profile)]
    order = {entry["key"]: sizes.get(str(entry.get("address")), 0)
             for entry in structure.get("retext_keys") or []}
    return sorted(keys, key=lambda key: -order.get(key, 0))


def _sized(text, like):
    """``text`` repeated or cut to about the length of ``like``."""
    width = max(len(like), 6)
    out = (text * (width // len(text) + 1))[:width]
    return out.rsplit(" ", 1)[0].strip() if " " in out.strip() else out.strip()


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
    numbers = sorted(int(path.name[len("slide-"):])
                     for path in pack.catalog_dir.glob("slide-[0-9]*")
                     if path.is_dir() and path.name[len("slide-"):].isdigit())
    # A cover has the fewest keys of all -- a title and a page number -- and
    # is no content slide: the probe retexted one and proved nothing.
    covers = set()
    for path in (pack.catalog_dir / "entries").glob("slide-*.json"):
        try:
            entry = json.loads(path.read_text(encoding="utf-8"))
        except ValueError:
            continue
        if "cover" in (entry.get("content_types") or ()):
            covers.add(entry.get("slide"))
    scored = [(len(_keys_for(pack, n)), n) for n in numbers
              if n not in (canvas, divider) and n not in covers
              and _keys_for(pack, n)]
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
        if slot == 0 and divider_keys:
            # The headline gets the probe's title; every other string the
            # design carries gets a same-length stand-in, so what is left on
            # the slide afterwards is the template's decoration and nothing
            # of its wording.
            words = {divider_keys[0]: "The look, checked"}
            words.update({key: _sized("Retexted. ", key)
                          for key in divider_keys[1:]})
            retext(page, words)
        elif slot == 1 and content_keys:
            # Every key, not the first three: the question this slide answers
            # is "is any word still the template's", and a key left alone is a
            # word left alone.  Each replacement is cut to the length of the
            # string it replaces, so a label stays a label and a paragraph a
            # paragraph and nothing overruns its box.
            retext(page, {key: _sized(f"Key {n} came from the catalog. ", key)
                          for n, key in enumerate(content_keys, start=1)})
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
            # The canvas's own title is often an empty placeholder, which
            # retext cannot reach; a probe slide with no headline tests
            # nothing about where the title sits.
            set_title(page, "Your colours, your type scale.")
            area = D.canvas(page, keep=keep)
            # Proportional, not absolute: the free area on a design with a
            # subtitle stripped out can be 50 mm tall, and a 40 mm KPI band
            # cut off the top of that leaves the cards no room at all.  The
            # cards take the smaller type roles for the same reason -- this
            # slide proves the colours, not the capacity.
            top, bottom = area.cut_top(int(area.h * 0.45), gap=mm(4))
            D.kpis(doc, page, top,
                   [("3", "slides"), ("1", "pack"), ("0", "invented colours")],
                   color=D.color("accent"),
                   max_size=min(36.0, D.size_for("display") / 2))
            D.cards(doc, page, bottom,
                    [("Accent", "One element per slide, and this is it."),
                     ("Muted", "Everything that is not the point."),
                     ("Alert", "Kept for the thing that is wrong.")],
                    cols=3, fill=D.color("muted"),
                    accent=[D.color("accent"), D.color("neutral"),
                            D.color("alert")],
                    title_size=D.size_for("caption"),
                    size=D.size_for("footnote"),
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
