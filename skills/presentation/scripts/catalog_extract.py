#!/usr/bin/env python3
"""
Build a per-layout dossier from a company template, ready to judge -- by the
vision-agent, or a person -- into the human catalog.

    python3 scripts/catalog_extract.py company.pptx -o catalog/

This is stage A of turning a brand template into a baked-in default; the whole
pipeline is in README.md.  It runs entirely on your machine, calls no network,
and emits only *structure* -- plus one render per slide.  Nothing here needs
eyes; the vision-agent comes next, reading what this writes.

**Everything eyes cannot be trusted with is measured here.**  The first
version of this pipeline asked the model for the catalog wholesale, and the
first catalog then grew, by hand, the fields that actually decide whether a
build works: the ``retext`` keys, which slots are empty, how much text lives
inside groups, how many characters a title box holds, and which shapes a canvas
keeps.  None of that is judgement -- it is transcription and arithmetic -- so it
is extracted, and the vision-agent is left with what needs eyes.

For each slide of the template it writes, under the output directory:

    catalog/
      manifest.json          one row per slide: number, layout, master, png, slots
      slide-01/
        structure.json       slots, retext keys, canvas hint, title budgets
        render.png           the slide as an image, for the vision-agent
      slide-02/ ...

`name_designs.py` turns each slide-NN/ pair into a worksheet (`designs.md`)
with the measured facts filled and three blanks per design; the vision-agent
fills the blanks from `render.png`, then `name_designs.py --import` and
``catalog_build.py`` assemble the catalog.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from probe import probe_data  # noqa: E402
from office.impress import PRESENTATION_EXTS, SofficeError  # noqa: E402
from office.rasterize import rasterize  # noqa: E402
from office.draw import advance_em, free_area  # noqa: E402

#: How much of a ``retext`` key to record.  A key is a substring, and a short
#: distinctive fragment matches more reliably than a whole line -- but it has to
#: be long enough to be unique on the slide.
KEY_CHARS = 44

#: UNO presentation-shape types: the slots a *design* provides, as opposed to
#: sample content someone dropped on the slide.  Used to guess what a canvas
#: should keep.
CHROME_TYPES = {
    "TitleTextShape", "OutlinerShape", "SubtitleShape",
    "HeaderShape", "FooterShape", "SlideNumberShape", "DateTimeShape",
}

#: A shape this small relative to the slide is a stub or a corner mark, never
#: content.  (One template's think-cell stub is 3 x 3, i.e. 0.000002 of
#: the slide.)  Kept deliberately tight: at 1% it also caught the navigator
#: chevrons, which are sample content a canvas has to drop.
TINY = 0.001

#: Above this share of the slide, a placeholder is the body, not chrome.
BODY_SHARE = 0.20


def _truncate(text, limit=KEY_CHARS):
    """First line of ``text``, cut to ``limit`` -- the form a catalog quotes.

    No trailing ellipsis: ``probe.py`` prints one and every build that pasted
    it into a ``retext`` key matched nothing.  What is written here is exactly
    what can be copied.
    """
    line = " ".join(str(text).split())
    return line[:limit].rstrip()


def _retext_keys(shape_entries):
    """Every string on the slide, with the address that carries it.

    This is the field a deck-builder actually uses, and the one no model can
    supply: it is transcription, not judgement.  Group children keep their
    dotted address (``6.2``), because ``retext`` reaches them and
    ``shapes(page)[i]`` does not.
    """
    keys = []
    for s in shape_entries:
        text = (s.get("text") or "").strip()
        if not text:
            continue
        keys.append({
            "address": s["index"],
            "key": _truncate(text),
            "grouped": "." in str(s["index"]),
        })
    return keys


def _empty_slots(shape_entries):
    """Shapes with no text at all.

    ``retext`` matches on wording, so it can never fill one of these: the
    catalog has to say so, or a build script spends an attempt discovering it.
    """
    return [
        {"address": s["index"], "name": s["name"],
         "placeholder": s.get("placeholder"), "type": s["type"],
         "width": s["width"], "height": s["height"]}
        for s in shape_entries
        if not (s.get("text") or "").strip()
        and (s.get("empty_placeholder") or s["type"] in CHROME_TYPES)
    ]


def _text_counts(shape_entries):
    """How much of this design's text is out of ``shapes(page)``'s reach."""
    top = sum(1 for s in shape_entries
              if (s.get("text") or "").strip() and "." not in str(s["index"]))
    grouped = sum(1 for s in shape_entries
                  if (s.get("text") or "").strip() and "." in str(s["index"]))
    return {"top_level": top, "grouped": grouped}


#: Mixed-case prose of average width, for measuring a box rather than the
#: sample text that happens to sit in it.  Measuring the placeholder's own
#: wording makes the budget depend on whether the designer typed an all-caps
#: word: two divider titles of the same width came out 19 and 11 characters.
REFERENCE_TEXT = "The quick brown fox jumps over the lazy dog"


def _fits_on_one_line(width, size, sample=REFERENCE_TEXT):
    """Characters of average text that fit ``width`` at ``size`` pt.

    The same estimate ``office.draw.fit_size`` inverts, so a title budget
    published in the catalog and a ``shrink_to_fit`` at build time agree.
    """
    if not (width and size):
        return None
    per_char = size * 35.28 * advance_em(sample or REFERENCE_TEXT)
    return int(width / per_char) if per_char else None


def _title_budgets(shape_entries, slide_height):
    """One-line budgets for the title-ish slots.

    A designed title box is drawn for a set number of lines at a set size, and
    master decoration behind it is not a shape on the page -- so a title that
    wraps is struck through by a background triangle and nothing on the slide
    collides.  No check can see that; the number has to be in the catalog.
    """
    out = []
    for s in shape_entries:
        if "." in str(s["index"]) or s["type"] not in CHROME_TYPES:
            continue
        if s["y"] > slide_height * 0.5:          # a footer, not a title
            continue
        sizes = [p["size"] for p in s.get("paragraphs", []) if p.get("size")]
        size = max(sizes) if sizes else None
        chars = _fits_on_one_line(s["width"], size)
        if chars:
            out.append({"address": s["index"], "width": s["width"],
                        "pt": round(size, 1), "chars_per_line": chars})
    return out


def _canvas_hint(shape_entries, slide_width, slide_height):
    """What ``draw.canvas`` would keep, and the free area it would return.

    Mirrors ``canvas``'s own measurement rule, on the geometry rather than on a
    live document: keep the design's own chrome -- the presentation
    placeholders that are not the body, plus anything too small to be content
    -- and drop the rest.  The free box is the union of what was dropped,
    ignoring slivers, which is the body placeholder's own rectangle: the
    designer's answer to where content goes.

    A suggestion, not a fact: run ``probe.py`` on the design before trusting it.
    """
    area = float(slide_width * slide_height) or 1.0
    keep, dropped = [], []
    for s in shape_entries:
        if "." in str(s["index"]):               # group children go with the group
            continue
        share = (s["width"] * s["height"]) / area
        is_chrome = s["type"] in CHROME_TYPES and share < BODY_SHARE
        (keep if (share < TINY or is_chrome) else dropped).append(s)

    def rect(s):
        return (s["x"], s["y"], s["x"] + s["width"], s["y"] + s["height"])

    # The same rule ``draw.canvas`` applies at build time, so the area the
    # catalog publishes is the one a build gets.
    box = free_area([rect(s) for s in dropped], [rect(s) for s in keep],
                    slide_width, slide_height)
    free = None
    if box:
        x0, y0, x1, y1 = box
        free = {"x": x0, "y": y0, "w": x1 - x0, "h": y1 - y0}
    # How much the strip-back demolishes, group contents included: a grouped
    # pyramid is one top-level shape and fourteen on the slide, and counted
    # as one it looked like a single body placeholder to the canvas picker.
    tops = {str(s["index"]) for s in dropped}
    demolished = sum(1 for s in shape_entries
                     if str(s["index"]).split(".")[0] in tops)
    return {
        "keep": [s["index"] for s in keep],
        "drops": [s["index"] for s in dropped],
        "dropped_shapes": demolished,
        "free_box": free,
    }


def _images(shape_entries):
    """Picture shapes, with geometry.  The model says what they *depict*."""
    return [
        {"address": s["index"], "type": s["type"],
         "x": s["x"], "y": s["y"], "width": s["width"], "height": s["height"]}
        for s in shape_entries
        if s["type"] in ("GraphicObjectShape", "MediaShape")
    ]


#: Share of a slot a picture drawn above it must cover to hide it.
COVERED = 0.5


def _covered_by(slot, shape_entries):
    """Index of a picture drawn *above* ``slot`` that hides it, or ``None``.

    Shape order is stacking order.  One template's divider has an empty title
    placeholder under a full-slide picture: the catalog said "set_title", the
    build wrote every section name into it, and three dividers shipped as
    plain blue slides -- the text was in the file and nowhere on the screen.
    """
    index = slot.get("index")
    if index is not None and "." in str(index):
        return None
    area = slot["width"] * slot["height"]
    if area <= 0:
        return None
    x0, y0 = slot["x"], slot["y"]
    x1, y1 = x0 + slot["width"], y0 + slot["height"]
    for s in shape_entries:
        # An inherited layout placeholder (no index) lies under every shape
        # on the slide itself.
        if ("." in str(s["index"])
                or (index is not None and s["index"] <= index)
                or s["type"] not in ("GraphicObjectShape", "MediaShape")):
            continue
        w = min(x1, s["x"] + s["width"]) - max(x0, s["x"])
        h = min(y1, s["y"] + s["height"]) - max(y0, s["y"])
        if w > 0 and h > 0 and w * h >= COVERED * area:
            return s["index"]
    return None


def _slots(shape_entries):
    """Placeholder-ish shapes a layout offers to fill.

    A slot is a text-bearing shape; we keep its index, name, sample text and
    the geometry that tells the model how much it holds.  Purely mechanical --
    the model decides which slots matter.
    """
    slots = []
    for s in shape_entries:
        has_text = bool(s.get("text", "").strip())
        empty_ph = s.get("empty_placeholder", False)
        if not (has_text or empty_ph):
            continue
        fonts = sorted({
            (p.get("font"), round(p["size"], 1) if p.get("size") else None)
            for p in s.get("paragraphs", [])
            if p.get("font") or p.get("size")
        })
        slots.append({
            "index": s["index"],
            "name": s["name"],
            "type": s["type"],
            # A filled slot shows sample_text; an empty inherited placeholder
            # (e.g. an action title) shows its kind so the model still knows the
            # layout offers that slot.
            "placeholder": s.get("placeholder"),
            "empty": empty_ph,
            "x": s["x"], "y": s["y"], "width": s["width"], "height": s["height"],
            "sample_text": s.get("text", ""),
            "fonts": [{"font": f, "size": sz} for f, sz in fonts],
        })
        covered = _covered_by(slots[-1], shape_entries)
        if covered is not None:
            slots[-1]["covered_by"] = covered
    return slots


def extract(source: Path, outdir: Path, dpi: int) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    data = probe_data(source)
    # ``hidden="show"``: a dossier's render has to be the slide it is filed
    # under, and a PDF drops hidden slides, which shifts every later image by
    # one against the numbers probe reports.
    images = rasterize(source, outdir / "_render", dpi=dpi, hidden="show")
    by_number = {}
    for img in images:
        # render.py names files slide-NN.png, 1-based
        stem = img.stem  # slide-01
        try:
            by_number[int(stem.split("-")[-1])] = img
        except ValueError:
            continue

    manifest = {
        "file": source.name,
        "slide_width": data["slide_width"],
        "slide_height": data["slide_height"],
        "slides": [],
    }
    for entry in data["slides"]:
        n = entry["number"]
        slide_dir = outdir / f"slide-{n:02d}"
        slide_dir.mkdir(exist_ok=True)
        slots = _slots(entry["shapes"])
        # Placeholders the layout offers but the slide has no shape for
        # (inherited empty from the master). Without these a slide that only
        # says "Click to add Title" would look slot-less to the model.
        for ls in entry.get("layout_slots", []):
            slots.append({
                "index": None,
                "name": ls.get("name", ""),
                "type": "layout-placeholder",
                "placeholder": ls["kind"],
                "empty": True,
                "inherited": True,
                "x": ls["x"], "y": ls["y"],
                "width": ls["width"], "height": ls["height"],
                "sample_text": "",
                "fonts": [],
            })
            covered = _covered_by(slots[-1], entry["shapes"])
            if covered is not None:
                slots[-1]["covered_by"] = covered
        shapes_ = entry["shapes"]
        structure = {
            "number": n,
            "layout": entry["layout"],
            "master": entry["master"],
            "hidden": entry.get("hidden", False),
            "background": entry["background"],
            "reading_order": entry["reading_order"],
            "slots": slots,
            "shape_count": len(shapes_),
            # Everything below is transcription, not judgement: it is what the
            # finished catalog needs and what a model was previously asked for
            # and got wrong.  The model never sees these as questions.
            "retext_keys": _retext_keys(shapes_),
            "empty_slots": _empty_slots(shapes_),
            "text_counts": _text_counts(shapes_),
            "title_budgets": _title_budgets(shapes_, data["slide_height"]),
            "canvas": _canvas_hint(shapes_, data["slide_width"],
                                   data["slide_height"]),
            "images": _images(shapes_),
        }
        (slide_dir / "structure.json").write_text(
            json.dumps(structure, indent=2, ensure_ascii=False)
        )
        png_rel = None
        if n in by_number:
            dest = slide_dir / "render.png"
            shutil.copyfile(by_number[n], dest)
            png_rel = str(dest.relative_to(outdir))
        manifest["slides"].append({
            "number": n,
            "layout": entry["layout"],
            "master": entry["master"],
            "hidden": entry.get("hidden", False),
            "render": png_rel,
            "structure": str((slide_dir / "structure.json").relative_to(outdir)),
            "slot_count": len(slots),
            "grouped_text": structure["text_counts"]["grouped"],
        })

    shutil.rmtree(outdir / "_render", ignore_errors=True)
    (outdir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False)
    )
    return manifest


def main():
    ap = argparse.ArgumentParser(
        description="Extract per-layout dossiers from a template for judging"
    )
    ap.add_argument("input", help="Template .pptx/.potx/.odp/.otp to survey")
    ap.add_argument("-o", "--outdir", default=None,
                    help="Output directory (default: <template>-catalog/ "
                         "beside the template)")
    ap.add_argument("--dpi", type=int, default=150, help="Render resolution")
    args = ap.parse_args()

    source = Path(args.input)
    if not source.is_file():
        ap.error(f"{source} does not exist")
    if source.suffix.lower() not in PRESENTATION_EXTS:
        ap.error(f"not a presentation: {source}")

    if args.outdir is None:
        args.outdir = str(source.resolve().with_name(source.stem + "-catalog"))
    try:
        manifest = extract(source, Path(args.outdir), args.dpi)
    except SofficeError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2
    n = len(manifest["slides"])
    masters = sorted({s["master"] for s in manifest["slides"]})
    print(f"Wrote {n} layout dossiers to {args.outdir}/")
    print(f"Masters seen: {', '.join(masters)}")
    print(f"Next: vision_brief.py designs {args.outdir} -- one brief per "
          f"design for the vision-agent -- then name_designs.py "
          f"{args.outdir} for the worksheet, then name_designs.py "
          f"{args.outdir} --import")
    return 0


if __name__ == "__main__":
    sys.exit(main())
