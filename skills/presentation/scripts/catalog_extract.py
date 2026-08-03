#!/usr/bin/env python3
"""
Build a per-layout dossier from a company template, ready to hand to a local
vision model (e.g. gemma) that writes the human catalog.

    python3 scripts/catalog_extract.py company.pptx -o catalog/

This is stage A of turning a brand template into a baked-in default (see
"Company template" in SKILL.md).  It runs entirely on your machine, calls no
network, and emits only *structure* -- layout numbers, placeholder slots,
geometry, fonts, colours -- plus one render per slide.  Nothing here needs a
model; the model comes next, reading what this writes.

For each slide of the template it writes, under the output directory:

    catalog/
      manifest.json          one row per slide: number, layout, master, png, slots
      slide-01/
        structure.json       the probe entry, trimmed to what a catalog needs
        render.png           the slide as an image, for the vision model
      slide-02/ ...

Feed each slide-NN/ pair (render.png + structure.json) to the model together
with references/gemma-catalog-prompt.md.  The model's job is the judgement this
script deliberately does not attempt: what each layout is *for*.
"""

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from probe import probe_data  # noqa: E402
from render import render  # noqa: E402
from office.impress import (  # noqa: E402
    PRESENTATION_EXTS, load_document, save_as, slides,
)


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
    return slots


def _all_visible_copy(source: Path, dest: Path):
    """Write a copy of ``source`` at ``dest`` with every slide made visible.

    render.py renders through PDF, which drops hidden slides -- so a deck with a
    hidden slide produces fewer PNGs than it has slides, and every PNG after the
    first hidden one is off by one against the slide numbers probe reports.
    Rendering an all-visible copy keeps one PNG per slide, in order.

    The source is copied first and only the copy is opened read-write: the
    original template is often installed read-only, and opening a non-writable
    file read-write fails obscurely as "type detection failed".
    """
    shutil.copyfile(source, dest)
    doc = load_document(dest, read_only=False)
    try:
        for page in slides(doc):
            if not getattr(page, "Visible", True):
                page.Visible = True
        save_as(doc, dest)
    finally:
        doc.close(False)


def extract(source: Path, outdir: Path, dpi: int) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    data = probe_data(source)
    hidden_any = any(s.get("hidden") for s in data["slides"])
    render_source = source
    tmp = None
    if hidden_any:
        # Render from an all-visible copy so PNG numbers line up with slides.
        tmp = Path(tempfile.mkdtemp()) / f"visible{source.suffix}"
        _all_visible_copy(source, tmp)
        render_source = tmp
    images = render(render_source, outdir / "_render", dpi, None)
    if tmp is not None:
        shutil.rmtree(tmp.parent, ignore_errors=True)
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
        structure = {
            "number": n,
            "layout": entry["layout"],
            "master": entry["master"],
            "hidden": entry.get("hidden", False),
            "background": entry["background"],
            "reading_order": entry["reading_order"],
            "slots": slots,
            "shape_count": len(entry["shapes"]),
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
        })

    shutil.rmtree(outdir / "_render", ignore_errors=True)
    (outdir / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False)
    )
    return manifest


def main():
    ap = argparse.ArgumentParser(
        description="Extract per-layout dossiers from a template for a vision model"
    )
    ap.add_argument("input", help="Template .pptx/.potx/.odp/.otp to survey")
    ap.add_argument("-o", "--outdir", default="catalog",
                    help="Output directory (default: catalog/)")
    ap.add_argument("--dpi", type=int, default=150, help="Render resolution")
    args = ap.parse_args()

    source = Path(args.input)
    if source.suffix.lower() not in PRESENTATION_EXTS:
        ap.error(f"not a presentation: {source}")

    manifest = extract(source, Path(args.outdir), args.dpi)
    n = len(manifest["slides"])
    masters = sorted({s["master"] for s in manifest["slides"]})
    print(f"Wrote {n} layout dossiers to {args.outdir}/")
    print(f"Masters seen: {', '.join(masters)}")
    print(f"Next: feed each slide-NN/ to your model with "
          f"references/gemma-catalog-prompt.md")


if __name__ == "__main__":
    main()
