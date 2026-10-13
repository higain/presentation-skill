#!/usr/bin/env python3
"""Collect the template's own icons into the pack.

    python3 scripts/harvest_icons.py acme                      # from the pack's template
    python3 scripts/harvest_icons.py acme --from deck.pptx     # and from a finished deck

A corporate master usually carries a few pictograms -- a tick, a magnifier, a
warning triangle -- inserted from Office's icon library.  They are the brand's
icon style, and without this step no deck can use them: ``D.icon`` finds
icons in ``templates/<pack>/icons/`` first and only then in an installed set.

What counts as an icon: an SVG picture on a slide, roughly square, small,
with no raster inside it.  Photos, logos drawn as wide SVGs and illustrations
are left alone.  Names come from Office's own id (``Icons_MagnifyingGlass``
becomes ``magnifying-glass``), else the picture's alt text, else
``icon-<slide>-<n>`` -- rename those in ``icons/index.json``.  Re-running
keeps every name and keyword already in the index: entries are matched on the
icon's content, so a rename survives.

No model is involved and nothing leaves the machine.  The icons are the
brand's, like the template: ``templates/*`` stays out of git.
"""

import argparse
import hashlib
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine  # noqa: E402

engine()

from office import icons  # noqa: E402
from office.pack import pack_named  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
SVG_BLIP = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip"

#: Larger than this is an illustration, not a pictogram.
MAX_BYTES = 40_000
#: Width over height an icon may have.
ASPECT = (0.6, 1.67)
#: Distinct paints an icon may use and still be recoloured as one.
MONO = 1

#: Generic shape names Office and LibreOffice give a picture; not a name.
GENERIC = re.compile(r"^(grafik|graphic|bild|picture|image|icon|symbol|"
                     r"object|objekt|inhaltsplatzhalter|platzhalter|"
                     r"content placeholder|placeholder)\s*\d*$", re.I)
#: Office's alt text boilerplate, in the languages a template arrives in.
ALT_NOISE = re.compile(
    r"\b(with solid fill|mit einfarbiger füllung|silhouette|outline|"
    r"umriss|kontur|mit füllung|einfarbig)\b", re.I)

#: Office icon ids that are the same picture mirrored.
DIRECTION = re.compile(r"_(LTR|RTL|M)$")


def slug(text):
    text = re.sub(r"([a-z0-9])([A-Z])", r"\1-\2", text)
    text = (text.lower().replace("ä", "ae").replace("ö", "oe")
            .replace("ü", "ue").replace("ß", "ss"))
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")


def office_id(svg):
    m = re.search(r'\bid="Icons_([A-Za-z0-9_]+)"', svg)
    return DIRECTION.sub("", m.group(1)) if m else None


def name_for(svg, alt, shape_name):
    oid = office_id(svg)
    if oid:
        return slug(oid), [w for w in slug(oid).split("-") if w]
    for text in (alt, shape_name):
        text = ALT_NOISE.sub("", text or "").strip(" ,.-")
        if text and not GENERIC.match(text):
            return slug(text)[:40], [w for w in slug(text).split("-") if w]
    return None, []


def is_icon(svg, size):
    if size > MAX_BYTES or "<image" in svg:
        return False
    return ASPECT[0] <= icons.aspect(svg) <= ASPECT[1]


def is_mono(svg):
    """One visible paint.  Office writes the theme colour twice -- a ``<style>``
    class and a ``fill`` attribute, not always the same value, and the class
    wins -- so the two places are counted apart: two colours means two in
    either one, not one in each."""
    blocks = "".join(re.findall(r"<style\b.*?</style>", svg, flags=re.S))
    rest = re.sub(r"<style\b.*?</style>", "", svg, flags=re.S)
    return max(len(icons.colours(blocks) - {"currentcolor"}),
               len(icons.colours(rest) - {"currentcolor"})) <= MONO


def fingerprint(svg):
    """Content identity regardless of colour, so a recoloured copy is one icon."""
    plain = icons.recolour(svg, 0)
    plain = re.sub(r'\sid="[^"]*"', "", plain)
    return hashlib.sha256(re.sub(r"\s+", "", plain).encode()).hexdigest()[:16]


def _slide_parts(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = {r.get("Id"): r.get("Target") for r in
            ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
            .findall(f"{REL}Relationship")}
    for s in pres.find(f"{P}sldIdLst"):
        target = rels[s.get(f"{R}id")]
        yield "ppt/" + target.lstrip("/").replace("ppt/", "", 1)


def harvest(deck):
    """``[(svg, alt, shape_name, where)]`` for every SVG picture in ``deck``."""
    out = []
    with zipfile.ZipFile(deck) as z:
        for number, part in enumerate(_slide_parts(z), start=1):
            rels_part = re.sub(r"slides/(slide\d+\.xml)$",
                               r"slides/_rels/\1.rels", part)
            try:
                rels = {r.get("Id"): r.get("Target") for r in
                        ET.fromstring(z.read(rels_part))
                        .findall(f"{REL}Relationship")}
                slide = ET.fromstring(z.read(part))
            except (KeyError, ET.ParseError):
                continue
            for pic in slide.iter(f"{P}pic"):
                blip = pic.find(f".//{SVG_BLIP}")
                if blip is None:
                    continue
                target = rels.get(blip.get(f"{R}embed"), "")
                media = "ppt/" + target.replace("../", "")
                try:
                    data = z.read(media)
                except KeyError:
                    continue
                svg = data.decode("utf-8", "replace")
                if not is_icon(svg, len(data)):
                    continue
                nv = pic.find(f".//{P}cNvPr")
                out.append((svg, nv.get("descr", "") if nv is not None else "",
                            nv.get("name", "") if nv is not None else "",
                            f"{Path(deck).name} slide {number}"))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("pack")
    ap.add_argument("--from", dest="extra", action="append", default=[],
                    metavar="DECK", help="also harvest a finished .pptx")
    args = ap.parse_args()

    pack = pack_named(args.pack)
    if pack is None or pack.template is None:
        sys.exit(f"no pack {args.pack!r} with a template -- run make_pack.py first")
    decks = [pack.template, *map(Path, args.extra)]
    for deck in decks:
        if deck.suffix.lower() not in (".pptx", ".potx"):
            sys.exit(f"{deck}: icons are read from .pptx/.potx only")

    folder = pack.root / icons.ICON_DIRNAME
    folder.mkdir(exist_ok=True)
    index = icons.pack_index(pack)
    by_print = {e.get("print"): n for n, e in index.items() if e.get("print")}
    added, unnamed, multi = [], [], []
    counter = {}

    for deck in decks:
        for svg, alt, shape_name, where in harvest(deck):
            fp = fingerprint(svg)
            if fp in by_print:
                continue
            name, keywords = name_for(svg, alt, shape_name)
            if not name:
                n = counter[where] = counter.get(where, 0) + 1
                name = slug(f"icon {where.rsplit(' ', 1)[-1]} {n}")
                unnamed.append(name)
            base, i = name, 2
            while name in index:
                name, i = f"{base}-{i}", i + 1
            mono = is_mono(svg)
            (folder / f"{name}.svg").write_text(svg, encoding="utf-8")
            index[name] = {"file": f"{name}.svg", "keywords": keywords,
                           "mono": mono, "source": where, "print": fp}
            by_print[fp] = name
            added.append(name)
            if not mono:
                multi.append(name)

    (folder / icons.INDEX_NAME).write_text(
        json.dumps({"icons": dict(sorted(index.items()))}, indent=1,
                   ensure_ascii=False) + "\n", encoding="utf-8")

    print(f"{pack.name}: {len(index)} icons in {folder} "
          f"({len(added)} new: {', '.join(added) or 'none'})")
    if unnamed:
        print(f"  Unnamed, rename in {icons.INDEX_NAME} (the file may keep "
              f"its name): {', '.join(unnamed)}")
    if multi:
        print(f"  Several colours, placed as they are, never recoloured: "
              f"{', '.join(multi)}")
    others = [label for label, _ in icons.sources(pack) if label != "pack"]
    print("  Beyond these, builds use: "
          + (", ".join(others) if others else
             "nothing -- " + icons.INSTALL_HINT))
    print("  Rebuild the catalog so catalog.md lists them: "
          "catalog_build.py <pack>/catalog")
    return 0


if __name__ == "__main__":
    sys.exit(main())
