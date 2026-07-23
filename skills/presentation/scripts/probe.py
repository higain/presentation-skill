#!/usr/bin/env python3
"""
Print a deck's shapes: index, geometry, colours and character formatting.

    python3 scripts/probe.py template.pptx
    python3 scripts/probe.py template.pptx --slides 4,7-9

``dump.py`` tells you what a slide *says*; this tells you what it *is*.  To
retext a template slide you need the index of each shape (``shapes(page)[5]``),
where it sits, how wide it is, and what font and colour its text carries --
none of which dump.py reports.  Probe once, then write against the indices.

Geometry is in 1/100 mm, the unit UNO works in.  Colours are hex RGB.

    ## Slide 7 — layout 20, master «DEFAULT», background #FDF6F1
      idx  shape       type              x      y      w      h  fill     line
      [ 2] Shape 2     ooxml-roundRect  2910   6625  22039   9413 #FFFFFF  #ECDFD6/46
      [ 4] Text 4      ooxml-rect       4544   7941   2203    725 -        -
           "SPEED"  ·  Arial 15.0 bold #FFFFFF kern 40 left

      reading order — 3 columns
      column 1    [3] 'Background'  [4] 'Set the context first. ...'
      column 2    [5] 'What we found'  [6] 'Present the...'  [7] 'Two shorter...'
      column 3    [8] 'What it means'  [9] 'Close with the implicat...'

The reading-order block groups the same shapes by column, because the table
above is in z-order and a column whose body is split across two text boxes
looks like an even grid there when it is not.  That single mistake -- assuming
three columns of two shapes on the slide sketched above -- is what puts column
three's heading into column two and leaves the template's own words on the
slide.  ``office.template.retext()`` avoids the indices altogether.
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import (  # noqa: E402
    PRESENTATION_EXTS,
    SofficeError,
    load_document,
    shapes,
    slides,
)
from office.template import read_text  # noqa: E402

ADJUST = {0: "left", 1: "right", 2: "block", 3: "center", 4: "stretch"}


def _hex(value):
    return f"#{value & 0xFFFFFF:06X}"


def _fill(shape):
    style = str(getattr(shape, "FillStyle", ""))
    if "NONE" in style:
        return "-"
    if "SOLID" in style:
        colour = _hex(shape.FillColor)
        alpha = getattr(shape, "FillTransparence", 0)
        return f"{colour}{f'@{alpha}%' if alpha else ''}"
    return style.split("'")[1] if "'" in style else "?"


def _line(shape):
    style = str(getattr(shape, "LineStyle", ""))
    if "NONE" in style:
        return "-"
    return f"{_hex(shape.LineColor)}/{getattr(shape, 'LineWidth', 0)}"


def _geometry_type(shape):
    """The ooxml/odf preset name, e.g. ooxml-roundRect -- what the shape IS."""
    geometry = getattr(shape, "CustomShapeGeometry", None)
    if geometry:
        for prop in geometry:
            if prop.Name == "Type":
                return prop.Value
    return shape.getShapeType().rsplit(".", 1)[-1]


def _format_line(cprops, pprops):
    """One line describing how a paragraph is set."""
    bits = []
    name = cprops.get("CharFontName")
    height = cprops.get("CharHeight")
    if name:
        bits.append(f"{name} {height:g}" if height else name)
    if (cprops.get("CharWeight") or 100) > 110:
        bits.append("bold")
    if "CharColor" in cprops:
        bits.append(_hex(cprops["CharColor"]))
    kern = cprops.get("CharKerning")
    if kern:
        bits.append(f"kern {kern}")
    adjust = pprops.get("ParaAdjust")
    if adjust is not None:
        bits.append(ADJUST.get(int(adjust), str(adjust)))
    spacing = pprops.get("ParaLineSpacing")
    if spacing is not None and getattr(spacing, "Height", 100) != 100:
        bits.append(f"lh {spacing.Height}%")
    return "  ·  ".join(bits)


def _columns(page, slide_width):
    """Text shapes grouped into columns by horizontal overlap, in reading order.

    ``(full_width, columns)``.  Shapes spanning most of the slide -- eyebrow,
    title, standfirst -- are not part of any column and come back separately;
    the rest are clustered on their left edge and sorted top to bottom.

    This exists because the shape table above is in z-order, which is the order
    shapes were added and says nothing about the layout.  A three-column slide
    whose middle column carries two body boxes rather than one looks like an
    even grid in z-order and is not one.
    """
    items = []
    for i, shape in enumerate(shapes(page)):
        try:
            text = " ".join(shape.getString().split())
            pos, size = shape.getPosition(), shape.getSize()
        except Exception:
            continue
        if text:
            items.append((i, pos.X, pos.Y, size.Width, text))

    # Cluster on the left edge.  A designed column shares one, and every other
    # signal is worse: overlapping x-ranges let a single wide caption chain a
    # row of separate bars into one bogus column, whichever way the comparison
    # is written.  Tolerance is a hair over 1% of the slide, and each shape is
    # compared with its column's seed so near-misses cannot chain.
    tolerance = max(slide_width * 0.012, 1)
    wide = [it for it in items if it[3] > slide_width * 0.55]
    columns = []
    for item in sorted((it for it in items if it[3] <= slide_width * 0.55),
                       key=lambda t: t[1]):
        for column in columns:
            if abs(item[1] - column[0][1]) <= tolerance:
                column.append(item)
                break
        else:
            columns.append([item])

    columns = sorted((sorted(c, key=lambda t: t[2]) for c in columns),
                     key=lambda c: min(t[1] for t in c))
    return sorted(wide, key=lambda t: t[2]), columns


def _reading_order(page, slide_width):
    """The columns block printed under a slide's shape table, or ``[]``."""
    wide, columns = _columns(page, slide_width)
    if len(columns) < 2:
        return []

    def row(label, items):
        cells = "  ".join(
            f"[{i}] {(text if len(text) <= 26 else text[:23] + '...')!r}"
            for i, _, _, _, text in items
        )
        return f"  {label:<11} {cells}"

    out = ["", f"  reading order — {len(columns)} columns"]
    if wide:
        out.append(row("full width", wide))
    for n, column in enumerate(columns, start=1):
        out.append(row(f"column {n}", column))

    # Only worth warning about where the slide reads as a repeated block: two
    # or more columns each holding a stack.  A cover with a title on the left
    # and a URL on the right is uneven by design and needs no warning.
    counts = [len(c) for c in columns]
    stacked = [n for n in counts if n > 1]
    if len(stacked) > 1 and len(set(stacked)) > 1:
        out.append(
            f"  ** columns hold {', '.join(map(str, counts))} shapes — this is "
            f"NOT an even grid. Take each column's indices from the list above; "
            f"do not compute them. **"
        )
    return out


def _background(page):
    bg = page.Background
    if bg is None:
        return "none"
    try:
        if "SOLID" in str(bg.FillStyle):
            return _hex(bg.FillColor)
        return str(bg.FillStyle).split("'")[1]
    except Exception:
        return "?"


def _wanted(spec, total):
    if not spec:
        return set(range(1, total + 1))
    out = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-", 1)
            out.update(range(int(lo), int(hi) + 1))
        elif part:
            out.add(int(part))
    return out


def probe(path: Path, slide_spec=None) -> str:
    doc = load_document(path, read_only=True)
    try:
        pages = slides(doc)
        want = _wanted(slide_spec, len(pages))
        width, height = pages[0].Width, pages[0].Height
        out = [
            f"# {path.name}",
            "",
            f"{len(pages)} slide(s), {width} × {height} (1/100 mm)",
            "",
            "`[i]` is exactly `shapes(page)[i]` — the index set_text() writes to.",
            "The table is in z-order, which is not reading order; where a slide "
            "has columns, the block under it groups the same shapes by column.",
            "",
        ]
        for number, page in enumerate(pages, start=1):
            if number not in want:
                continue
            head = (
                f"## Slide {number} — layout {page.Layout}, "
                f"master «{page.MasterPage.Name}», background {_background(page)}"
            )
            if not getattr(page, "Visible", True):
                head += "  *(hidden)*"
            out += [head, ""]
            out.append(
                f"{'idx':>5} {'shape':<12} {'type':<18}"
                f"{'x':>7}{'y':>7}{'w':>7}{'h':>7}  {'fill':<10}{'line'}"
            )
            for i, shape in enumerate(shapes(page)):
                pos, size = shape.getPosition(), shape.getSize()
                out.append(
                    f"{'[' + str(i) + ']':>5} {shape.Name[:12]:<12} "
                    f"{_geometry_type(shape)[:18]:<18}"
                    f"{pos.X:>7}{pos.Y:>7}{size.Width:>7}{size.Height:>7}  "
                    f"{_fill(shape):<10}{_line(shape)}"
                )
                try:
                    paras = read_text(shape)
                except Exception:
                    paras = []
                for text, cprops, pprops in paras:
                    if not text.strip():
                        continue
                    shown = text if len(text) <= 64 else text[:61] + "..."
                    out.append(f"{'':>5} {shown!r}")
                    style = _format_line(cprops, pprops)
                    if style:
                        out.append(f"{'':>7} {style}")
            out += _reading_order(page, width)
            out.append("")
        return "\n".join(out).rstrip() + "\n"
    finally:
        doc.close(False)


def probe_data(path: Path, slide_spec=None) -> dict:
    """The same survey as ``probe()``, as data -- what ``--json`` prints.

    Shapes keep their ``index``, so anything selected here can be written to
    with ``shapes(page)[index]`` straight away.
    """
    doc = load_document(path, read_only=True)
    try:
        pages = slides(doc)
        want = _wanted(slide_spec, len(pages))
        width, height = pages[0].Width, pages[0].Height
        out = {
            "file": path.name,
            "slide_width": width,
            "slide_height": height,
            "slides": [],
        }
        for number, page in enumerate(pages, start=1):
            if number not in want:
                continue
            entry = {
                "number": number,
                "layout": page.Layout,
                "master": page.MasterPage.Name,
                "background": _background(page),
                "hidden": not getattr(page, "Visible", True),
                "shapes": [],
            }
            for i, shape in enumerate(shapes(page)):
                pos, size = shape.getPosition(), shape.getSize()
                try:
                    paras = read_text(shape)
                except Exception:
                    paras = []
                entry["shapes"].append({
                    "index": i,
                    "name": shape.Name,
                    "type": _geometry_type(shape),
                    "x": pos.X, "y": pos.Y,
                    "width": size.Width, "height": size.Height,
                    "fill": _fill(shape),
                    "line": _line(shape),
                    "text": "\n".join(t for t, _, _ in paras),
                    "paragraphs": [
                        {
                            "text": text,
                            "font": cprops.get("CharFontName"),
                            "size": cprops.get("CharHeight"),
                            "bold": (cprops.get("CharWeight") or 100) > 110,
                            "color": (_hex(cprops["CharColor"])
                                      if "CharColor" in cprops else None),
                        }
                        for text, cprops, _ in paras
                    ],
                })
            wide, columns = _columns(page, width)
            entry["reading_order"] = {
                "full_width": [it[0] for it in wide],
                "columns": [[it[0] for it in column] for column in columns],
            }
            out["slides"].append(entry)
        return out
    finally:
        doc.close(False)


def main():
    parser = argparse.ArgumentParser(
        description="Print a deck's shape indices, geometry and formatting"
    )
    parser.add_argument("input", help="Presentation to read")
    parser.add_argument("-o", "--output", help="Write here (default: stdout)")
    parser.add_argument("--slides", help="Only these slides, e.g. 4,7-9")
    parser.add_argument(
        "--json", action="store_true",
        help="Emit JSON instead of the table, to pick shapes programmatically",
    )
    args = parser.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"Error: {path} does not exist", file=sys.stderr)
        return 2
    if path.suffix.lower() not in PRESENTATION_EXTS:
        print(
            f"Error: {path} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        return 2

    try:
        if args.json:
            text = json.dumps(probe_data(path, args.slides), indent=2,
                              ensure_ascii=False) + "\n"
        else:
            text = probe(path, args.slides)
    except (SofficeError, FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
