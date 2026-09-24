#!/usr/bin/env python3
"""
A render, coarse enough to read as text: where the ink is, and where it is not.

    scripts/wireframe.py render/slide-01.png
    scripts/wireframe.py render/ --slides 1,7-9
    scripts/wireframe.py render/slide-01.png --cols 80 --rows 30

**This is for the reviewer who cannot see the image.**  Prefer delegating the
render to the vision-agent, which has eyes; this needs nothing but the PNG and
is the fallback when the vision-agent is unavailable.  It cannot tell you a
diagram is ugly, but it does tell you -- in a form that survives being read as
text -- the two things that went wrong most often in the runs it was written
for: a slide that is mostly empty, and a block of content that is not where the
rest of the slide says it should be.

What one looks like, on a title slide nobody finished::

     0|    .                                                           |
     1|    .:.--.:-:.-:::::=:::-::::--:-.-:--                          |
     2|                                                                |
       (sixteen more blank rows)
    18|                  ----.=--:==--.--:==-:-=---=-                  |
    ink 6% · centroid x=0.45 y=0.47 · largest blank band 67% of the height

Each cell is the mean darkness of that patch of the slide, ramped
`` .:-=+*#%@``.  The grid is deliberately small: at 64x24 a paragraph is a
texture and a heading is a line, which is the level a layout is judged at.

Reading it:

- **ink** -- share of cells with anything in them at all, a pale panel fill
  included.  Under ~8% is an empty slide unless it is a one-line statement.
  It is not ``qa.py``'s "paints N% of the slide", which measures the shapes.
- **centroid** -- 0.5/0.5 is dead centre.  A figure that should be centred and
  reads 0.62 is off to the right, and a slide whose content all sits above
  y=0.35 has a bare lower half.
- **blank band** -- the tallest run of rows nothing paints, which is the shape
  "content crammed under the title" makes.
- **margins** -- the empty gutter on each side, as a share of the slide.
  Unequal left and right margins are a block that is not centred.

Exit codes: 0 drew something, 1 nothing could be read, 2 bad usage.
"""

import argparse
import re
import sys
from pathlib import Path

#: Small enough that a wall of body text is one texture rather than forty rows
#: of noise, wide enough that two columns are visibly two columns.
COLS, ROWS = 64, 24

#: Blank to solid.  The first character is a space, so an empty cell is empty.
RAMP = " .:-=+*#%@"

#: Anything lighter than this fraction of full black is paper, not ink.  A JPEG
#: contact sheet and an anti-aliased render both leave a haze of 1-2% grey over
#: white, and without a floor every cell counts as painted.
#:
#: Low on purpose.  The house template's own panel fill is #F2F2F2 -- 5% ink --
#: so a floor set to keep JPEG noise out at 6% also erased every card on the
#: slide, and a row of three empty cards then measured identically to a row of
#: nothing.  Those are different defects.
INK_FLOOR = 0.02


def slide_number(png):
    """The slide number in a ``slide-07.png``, or ``None``."""
    hit = re.search(r"(\d+)", Path(png).stem)
    return int(hit.group(1)) if hit else None


def find_renders(where, wanted=None):
    """The per-slide PNGs under ``where``, in render order.

    A file is returned as itself; a directory is searched, with contact sheets
    (which have no slide number) left out.
    """
    where = Path(where)
    if where.is_file():
        return [where]
    pngs = [p for p in where.glob("*.png") if slide_number(p) is not None]
    pngs.sort(key=slide_number)
    if wanted:
        pngs = [p for p in pngs if slide_number(p) in wanted]
    return pngs


def ink_grid(png, cols=COLS, rows=ROWS):
    """``[[0.0-1.0, ...], ...]`` -- mean ink per cell, ``rows`` x ``cols``.

    ``Image.BOX`` rather than the default filter on purpose: box resampling
    averages every source pixel in the cell, so a one-pixel hairline rule still
    registers.  A smoothing filter would drop it, and hairlines are how a
    designed slide draws its divisions.
    """
    from PIL import Image                      # optional at import, required here

    with Image.open(png) as im:
        grey = im.convert("L").resize((cols, rows), Image.BOX)
        pixels = grey.load()
        return [[max(0.0, (255 - pixels[x, y]) / 255.0 - INK_FLOOR) / (1 - INK_FLOOR)
                 for x in range(cols)] for y in range(rows)]


def _runs(flags):
    """``[(start, end_inclusive)]`` for each run of ``True`` in ``flags``."""
    out, start = [], None
    for i, flag in enumerate(flags):
        if flag and start is None:
            start = i
        elif not flag and start is not None:
            out.append((start, i - 1))
            start = None
    if start is not None:
        out.append((start, len(flags) - 1))
    return out


def measure(grid):
    """``{ink, cx, cy, band, band_at, margins}`` for one grid.

    ``band`` is the tallest empty run as a share of the height; ``margins`` is
    ``(left, right, top, bottom)``, each the empty gutter as a share of the
    slide.  Everything is a fraction, so none of it depends on the grid size.
    """
    rows, cols = len(grid), len(grid[0])
    painted = [[cell > 0 for cell in row] for row in grid]
    total = sum(sum(row) for row in painted)
    if not total:
        return {"ink": 0.0, "cx": 0.5, "cy": 0.5, "band": 1.0, "band_at": 0.0,
                "margins": (0.0, 0.0, 0.0, 0.0)}

    # Weighted by ink, not by cell: a heading is where the slide's mass is, and
    # a faint rule across the bottom should not drag the centroid onto it.
    weight = sum(sum(row) for row in grid)
    cx = sum(grid[y][x] * (x + 0.5) for y in range(rows)
             for x in range(cols)) / weight / cols
    cy = sum(grid[y][x] * (y + 0.5) for y in range(rows)
             for x in range(cols)) / weight / rows

    empty_rows = [not any(row) for row in painted]
    bands = _runs(empty_rows)
    band = max(bands, key=lambda r: r[1] - r[0]) if bands else None

    used_rows = [y for y in range(rows) if any(painted[y])]
    used_cols = [x for x in range(cols) if any(painted[y][x] for y in range(rows))]
    margins = (used_cols[0] / cols, (cols - 1 - used_cols[-1]) / cols,
               used_rows[0] / rows, (rows - 1 - used_rows[-1]) / rows)

    return {
        "ink": total / (rows * cols),
        "cx": cx, "cy": cy,
        "band": ((band[1] - band[0] + 1) / rows) if band else 0.0,
        "band_at": (band[0] / rows) if band else 0.0,
        "margins": margins,
    }


def summary(stats):
    """The measurements as one line of prose."""
    left, right, top, bottom = stats["margins"]
    parts = [
        f"ink {stats['ink']:.0%}",
        f"centroid x={stats['cx']:.2f} y={stats['cy']:.2f} (0.50/0.50 = centred)",
    ]
    if stats["band"] >= 0.12:
        parts.append(f"largest blank band {stats['band']:.0%} of the height, "
                     f"from y={stats['band_at']:.0%}")
    parts.append(f"margins l/r {left:.0%}/{right:.0%} · t/b {top:.0%}/{bottom:.0%}")
    return " · ".join(parts)


#: Ramp gain.  A cell holding one line of body text is only ~8% ink, so at
#: 1:1 every text cell lands on the first ramp character and a slide of prose
#: reads as a uniform field of dots.  At 3 the ramp spends its range where
#: slides actually live -- text, panels, solid fills -- and only a filled block
#: reaches ``@``.
GAIN = 3.0


def draw(grid):
    """The grid as ``[str]``, one string per row, no numbering."""
    top = len(RAMP) - 1
    return ["".join(RAMP[min(top, int(cell * GAIN * top) + 1)] if cell > 0
                    else RAMP[0] for cell in row) for row in grid]


def block(png, cols=COLS, rows=ROWS, number=None):
    """One slide as markdown: a fenced grid and its measurements."""
    grid = ink_grid(png, cols, rows)
    stats = measure(grid)
    label = f"slide {number}" if number is not None else Path(png).name
    lines = [f"### {label}", "", "```"]
    lines += [f"{i:2d}|{row}|" for i, row in enumerate(draw(grid))]
    lines += ["```", "", summary(stats), ""]
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Read a slide render as text: where the ink is",
    )
    parser.add_argument("path", help="a PNG, or the render directory")
    parser.add_argument("--slides", help="only these, e.g. 4,7-9")
    parser.add_argument("--cols", type=int, default=COLS)
    parser.add_argument("--rows", type=int, default=ROWS)
    args = parser.parse_args()

    wanted = None
    if args.slides:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from office.impress import parse_slide_spec
        try:
            wanted = set(parse_slide_spec(args.slides))
        except ValueError as e:
            print(f"Error: --slides {e}", file=sys.stderr)
            return 2

    target = Path(args.path)
    if not target.exists():
        print(f"Error: {target} does not exist", file=sys.stderr)
        return 2

    pngs = find_renders(target, wanted)
    if not pngs:
        print(f"Error: no slide PNGs in {target}", file=sys.stderr)
        return 1

    drawn = 0
    for png in pngs:
        try:
            print(block(png, args.cols, args.rows, slide_number(png)))
        except Exception as e:                 # a truncated PNG, no PIL
            print(f"### {png.name}\n\ncould not read: {type(e).__name__}: {e}\n")
            continue
        drawn += 1
    return 0 if drawn else 1


if __name__ == "__main__":
    sys.exit(main())
