#!/usr/bin/env python3
"""
Render a deck to one PNG per slide, for visual QA.  Works on .odp and .pptx.

    python3 scripts/render.py deck.odp                    # -> render/slide-01.png ...
    python3 scripts/render.py deck.pptx -o qa --dpi 150
    python3 scripts/render.py deck.odp --slides 3,7-9     # only those slides

Prints the absolute path of every image it wrote, one per line, so the paths can
be passed straight to an image viewer.  The output directory is cleared of stale
slide-*.png first -- otherwise images from a previous run survive alongside the
new ones and you review a deck that no longer exists.

**Image N is slide N.**  Hidden slides are rendered too, from an all-visible
copy, so the numbering lines up with `qa.py`, `probe.py` and `dump.py` and a
finding on slide 7 is in `slide-07.png`.  A deck as it would be presented --
hidden slides left out -- is what `convert.py deck.pptx deck.pdf` gives you.

Rendering goes through PDF, which is what LibreOffice itself would print, so
what you see here is what the deck looks like -- with one caveat: fonts the
machine does not have are substituted, and a substitute with different metrics
changes where text wraps.  See the Typography section of SKILL.md.
"""

import argparse
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import PRESENTATION_EXTS, SofficeError, parse_slide_spec  # noqa: E402
from office.rasterize import rasterize  # noqa: E402

_parse_slides = parse_slide_spec


def render(source: Path, outdir: Path, dpi: int,
           wanted: list[int] | None = None, hidden: str = "show") -> list[Path]:
    """One PNG per rendered slide.  See ``office.rasterize`` for ``hidden``."""
    return rasterize(source, outdir, dpi=dpi, fmt="png", wanted=wanted,
                     hidden=hidden)


def main():
    parser = argparse.ArgumentParser(description="Render a deck to one PNG per slide")
    parser.add_argument("input", help="Presentation to render")
    parser.add_argument(
        "-o", "--outdir", default="render", help="Output directory (default: render/)"
    )
    parser.add_argument(
        "--dpi", type=int, default=150, help="Resolution (default: 150)"
    )
    parser.add_argument(
        "--slides", help="Only these slides, 1-based, e.g. '3,7-9' (default: all)"
    )
    args = parser.parse_args()

    source = Path(args.input)
    if not source.exists():
        print(f"Error: {source} does not exist", file=sys.stderr)
        return 2
    if source.suffix.lower() not in PRESENTATION_EXTS:
        print(
            f"Error: {source} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        return 2
    if args.dpi < 1:
        print("Error: --dpi must be positive", file=sys.stderr)
        return 2
    if shutil.which("soffice") is None and shutil.which("libreoffice") is None:
        print(
            "Error: soffice not found on PATH; install LibreOffice "
            "(apt install libreoffice-impress)",
            file=sys.stderr,
        )
        return 2

    wanted = None
    if args.slides:
        try:
            wanted = _parse_slides(args.slides)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            return 2

    try:
        images = render(source, Path(args.outdir), args.dpi, wanted)
    except SofficeError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    if not images:
        print("Error: no images were produced", file=sys.stderr)
        return 1
    for image in images:
        print(image.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())
