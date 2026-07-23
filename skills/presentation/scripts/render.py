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

Rendering goes through PDF, which is what LibreOffice itself would print, so
what you see here is what the deck looks like -- with one caveat: fonts the
machine does not have are substituted, and a substitute with different metrics
changes where text wraps.  See the Typography section of SKILL.md.
"""

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import PRESENTATION_EXTS, SofficeError  # noqa: E402
from office.soffice import run_soffice  # noqa: E402


def _parse_slides(spec: str) -> list[int]:
    """'3,7-9' -> [3, 7, 8, 9]."""
    out: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, _, hi = part.partition("-")
            try:
                lo, hi = int(lo), int(hi)
            except ValueError:
                raise ValueError(f"not a slide range: {part!r}") from None
            if lo > hi:
                raise ValueError(f"reversed slide range: {part!r}")
            out.update(range(lo, hi + 1))
        else:
            try:
                out.add(int(part))
            except ValueError:
                raise ValueError(f"not a slide number: {part!r}") from None
    if not out or min(out) < 1:
        raise ValueError("slide numbers start at 1")
    return sorted(out)


def to_pdf(source: Path, outdir: Path) -> Path:
    """Convert a deck to PDF and return the path."""
    result = run_soffice(
        ["--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(source)],
        timeout=300,
    )
    pdf = outdir / f"{source.stem}.pdf"
    if not pdf.is_file():
        detail = (result.stderr or result.stdout or "").strip()
        raise SofficeError(
            f"LibreOffice produced no PDF for {source}"
            + (f": {detail[:400]}" if detail else "")
        )
    return pdf


def render(source: Path, outdir: Path, dpi: int, wanted: list[int] | None) -> list[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    for stale in outdir.glob("slide-*.png"):
        stale.unlink()

    with tempfile.TemporaryDirectory(prefix="render-pdf-") as tmp:
        pdf = to_pdf(source, Path(tmp))
        cmd = ["pdftoppm", "-png", "-r", str(dpi)]
        if wanted:
            cmd += ["-f", str(min(wanted)), "-l", str(max(wanted))]
        cmd += [str(pdf), str(outdir / "slide")]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True)
        except FileNotFoundError:
            raise SofficeError(
                "pdftoppm not found; install it with: apt install poppler-utils"
            ) from None
        if proc.returncode != 0:
            raise SofficeError(f"pdftoppm failed: {(proc.stderr or '').strip()[:400]}")

    images = sorted(outdir.glob("slide-*.png"))
    if wanted:
        keep = set(wanted)
        for image in list(images):
            match = re.search(r"slide-(\d+)\.png$", image.name)
            if match and int(match.group(1)) not in keep:
                image.unlink()
                images.remove(image)
    return images


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
