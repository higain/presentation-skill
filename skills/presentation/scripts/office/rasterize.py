"""Deck to images: the one PDF-and-``pdftoppm`` pipeline in this skill.

Every visual step here -- ``render.py``'s per-slide PNGs, ``thumbnail.py``'s
contact sheet, ``catalog_extract.py``'s per-layout renders -- is the same two
subprocesses: LibreOffice prints the deck to PDF, ``pdftoppm`` cuts the PDF into
one image per page.  It used to be written out three times, and the three copies
had drifted into three different answers to the hidden-slide problem:
``render.py`` let the numbering shift and said so in its docstring,
``thumbnail.py`` substituted a grey tile, ``catalog_extract.py`` opened a
writable copy and un-hid every slide.  All three are right for their caller and
none of them is right for the others, so the policy is now an argument.

``hidden=``:

``"drop"``
    What LibreOffice does: hidden slides are not printed.  The images are the
    deck as it would be presented, and **image N is not slide N** after the
    first hidden slide.
``"show"``
    Render from an all-visible copy, so there is exactly one image per slide, in
    order.  Use this whenever an image has to line up with a slide number.

The caller that wants a placeholder tile for a hidden slide -- ``thumbnail.py``,
which labels one -- asks for ``"drop"`` and fills the gaps itself; it is the
only one that has anything to draw a placeholder with.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from .impress import PRESENTATION_EXTS, SofficeError, load_document, save_as, slides
from .soffice import run_soffice

#: pdftoppm names its output ``<prefix>-<n>.<ext>``, zero-padded to the page
#: count.  Everything downstream globs for this.
SLIDE_PREFIX = "slide"

_NUMBER_RE = re.compile(r"-(\d+)\.[a-z]+$", re.I)


def image_number(path: Path) -> int:
    """The page number ``pdftoppm`` put in a file name, or 0."""
    match = _NUMBER_RE.search(path.name)
    return int(match.group(1)) if match else 0


def existing_images(directory: Path, prefix: str = SLIDE_PREFIX) -> list[Path]:
    """Images already rendered into ``directory``, in page order.

    Sorted by the number in the name, so 10 does not sort before 2.
    """
    found = [p for p in Path(directory).glob(f"{prefix}-*")
             if p.suffix.lower() in (".png", ".jpg", ".jpeg")]
    return sorted(found, key=image_number)


def to_pdf(source: Path, outdir: Path, timeout: int = 300) -> Path:
    """Print a deck to PDF and return the path.  This is the slow step."""
    result = run_soffice(
        ["--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(source)],
        timeout=timeout,
    )
    pdf = Path(outdir) / f"{Path(source).stem}.pdf"
    if not pdf.is_file():
        detail = (result.stderr or result.stdout or "").strip()
        raise SofficeError(
            f"LibreOffice produced no PDF for {source}"
            + (f": {detail[:400]}" if detail else "")
        )
    return pdf


def all_visible_copy(source: Path, dest: Path) -> Path:
    """Copy ``source`` to ``dest`` with every slide made visible.

    The source is copied *first* and only the copy is opened read-write: a
    template is often installed read-only, and opening a non-writable file
    read-write fails obscurely as "type detection failed".
    """
    shutil.copyfile(source, dest)
    doc = load_document(dest, read_only=False)
    try:
        changed = False
        for page in slides(doc):
            if not getattr(page, "Visible", True):
                page.Visible = True
                changed = True
        if changed:
            save_as(doc, dest)
    finally:
        doc.close(False)
    return dest


def rasterize(source, outdir, dpi: int = 150, fmt: str = "png",
              wanted: list[int] | None = None, hidden: str = "drop",
              prefix: str = SLIDE_PREFIX, clean: bool = True) -> list[Path]:
    """Write one image per rendered page into ``outdir``; return them in order.

    ``wanted`` is a list of 1-based page numbers, ``fmt`` is ``"png"`` or
    ``"jpeg"``, ``hidden`` is ``"drop"`` or ``"show"`` (see the module
    docstring).  ``clean`` removes stale images from a previous run first --
    without it you review a deck that no longer exists.
    """
    source, outdir = Path(source), Path(outdir)
    if source.suffix.lower() not in PRESENTATION_EXTS:
        raise ValueError(
            f"{source} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}")
    if hidden not in ("drop", "show"):
        raise ValueError(f"hidden must be 'drop' or 'show', not {hidden!r}")
    if dpi < 1:
        raise ValueError("dpi must be positive")

    outdir.mkdir(parents=True, exist_ok=True)
    if clean:
        for stale in existing_images(outdir, prefix):
            stale.unlink()

    ext = "jpg" if fmt in ("jpeg", "jpg") else "png"
    with tempfile.TemporaryDirectory(prefix="rasterize-") as tmp:
        render_source = source
        if hidden == "show":
            render_source = all_visible_copy(
                source, Path(tmp) / f"visible{source.suffix}")
        pdf = to_pdf(render_source, Path(tmp))
        cmd = ["pdftoppm", f"-{'jpeg' if ext == 'jpg' else 'png'}", "-r", str(dpi)]
        if wanted:
            cmd += ["-f", str(min(wanted)), "-l", str(max(wanted))]
        cmd += [str(pdf), str(outdir / prefix)]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True)
        except FileNotFoundError:
            raise SofficeError(
                "pdftoppm not found; install it with: apt install poppler-utils"
            ) from None
        if proc.returncode != 0:
            raise SofficeError(
                f"pdftoppm failed: {(proc.stderr or '').strip()[:400]}")

    images = existing_images(outdir, prefix)
    if wanted:
        keep = set(wanted)
        for image in list(images):
            if image_number(image) not in keep:
                image.unlink()
                images.remove(image)
    return images
