#!/usr/bin/env python3
"""Create a labelled contact sheet of every slide in a deck -- .odp or .pptx.

One grid image of all slides, for picking which template layout to reuse.  Each
thumbnail is labelled with its slide number and master; hidden slides get a
crossed-out placeholder instead of a render.

For full-resolution visual QA of a finished deck, use render.py instead -- this
is for surveying a deck at a glance.

Usage:
    python3 scripts/thumbnail.py template.odp [output_prefix] [--cols N]

Examples:
    python3 scripts/thumbnail.py template.odp
    # Creates: template.jpg  (named after the deck, so two decks never collide)

    python3 scripts/thumbnail.py template.pptx grid --cols 4
    # Creates: grid.jpg (or grid-1.jpg, grid-2.jpg for large decks)
"""

import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))

from office.impress import PRESENTATION_EXTS, load_document, slides as doc_slides
from office.rasterize import existing_images, rasterize
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as exc:                      # pragma: no cover
    raise ImportError("thumbnail.py needs Pillow: apt install python3-pil") from exc


THUMBNAIL_WIDTH = 300
CONVERSION_DPI = 100
MAX_COLS = 6
DEFAULT_COLS = 3
JPEG_QUALITY = 95
GRID_PADDING = 20
BORDER_WIDTH = 2
FONT_SIZE_RATIO = 0.10
LABEL_PADDING_RATIO = 0.4


def main():
    parser = argparse.ArgumentParser(
        description="Create a labelled contact sheet of a deck's slides."
    )
    parser.add_argument("input", help="Presentation (.odp, .pptx, ...)")
    parser.add_argument(
        "output_prefix",
        nargs="?",
        default=None,
        help="Output prefix for image files (default: the deck's own name)",
    )
    parser.add_argument(
        "--cols",
        type=int,
        default=DEFAULT_COLS,
        help=f"Number of columns (default: {DEFAULT_COLS}, max: {MAX_COLS})",
    )
    parser.add_argument(
        "--from-images",
        metavar="DIR",
        help="Reuse the per-slide images already in DIR -- as render.py writes "
             "them -- instead of converting the deck a second time",
    )

    args = parser.parse_args()

    if args.cols < 1:
        print("Error: --cols must be at least 1", file=sys.stderr)
        sys.exit(2)
    cols = min(args.cols, MAX_COLS)
    if args.cols > MAX_COLS:
        print(f"Warning: Columns limited to {MAX_COLS}")

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: {input_path} does not exist", file=sys.stderr)
        sys.exit(2)
    if input_path.suffix.lower() not in PRESENTATION_EXTS:
        print(
            f"Error: {input_path} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        sys.exit(2)

    # Defaulting to the deck's own name keeps two decks thumbnailed in one
    # directory from silently overwriting each other's grids.
    output_path = Path(f"{args.output_prefix or input_path.stem}.jpg")

    try:
        slide_info = get_slide_info(input_path)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            # Converting the deck costs a full LibreOffice round trip.  When the
            # caller has already rendered every slide -- which make_deck.py has,
            # one step earlier -- doing it again is half the wall clock of the
            # command for a second copy of the same pixels.
            if args.from_images:
                visible_images = existing_images(Path(args.from_images))
                if not visible_images:
                    print(f"Error: no slide images in {args.from_images}",
                          file=sys.stderr)
                    sys.exit(1)
            else:
                visible_images = convert_to_images(input_path, temp_path)

            if not visible_images and not any(s["hidden"] for s in slide_info):
                print("Error: No slides found", file=sys.stderr)
                sys.exit(1)

            slides = build_slide_list(slide_info, visible_images, temp_path)

            grid_files = create_grids(slides, cols, THUMBNAIL_WIDTH, output_path)

            print(f"Created {len(grid_files)} grid(s):")
            for grid_file in grid_files:
                print(f"  {grid_file}")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


def get_slide_info(path: Path) -> list[dict]:
    """Slide number, label and hidden flag, straight from the document model.

    Read through UNO rather than the OOXML package, so this works on .odp as
    well -- and reports the deck as LibreOffice will render it, which is what
    the thumbnails have to line up with.
    """
    doc = load_document(path, read_only=True)
    try:
        info = []
        for number, page in enumerate(doc_slides(doc), start=1):
            label = f"{number}. {page.MasterPage.Name}"
            info.append({"name": label, "hidden": not getattr(page, "Visible", True)})
        return info
    finally:
        doc.close(False)


def build_slide_list(
    slide_info: list[dict],
    visible_images: list[Path],
    temp_dir: Path,
) -> list[tuple[Path, str]]:
    visible_count = sum(1 for info in slide_info if not info["hidden"])
    rendered_hidden = len(visible_images) == len(slide_info) != visible_count

    if not rendered_hidden and visible_count != len(visible_images):
        raise ValueError(
            f"LibreOffice rendered {len(visible_images)} page(s) for {visible_count} "
            f"visible slide(s) of {len(slide_info)}; thumbnails would be mislabeled"
        )

    if visible_images:
        with Image.open(visible_images[0]) as img:
            placeholder_size = img.size
    else:
        placeholder_size = (1920, 1080)

    slides = []
    visible_idx = 0

    for info in slide_info:
        if info["hidden"] and not rendered_hidden:
            placeholder_path = temp_dir / f"hidden-{info['name']}.jpg"
            placeholder_img = create_hidden_placeholder(placeholder_size)
            placeholder_img.save(placeholder_path, "JPEG")
            slides.append((placeholder_path, f"{info['name']} (hidden)"))
        else:
            label = f"{info['name']} (hidden)" if info["hidden"] else info["name"]
            slides.append((visible_images[visible_idx], label))
            visible_idx += 1

    return slides


def create_hidden_placeholder(size: tuple[int, int]) -> Image.Image:
    img = Image.new("RGB", size, color="#F0F0F0")
    draw = ImageDraw.Draw(img)
    line_width = max(5, min(size) // 100)
    draw.line([(0, 0), size], fill="#CCCCCC", width=line_width)
    draw.line([(size[0], 0), (0, size[1])], fill="#CCCCCC", width=line_width)
    return img


def convert_to_images(deck_path: Path, temp_dir: Path) -> list[Path]:
    """Thumbnail-resolution JPEGs, one per rendered slide.

    ``hidden="drop"`` because this is the one caller that has something better
    to put in the gap: ``build_slide_list`` draws a crossed-out placeholder and
    labels it, which a reader of a contact sheet wants to see.
    """
    return rasterize(deck_path, temp_dir, dpi=CONVERSION_DPI, fmt="jpeg",
                     hidden="drop")


def create_grids(
    slides: list[tuple[Path, str]],
    cols: int,
    width: int,
    output_path: Path,
) -> list[str]:
    max_per_grid = cols * (cols + 1)
    grid_files = []

    for chunk_idx, start_idx in enumerate(range(0, len(slides), max_per_grid)):
        end_idx = min(start_idx + max_per_grid, len(slides))
        chunk_slides = slides[start_idx:end_idx]

        grid = create_grid(chunk_slides, cols, width)

        if len(slides) <= max_per_grid:
            grid_filename = output_path
        else:
            stem = output_path.stem
            suffix = output_path.suffix
            grid_filename = output_path.parent / f"{stem}-{chunk_idx + 1}{suffix}"

        grid_filename.parent.mkdir(parents=True, exist_ok=True)
        grid.save(str(grid_filename), quality=JPEG_QUALITY)
        grid_files.append(str(grid_filename))

    return grid_files


def create_grid(
    slides: list[tuple[Path, str]],
    cols: int,
    width: int,
) -> Image.Image:
    font_size = int(width * FONT_SIZE_RATIO)
    label_padding = int(font_size * LABEL_PADDING_RATIO)

    with Image.open(slides[0][0]) as img:
        aspect = img.height / img.width
    height = int(width * aspect)

    rows = (len(slides) + cols - 1) // cols
    grid_w = cols * width + (cols + 1) * GRID_PADDING
    grid_h = rows * (height + font_size + label_padding * 2) + (rows + 1) * GRID_PADDING

    grid = Image.new("RGB", (grid_w, grid_h), "white")
    draw = ImageDraw.Draw(grid)

    try:
        font = ImageFont.load_default(size=font_size)
    except Exception:
        font = ImageFont.load_default()

    for i, (img_path, slide_name) in enumerate(slides):
        row, col = i // cols, i % cols
        x = col * width + (col + 1) * GRID_PADDING
        y_base = (
            row * (height + font_size + label_padding * 2) + (row + 1) * GRID_PADDING
        )

        label = slide_name
        bbox = draw.textbbox((0, 0), label, font=font)
        text_w = bbox[2] - bbox[0]
        draw.text(
            (x + (width - text_w) // 2, y_base + label_padding),
            label,
            fill="black",
            font=font,
        )

        y_thumbnail = y_base + label_padding + font_size + label_padding

        with Image.open(img_path) as img:
            img.thumbnail((width, height), Image.Resampling.LANCZOS)
            w, h = img.size
            tx = x + (width - w) // 2
            ty = y_thumbnail + (height - h) // 2
            grid.paste(img, (tx, ty))

            if BORDER_WIDTH > 0:
                draw.rectangle(
                    [
                        (tx - BORDER_WIDTH, ty - BORDER_WIDTH),
                        (tx + w + BORDER_WIDTH - 1, ty + h + BORDER_WIDTH - 1),
                    ],
                    outline="gray",
                    width=BORDER_WIDTH,
                )

    return grid


if __name__ == "__main__":
    sys.exit(main())
