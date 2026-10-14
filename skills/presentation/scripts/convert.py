#!/usr/bin/env python3
"""
Convert a presentation between formats.

    python3 scripts/convert.py deck.odp deck.pptx     # ODF  -> OOXML
    python3 scripts/convert.py deck.pptx deck.odp     # OOXML -> ODF
    python3 scripts/convert.py old.ppt new.pptx       # legacy binary -> OOXML
    python3 scripts/convert.py deck.odp out.pdf       # for rendering / review

Converting is lossy in one direction and not the other: ODF is LibreOffice's own
format, so .odp -> .odp is exact, while .odp -> .pptx goes through an export
filter.  Convert only when the target format is what you actually need to ship;
to *edit* a deck, open it and save it in its own format instead (office.impress).
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))

from office.impress import (  # noqa: E402
    FILTERS,
    PRESENTATION_EXTS,
    SofficeError,
    filter_for,
    impress_document,
    save_as,
)


def convert(source: Path, target: Path) -> Path:
    """Open ``source`` and write it to ``target`` in the target's format."""
    with impress_document(source, save=False) as doc:
        return save_as(doc, target)


def main():
    parser = argparse.ArgumentParser(description="Convert a presentation between formats")
    parser.add_argument("input", help="Source presentation")
    parser.add_argument(
        "output",
        help=f"Target file; its extension picks the filter ({', '.join(sorted(FILTERS))})",
    )
    args = parser.parse_args()

    source = Path(args.input)
    target = Path(args.output)

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
    if source.resolve() == target.resolve():
        print("Error: input and output are the same file", file=sys.stderr)
        return 2
    if not target.parent.is_dir():
        print(f"Error: {target.parent} does not exist", file=sys.stderr)
        return 2

    try:
        filter_for(target)  # fail before starting LibreOffice on a bad extension
        written = convert(source, target)
    except (SofficeError, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    print(f"Wrote {written} ({written.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
