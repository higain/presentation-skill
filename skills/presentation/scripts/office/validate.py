#!/usr/bin/env python3
"""
Command line tool to validate PowerPoint XML files against XSD schemas.

Usage:
    python validate.py <path> [--original <original_file>] [--auto-repair]

The first argument can be either:
- An unpacked directory containing the presentation XML files
- A packed .pptx or .potx file, which will be unpacked to a temp directory

This checks OOXML packages only.  ODF presentations (.odp/.otp) have no OOXML
schema to check against and do not need one -- LibreOffice writes them itself.
To QA an .odp, render it (scripts/render.py) and read it back (scripts/dump.py).

Auto-repair fixes:
- Missing xml:space="preserve" on <a:t> elements with leading/trailing whitespace
"""

import argparse
import sys
import tempfile
import zipfile
from pathlib import Path

from helpers import OOXML_FAMILY, rezip, safe_extract
from validators import PPTXSchemaValidator

ODF_PRESENTATION_EXTS = {".odp", ".otp", ".fodp"}

PPTX_EXTS = sorted(ext for ext, family in OOXML_FAMILY.items() if family == "pptx")


def _fail(message: str):
    print(f"Error: {message}", file=sys.stderr)
    sys.exit(2)


def _reject_unsupported(path: Path):
    """Turn a wrong-format argument into advice rather than a generic type error."""
    ext = path.suffix.lower()
    if ext in ODF_PRESENTATION_EXTS:
        _fail(
            f"{path} is an ODF presentation. There is no OOXML schema to validate it "
            "against, and it does not need one. To QA it, render it with "
            "scripts/render.py and read it back with scripts/dump.py; to schema-check "
            "it as OOXML, convert it first: scripts/convert.py deck.odp deck.pptx"
        )
    if OOXML_FAMILY.get(ext) in ("docx", "xlsx"):
        _fail(
            f"{path} is a {OOXML_FAMILY[ext]} file; this validator only handles "
            f"presentations ({', '.join(PPTX_EXTS)})."
        )


def main():
    parser = argparse.ArgumentParser(description="Validate PowerPoint XML files")
    parser.add_argument(
        "path",
        help="Path to unpacked directory or packed .pptx/.potx file",
    )
    parser.add_argument(
        "--original",
        required=False,
        default=None,
        help="Path to the template or source .pptx/.potx this deck came from. "
        "Errors the original already had are then not reported as yours.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose output",
    )
    parser.add_argument(
        "--auto-repair",
        action="store_true",
        help="Automatically add missing xml:space=\"preserve\". "
        "Modifies the input in place: repairs to a packed file are written back to it.",
    )
    args = parser.parse_args()

    path = Path(args.path)
    if not path.exists():
        _fail(f"{path} does not exist")
    _reject_unsupported(path)

    original_file = None
    if args.original:
        original_file = Path(args.original)
        if not original_file.is_file():
            _fail(f"{original_file} is not a file")
        _reject_unsupported(original_file)
        if OOXML_FAMILY.get(original_file.suffix.lower()) != "pptx":
            _fail(f"--original must be one of: {', '.join(PPTX_EXTS)}")

    if path.is_file() and OOXML_FAMILY.get(path.suffix.lower()) != "pptx":
        _fail(f"{path} is not a presentation; expected one of: {', '.join(PPTX_EXTS)}")

    packed_file = None
    temp_dir_ctx = None
    if path.is_file() and path.suffix.lower() in OOXML_FAMILY:
        packed_file = path
        temp_dir_ctx = tempfile.TemporaryDirectory()
        unpacked_dir = Path(temp_dir_ctx.name)
        try:
            with zipfile.ZipFile(path, "r") as zf:
                safe_extract(zf, unpacked_dir)
        except (zipfile.BadZipFile, ValueError, OSError) as e:
            _fail(f"cannot unpack {path}: {e}")
    else:
        if not path.is_dir():
            _fail(f"{path} is not a directory or a .pptx/.potx file")
        unpacked_dir = path
        if not (unpacked_dir / "ppt").is_dir():
            _fail(
                f"{path} has no ppt/ directory, so it is not an unpacked presentation. "
                "Point this at the directory you extracted the .pptx into."
            )

    validators = [PPTXSchemaValidator(unpacked_dir, original_file, verbose=args.verbose)]

    if args.auto_repair:
        total_repairs = sum(v.repair() for v in validators)
        if total_repairs:
            print(f"Auto-repaired {total_repairs} issue(s)")
            if packed_file is not None:
                rezip(unpacked_dir, packed_file)
                print(f"Wrote repaired file to {packed_file}")

    success = all([v.validate() for v in validators])

    if temp_dir_ctx is not None:
        temp_dir_ctx.cleanup()

    if success:
        print("All validations PASSED!")

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
