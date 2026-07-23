"""
Convert a spreadsheet between formats with LibreOffice, recalculating first so
the result never carries stale or missing cached values.

    python3 scripts/convert.py model.xlsx model.ods     # explicit destination
    python3 scripts/convert.py model.xlsx ods           # alongside the source
    python3 scripts/convert.py model.ods report.pdf

Targets: ods, fods, xlsx, xlsm, xls, csv, pdf, html.  Converting to .csv writes
the first sheet only.

Conversion is the safe way to *create* a new .ods (build it with openpyxl, then
convert), but not a way to *edit* an existing one -- a round trip through a
foreign format is lossy.  Edit .ods in place with office.calc instead.
"""

import sys
from pathlib import Path

import uno

from office.calc import SofficeError, load_document, shutdown

FILTERS = {
    ".ods": "calc8",
    ".fods": "OpenDocument Spreadsheet Flat XML",
    ".ots": "calc8_template",
    ".xlsx": "Calc MS Excel 2007 XML",
    ".xlsm": "Calc MS Excel 2007 VBA XML",
    ".xltx": "Calc MS Excel 2007 XML Template",
    ".xls": "MS Excel 97",
    ".csv": "Text - txt - csv (StarCalc)",
    ".pdf": "calc_pdf_Export",
    ".html": "HTML (StarCalc)",
}

# comma-separated, double-quoted, UTF-8, from row 1, default formats, and the
# last token false: write the stored value (1000), not the formatted one ($1.000)
CSV_OPTIONS = "44,34,76,1,,0,false,true,false"


def _props(**kwargs):
    from com.sun.star.beans import PropertyValue

    out = []
    for name, value in kwargs.items():
        p = PropertyValue()
        p.Name = name
        p.Value = value
        out.append(p)
    return tuple(out)


def convert(source, target, recalc=True):
    source = Path(source).absolute()
    target = Path(target)
    if not target.suffix:  # a bare extension: "ods"
        target = source.with_suffix("." + str(target).lstrip("."))
    target = target.absolute()

    suffix = target.suffix.lower()
    if suffix not in FILTERS:
        raise SofficeError(
            f"Cannot write {suffix or target.name}; supported: {', '.join(sorted(FILTERS))}"
        )
    if target == source:
        raise SofficeError("Source and destination are the same file")
    target.parent.mkdir(parents=True, exist_ok=True)

    # Read-only: conversion writes the target with storeToURL and never touches
    # the source, so a source the user cannot write is still fair game.
    doc = load_document(source, read_only=True)
    try:
        if recalc:
            doc.calculateAll()
        options = {"FilterName": FILTERS[suffix], "Overwrite": True}
        if suffix == ".csv":
            options["FilterOptions"] = CSV_OPTIONS
        doc.storeToURL(uno.systemPathToFileUrl(str(target)), _props(**options))
    finally:
        try:
            doc.close(False)
        except Exception:
            pass
    return target


def main():
    if len(sys.argv) != 3:
        print(__doc__.strip())
        sys.exit(1)
    try:
        target = convert(sys.argv[1], sys.argv[2])
    except (SofficeError, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:  # UNO exceptions do not share a common base class
        print(f"Error: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        shutdown()
    print(f"Wrote {target}")


if __name__ == "__main__":
    main()
