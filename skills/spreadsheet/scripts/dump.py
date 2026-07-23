"""
Print a spreadsheet as markdown, with cell coordinates -- so you can plan edits
from what you read.  Reads every format Calc opens: .ods, .fods, .xlsx, .xlsm,
.xls, .csv.

    python3 scripts/dump.py book.ods
    python3 scripts/dump.py book.ods --sheet "Q3" --max-rows 200
    python3 scripts/dump.py book.ods --formulas    # formulas in the grid

Cells show what LibreOffice displays (formatted dates, rounded currency, and
formula *results*), not the raw stored numbers.  Formula cells are also listed
verbatim under each sheet.  The file is opened read-only and never modified.
"""

import argparse
import sys

from office.calc import (
    CELLFLAG_FORMULA,
    SofficeError,
    calc_document,
    column_letter,
    sheet_formula_count,
    sheets,
    shutdown,
    used_range,
)

MAX_FORMULAS = 200


def _escape(text):
    return text.replace("|", "\\|").replace("\n", " ")


def _grid(sheet, last_col, last_row, max_rows, max_cols, formulas):
    cols = min(last_col + 1, max_cols)
    rows = min(last_row + 1, max_rows)

    header = ["", *(column_letter(c) for c in range(cols))]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]
    for r in range(rows):
        cells = []
        for c in range(cols):
            cell = sheet.getCellByPosition(c, r)
            text = cell.getFormula() if formulas else cell.getString()
            cells.append(_escape(text))
        lines.append(f"| **{r + 1}** | " + " | ".join(cells) + " |")

    if rows <= last_row:
        lines.append(f"\n_rows {rows + 1}-{last_row + 1} not shown (raise --max-rows)_")
    if cols <= last_col:
        lines.append(
            f"\n_columns {column_letter(cols)}-{column_letter(last_col)} not shown "
            "(raise --max-cols)_"
        )
    return lines


def _formula_list(sheet, limit=MAX_FORMULAS):
    """Up to `limit` formulas, and the true total.  Reading a cell at a time is
    slow across the UNO bridge, so a sheet with 20k formulas is counted whole
    and only the first few are spelled out."""
    total = sheet_formula_count(sheet)
    out = []
    enumeration = sheet.queryContentCells(CELLFLAG_FORMULA).getCells().createEnumeration()
    while enumeration.hasMoreElements() and len(out) < limit:
        cell = enumeration.nextElement()
        addr = cell.CellAddress
        out.append(
            f"- `{column_letter(addr.Column)}{addr.Row + 1}`: `{cell.getFormula()}`"
            f"  -> {cell.getString()}"
        )
    return out, total


def dump(path, sheet_name=None, max_rows=100, max_cols=30, formulas=False):
    lines = []
    with calc_document(path, save=False, recalc=False) as doc:
        selected = sheets(doc)
        if sheet_name is not None:
            selected = [s for s in selected if s.Name == sheet_name]
            if not selected:
                names = ", ".join(repr(s.Name) for s in sheets(doc))
                raise SystemExit(f"No sheet named {sheet_name!r}. Sheets: {names}")

        lines.append(f"# {path} ({doc.Sheets.Count} sheet(s))")
        for sheet in selected:
            last_col, last_row = used_range(sheet)
            if last_col < 0:
                lines.append(f"\n## {sheet.Name} — empty")
                continue
            extent = f"A1:{column_letter(last_col)}{last_row + 1}"
            lines.append(f"\n## {sheet.Name} — used range {extent}")
            lines += _grid(sheet, last_col, last_row, max_rows, max_cols, formulas)

            found, total = _formula_list(sheet)
            if found and not formulas:
                lines.append(f"\n### formulas in {sheet.Name} ({total})")
                lines += found
                if total > len(found):
                    lines.append(f"_...and {total - len(found)} more_")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("file")
    parser.add_argument("--sheet", default=None, help="only this sheet (default: all)")
    parser.add_argument("--max-rows", type=int, default=100)
    parser.add_argument("--max-cols", type=int, default=30)
    parser.add_argument(
        "--formulas", action="store_true", help="show formulas in the grid instead of results"
    )
    args = parser.parse_args()

    try:
        print(dump(args.file, args.sheet, args.max_rows, args.max_cols, args.formulas))
    except (SofficeError, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        shutdown()


if __name__ == "__main__":
    main()
