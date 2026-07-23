---
name: maxi-spreadsheet
description: 'Use this skill whenever a spreadsheet file is the primary input or output - .ods, .fods, .xlsx, .xlsm, .xls, .csv or .tsv. Covers opening, reading, editing or fixing an existing spreadsheet (columns, formulas, formatting, charts, messy data), creating one from scratch or from other data, and converting between tabular formats. Trigger on whatever the user calls it - "Excel file", "workbook", "spreadsheet", "Tabelle", "Tabellenkalkulation", "Calc file", "ODS" or just "the sheet" - whatever the extension turns out to be: people say "Excel" for an .ods and "spreadsheet" for an .xlsx. Trigger especially when they point at such a file by name or path, even casually ("the excel in my downloads", "meine Kalkulation"). Also trigger for cleaning messy tabular data into a proper spreadsheet, and for budgets, financial models and price lists. The deliverable must be a spreadsheet file. Do NOT trigger when the deliverable is a Word document, HTML report, standalone script, database pipeline or Google Sheets API.'
license: Proprietary. LICENSE.txt has complete terms
compatibility: Linux with LibreOffice (soffice) and python3-uno, plus python3-openpyxl, python3-pandas and python3-odf from apt. No pip or npm. markitdown.sh is optional.
---

# Spreadsheet creation, editing, and analysis — .ods and .xlsx

| Task | Approach |
|---|---|
| **Read** any format, with cell coordinates | `python3 scripts/dump.py file.ods` |
| **Edit an existing file** — .ods, or any file whose formatting must survive | `office.calc`, the LibreOffice API — see below |
| **Build a new .xlsx** from scratch | `openpyxl`, then `recalc.py` |
| **Build a new .ods** from scratch | `openpyxl` → `convert.py file.xlsx ods` → `recalc.py` |
| **Bulk data** in or out | `pandas` — `read_excel(f, engine="odf")` for .ods, default engine for .xlsx |
| **Convert** between formats | `python3 scripts/convert.py in.xlsx out.ods` |
| **Verify** — mandatory whenever the file has formulas | `python3 scripts/recalc.py file` |

> Script paths are relative to this skill's directory. Everything is preinstalled from apt —
> **there is no pip and no npm here.** If something is genuinely missing, ask the user to
> `sudo apt install` it; never try to install a Python package yourself.

**"Excel file" is not a format.** Most people here work in LibreOffice and call every spreadsheet
an Excel file (or *Tabelle*) whatever it is. The extension on disk decides which engine you use
and which formula syntax is legal, so look at the actual filename before you choose — and when you
create a file, deliver the format the user's other files are in unless they asked for another one.
Keep the original extension when editing: never quietly hand back an .xlsx for an .ods.

**Which engine.** `office.calc` drives a headless LibreOffice, so it edits the real document and
saves it in its own format — an .ods stays ODF, and charts, pivot tables, conditional formats and
styles you did not touch come through untouched (verified). Use it for **every edit of an existing
file**, and for anything .ods. `openpyxl` writes .xlsx bytes directly: faster for building a big
sheet from nothing, and the right tool for styling .xlsx, but it **cannot read or write .ods at
all**, and it rebuilds a file from what it understands — anything it does not model is dropped.

## Requirements for every output

- **Professional font** (Arial, Liberation Sans, Times New Roman) throughout, unless the user says otherwise.
- **Zero formula errors.** Never ship while `recalc.py` reports `errors_found`. If you think an error predates you, prove it: open the *original* with `dump.py` and look at that cell. An error you introduced looks exactly like one you inherited.
- **Use formulas, never hardcoded results.** Write `=SUM(B2:B9)`, not the Python-computed total. The sheet must recalculate when its inputs change.
- **Follow the user's spec literally.** Exact tab names, exact column headers, and the formula they spelled out. A redesign that computes something else fails, however elegant.
- **Document every assumption and hardcoded number** where the reader will see it — a cell comment, or an adjacent cell at a table's end. Cite a real source when one exists (`Quelle: Geschäftsbericht 2024, S. 45`); when the number came from the user, say so plainly.
- **A workbook *you create* for someone to fill in** needs a short legend naming which cells to edit, and one example row of realistic values showing the expected format. Never add such a row to a file you were asked to edit.
- **Editing an existing file: match its conventions exactly.** They override every guideline here. Find its designated input cells first — a distinct font color, fill, or shading marks them — write only there, and leave every existing formula untouched.

## Editing with office.calc (LibreOffice)

Put the skill's `scripts` directory on the path — the one that *contains*
`office/`, never `office/` itself, which is what produces
`ModuleNotFoundError: No module named 'office'`:

```python
import sys; sys.path.insert(0, "/absolute/path/to/spreadsheet/scripts")
from office.calc import calc_document, number_format

with calc_document("budget.ods") as doc:          # save=False to inspect without saving
    sheet = doc.Sheets.getByName("Q3")
    sheet.getCellRangeByName("A10").setString("Summe")
    sheet.getCellRangeByName("B10").setFormula("=SUM(B2:B9)")
    sheet.getCellRangeByName("B10").NumberFormat = number_format(doc, "#,##0.00")
# recalculated and saved in place on exit; an exception leaves the file untouched
```

`getCellRangeByName("B10")` is A1-style, `getCellByPosition(1, 9)` is 0-based (column, row).
`setValue(3.5)` writes a number, `setString("007")` forces text, `setFormula` takes either plus
formulas. Useful properties on any cell or range: `CharFontName`, `CharHeight`, `CharWeight`
(`150.0` = bold), `CharColor`, `CellBackColor` (`0xRRGGBB`), `HoriJustify`, `NumberFormat`.
Also `sheet.Columns.getByIndex(0).Width` (1/100 mm), `doc.CurrentController.freezeAtPosition(col,
row)`, and `add_comment(sheet, "B7", "Quelle: ...")` from the same module for a cell comment.

**Formula syntax through `setFormula` is not Excel syntax.** These four rules are absolute:

- **Separate arguments with `;`** — `=ROUND(B2;2)`. A comma silently means something else and
  yields `Err:508`/`Err:509`, always on .ods and often on .xlsx.
- **Cross-sheet references use a dot** — `=Assumptions.B1`, or `=$'Meine Annahmen'.B1` when the
  name has a space. Excel's `Assumptions!B1` is a `#NAME?` error here.
- **Decimal literals use `.`** regardless of the machine's locale (`setFormula("1.5")`); `"1,5"`
  is stored as text. For plain numbers prefer `setValue`.
- **No `_xlfn.` prefixes.** Write `=TEXTJOIN(...)`; LibreOffice adds the prefix itself when it
  exports to .xlsx.

**A formula LibreOffice could not parse is stored lowercased** — a quick tell next to a `#NAME?`.

**Functions this LibreOffice evaluates** (25.2): everything Excel-2007-era, plus `TEXTJOIN`,
`CONCAT`, `IFS`, `SWITCH`, `MAXIFS`, `MINIFS`, `XLOOKUP`, `XMATCH`, `LET`, and the dynamic-array
functions `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`. **`LAMBDA` and `TEXTSPLIT` do not exist here**
and become `#NAME?`. A dynamic-array function only fills one cell via `setFormula`; to spill it,
call `setArrayFormula` on the whole target range:
`sheet.getCellRangeByName("H1:H20").setArrayFormula("=FILTER(B1:B99;B1:B99>10)")`.

## Editing with openpyxl (.xlsx only)

- **Formulas are strings written verbatim into the XML**, so here it *is* Excel syntax: commas between arguments, `='Assumptions Inputs'!$B$5` across sheets (quote any name with a space).
- **Post-2007 functions need the `_xlfn.` prefix** when openpyxl writes them: `_xlfn.TEXTJOIN`, `_xlfn.CONCAT`, `_xlfn.IFS`, `_xlfn.SWITCH`, `_xlfn.MAXIFS`, `_xlfn.MINIFS`. Bare, each yields `#NAME?`.
- **Never write `XLOOKUP`, `FILTER`, `SORT`, `UNIQUE` or `SEQUENCE` through openpyxl.** They are spilling functions and an openpyxl-written file has no spill metadata, so only the top-left cell gets a value — and `recalc.py` reports `total_errors: 0` on the truncated result. Use `INDEX`/`MATCH`, and sort, filter and de-duplicate in Python before writing cells.
- **Reading a model takes two loads.** `data_only=True` yields cached values with the formulas gone; the default yields formula strings with no values. One pass cannot give you both.
- **`data_only=True` is destructive if you save.** That workbook has no formulas left, so saving replaces every one with a literal — permanently.
- **`data_only=True` on a file openpyxl just wrote returns `None` everywhere** — run `recalc.py` first. (A formula whose result is `""` also reads back as `None`.)
- **Merged cells: write the top-left anchor only.** Every other cell in the range is a `MergedCell` whose `.value` is read-only.
- **`.xlsm` loses its macros unless you pass `keep_vba=True`** to `load_workbook`. A LibreOffice round-trip is not proven to keep them either — for a macro workbook, ask the user before touching it.
- **Do the styling of an .xlsx in openpyxl, not in office.calc.** Setting `CharColor` on a cell that still carries openpyxl's default theme font is dropped by LibreOffice's .xlsx export (a real bug, .ods is unaffected). Fills, fonts and number formats set *by openpyxl* survive a recalc round-trip intact.

## Recalculate — mandatory whenever the file contains formulas

```bash
python3 scripts/recalc.py output.ods [timeout_seconds] [--force]   # default 60
```

LibreOffice computes every formula, the file is **rewritten in place**, and you get JSON:
`status` (`success` | `errors_found`), `total_formulas`, `total_errors`, an `error_summary` per
error type (up to 100 cells each, `locations_truncated` says how many it withheld) and `errors`,
the first 20 offenders **with the formula that failed**. Fix what it names and run it again.
**JSON with an `error` key instead of a `status` means nothing was recalculated**, and only that
case exits non-zero — `errors_found` exits 0, so never treat a clean exit as a clean workbook.

An openpyxl-written file has no cached values at all until this runs: until then every formula
cell reads back as `None` to `pandas`, `data_only=True`, and most previewers.

**A green recalc proves your formulas *evaluate*, not that they are *right*.** An off-by-one range
or a reference to the wrong row yields a clean, error-free file with wrong numbers. Write 2–3
formulas first, `dump.py` them, and check they pull the values you expect before building a grid.

**Files that link to another workbook.** `recalc.py` reports them under `external_links` and
refuses to run when a linked file is missing *and* the cached copy of its data is gone — that
combination recalculates to silent zeros, not to an error. Restore the linked file beside it, or
`--force` to accept the zeros. Editing through `office.calc` never resolves external links (it
keeps their cached values); rewriting such a file with pandas is what destroys them.

## Financial models

Unless the user says otherwise, or the existing file already does something else.

**Color:** blue text (`0x0000FF`) for hardcoded inputs and scenario levers · black for formulas ·
green (`0x008000`) for links to another sheet · red (`0xFF0000`) for links to another file ·
yellow fill (`0xFFFF00`) for key assumptions and cells the user should fill in.

**Numbers:** write format codes the English way — `#,##0.00`, `$#,##0` — whatever the machine's
locale; LibreOffice renders them in the user's locale (so a German user sees `1.234,50`). Name the
unit in the header (`Umsatz (T€)`) · zeros render as `-` (`#,##0;(#,##0);-`) · negatives in
parentheses · percentages `0.0%`, **stored as fractions** (`0.15` renders `15.0%`; storing `15`
renders `1500.0%`) · valuation multiples `0.0x` · years as text (`"2024"`, never `2,024`).

**Structure:** every assumption in its own labeled cell, referenced by the formulas that use it
(`=B5*(1+$B$6)`, never `=B5*1.05`) · formulas consistent across every projection period, since a
lone edited cell mid-row is the commonest silent error · guard denominators that can be zero.

## Reading a file

`python3 scripts/dump.py file.ods` prints each sheet as a markdown grid **with row numbers and
column letters**, plus every formula it contains — the only reader here that works on .ods and the
only one that gives you coordinates to plan an edit from. `--sheet NAME`, `--max-rows`,
`--max-cols`, `--formulas` (formulas inside the grid) narrow it.

Formulas print in LibreOffice syntax whatever the file's format, so a cross-sheet reference in an
.xlsx reads `=$'Backend API'.B4`, not Excel's `!`. That is the same formula, correctly read — not
damage, and nothing to "fix". The syntax rules are under
[Editing with office.calc](#editing-with-officecalc-libreoffice).

`markitdown.sh book.xlsx /tmp/book.md` (writes to the file you name, second argument required) is
an alternative quick look for **.xlsx only** — it runs in a container and rejects .ods. It is a
shell script already on `PATH`, *not* a file in this skill's `scripts/`: run it as
`markitdown.sh ...`, never `python3 .../scripts/markitdown.sh`. It renders
each sheet as one plain table with no cell coordinates and no formulas, drops empty leading
columns, and shows cached values, so a file openpyxl just wrote looks empty until `recalc.py`
runs. Prefer `dump.py` unless you specifically want markitdown's flatter output.

## Dependencies

All from apt, all present: `libreoffice-calc` + `python3-uno` (the engine behind every script
here), `python3-openpyxl`, `python3-pandas`, `python3-odf` (lets pandas read .ods).
`markitdown.sh` additionally needs docker and the `markitdown:latest` image.
