"""
Native spreadsheet editing through LibreOffice's UNO API (python3-uno).

LibreOffice opens the real document and saves it back in its own format, so an
.ods stays ODF and an .xlsx stays OOXML -- no conversion, nothing dropped.
Formulas, styles, charts, named ranges and pivot tables survive an edit here
that a rewrite-from-scratch library would lose.

    import sys; sys.path.insert(0, "<skill>/scripts")
    from office.calc import calc_document

    with calc_document("budget.ods") as doc:
        sheet = doc.Sheets.getByName("Q3")
        sheet.getCellByPosition(1, 9).setFormula("=SUM(B2:B9)")   # B10, 0-based
    # recalculated and saved on exit

Cell access is 0-based by position, or A1-style by name:
``sheet.getCellRangeByName("B10")``.  ``setFormula`` takes formulas *and*
literals, ``setString``/``setValue`` force text/number.

One headless LibreOffice is started on first use and shut down when the process
ends, so many documents in one script cost one startup.
"""

import atexit
import contextlib
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue

from office.soffice import SOFFICE_MISSING, free_port, get_soffice_env, profile_arg, soffice_path

# com.sun.star.sheet.CellFlags.FORMULA / com.sun.star.sheet.FormulaResult.ERROR
CELLFLAG_FORMULA = 16
FORMULARESULT_ERROR = 4

# LibreOffice error codes worth naming; anything else is reported as Err:<code>.
ERROR_NAMES = {
    501: "#NAME?",  # invalid character
    502: "#VALUE!",  # invalid argument
    503: "#NUM!",  # invalid floating point operation
    508: "#VALUE!",  # missing pair (bracket)
    509: "#VALUE!",  # missing operator
    510: "#VALUE!",  # missing variable
    511: "#VALUE!",  # missing variable
    519: "#VALUE!",
    522: "#REF!",  # circular reference
    523: "#NUM!",  # calculation does not converge
    524: "#REF!",
    525: "#NAME?",
    526: "#NAME?",
    527: "#REF!",  # nesting too deep
    532: "#DIV/0!",
    533: "#NAME?",  # nested array not supported
    538: "#VALUE!",
    539: "#VALUE!",
    540: "#VALUE!",  # external add-in error
    32767: "#N/A",
}


class SofficeError(RuntimeError):
    pass


def column_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_name(sheet_name: str, column: int, row: int) -> str:
    return f"{sheet_name}!{column_letter(column)}{row + 1}"


class _Server:
    """A headless soffice listening on a private loopback port."""

    def __init__(self):
        self.proc = None
        self.ctx = None
        self.profile = None

    def desktop(self, timeout: float = 60.0):
        if self.ctx is None:
            self._start(timeout)
        return self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

    def _start(self, timeout: float):
        port = free_port()
        self.profile = tempfile.mkdtemp(prefix="lo-uno-profile-")
        try:
            exe = soffice_path()
        except FileNotFoundError:
            raise SofficeError(SOFFICE_MISSING) from None

        self.proc = subprocess.Popen(
            [
                exe,
                "--headless",
                "--invisible",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                "--nologo",
                profile_arg(self.profile),
                f"--accept=socket,host=127.0.0.1,port={port};urp;",
            ],
            env=get_soffice_env(),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        atexit.register(self.stop)

        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local
        )
        url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        deadline = time.monotonic() + timeout
        while True:
            if self.proc.poll() is not None:
                self.stop()
                raise SofficeError(
                    f"LibreOffice exited with code {self.proc.returncode} before accepting "
                    "a connection"
                )
            try:
                self.ctx = resolver.resolve(url)
                return
            except Exception:  # NoConnectException until the listener is up
                if time.monotonic() > deadline:
                    self.stop()
                    raise SofficeError(
                        f"LibreOffice did not accept a UNO connection within {timeout:g}s"
                    ) from None
                time.sleep(0.25)

    def stop(self):
        if self.ctx is not None:
            with contextlib.suppress(Exception):
                self.ctx.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.ctx
                ).terminate()
            self.ctx = None
        if self.proc is not None:
            try:
                self.proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                self.proc.kill()
                self.proc.wait(timeout=10)
            self.proc = None
        if self.profile:
            shutil.rmtree(self.profile, ignore_errors=True)
            self.profile = None


_SERVER = _Server()


def shutdown():
    """Stop the shared LibreOffice.  Called automatically at process exit."""
    _SERVER.stop()


def kill():
    """Hard-kill the shared LibreOffice, for when it stops responding."""
    proc, _SERVER.proc, _SERVER.ctx = _SERVER.proc, None, None
    if proc is not None and proc.poll() is None:
        proc.kill()


def _props(**kwargs):
    out = []
    for name, value in kwargs.items():
        p = PropertyValue()
        p.Name = name
        p.Value = value
        out.append(p)
    return tuple(out)


def load_document(path, read_only: bool = False, update_links: bool = False, timeout: float = 60.0):
    """Open a spreadsheet and return the UNO document.  Caller must close it."""
    path = Path(path).absolute()
    if not path.exists():
        raise FileNotFoundError(f"{path} does not exist")
    if not os.access(path, os.R_OK):
        raise SofficeError(f"{path} is not readable")
    if not read_only and not os.access(path, os.W_OK):
        raise SofficeError(
            f"{path} is not writable, so it cannot be opened for editing. "
            "Copy it first, or open it with read_only=True."
        )

    options = {
        "Hidden": True,
        "MacroExecutionMode": 1,  # NEVER_EXECUTE
        # 0 = NO_UPDATE: keep the cached values of links to other files
        # instead of trying to resolve files that are probably absent.
        "UpdateDocMode": 3 if update_links else 0,  # FULL_UPDATE / NO_UPDATE
    }
    # Pass ReadOnly only to turn it ON.  Explicitly asking for ReadOnly=False on
    # a file the process cannot write makes type detection fail outright, and
    # LibreOffice reports it as the useless "Unsupported URL ... type detection
    # failed" rather than a permission error.
    if read_only:
        options["ReadOnly"] = True

    desktop = _SERVER.desktop(timeout=timeout)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(str(path)), "_blank", 0, _props(**options)
    )
    if doc is None:
        raise SofficeError(f"LibreOffice could not open {path} as a document")
    if not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        doc.close(False)
        raise SofficeError(f"{path} is not a spreadsheet")
    return doc


@contextlib.contextmanager
def calc_document(path, save: bool = True, recalc: bool = True, update_links: bool = False):
    """Open a spreadsheet; recalculate and save it in place on a clean exit.

    ``save=False`` opens read-only -- use it for inspection, and whenever you
    must not touch the file.  An exception inside the block leaves the file on
    disk untouched.
    """
    doc = load_document(path, read_only=not save, update_links=update_links)
    try:
        yield doc
        if recalc:
            doc.calculateAll()
        if save:
            doc.setModified(True)  # store() is a no-op on an unmodified document
            doc.store()
    finally:
        with contextlib.suppress(Exception):
            doc.close(False)


def number_format(doc, code: str, language: str = "en", country: str = "US") -> int:
    """Key for a number format code, registered on the document if it is new.

        cells = sheet.getCellRangeByName("B2:B20")
        cells.NumberFormat = number_format(doc, "#,##0.00")

    Write the code the English way -- ``.`` decimal, ``,`` thousands, ``$#,##0``
    -- whatever the machine's locale is; LibreOffice renders it in the user's
    locale.  Passing a locale-specific code with the default en-US argument is
    the usual reason a format silently does not apply.
    """
    locale = uno.createUnoStruct("com.sun.star.lang.Locale")
    locale.Language = language
    locale.Country = country
    formats = doc.getNumberFormats()
    key = formats.queryKey(code, locale, False)
    if key == -1:
        key = formats.addNew(code, locale)
    return key


def add_comment(sheet, cell_ref: str, text: str):
    """Attach a cell comment (an annotation), e.g. add_comment(sheet, "B7", "...").

    Use it to record where a hardcoded number came from.  LibreOffice writes it
    to .ods and .xlsx alike (openpyxl reads it back as ``cell.comment``); only
    reading its text back through UNO from an .xlsx comes up empty, which is a
    quirk of the reader, not a lost comment.
    """
    target = sheet.getCellRangeByName(cell_ref).CellAddress
    sheet.Annotations.insertNew(target, text)


def sheets(doc):
    """The document's sheets, in order."""
    return [doc.Sheets.getByIndex(i) for i in range(doc.Sheets.Count)]


def used_range(sheet):
    """(last_column, last_row) of the used area, 0-based; (-1, -1) if empty."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    addr = cursor.RangeAddress
    if addr.EndColumn == 0 and addr.EndRow == 0 and not sheet.getCellByPosition(0, 0).getString():
        return -1, -1
    return addr.EndColumn, addr.EndRow


def _range_cell_count(ranges) -> int:
    total = 0
    for addr in ranges.getRangeAddresses():
        total += (addr.EndColumn - addr.StartColumn + 1) * (addr.EndRow - addr.StartRow + 1)
    return total


def sheet_formula_count(sheet) -> int:
    """How many cells on one sheet hold a formula.  Counts whole ranges, so it
    stays fast on a sheet with tens of thousands of formulas."""
    return _range_cell_count(sheet.queryContentCells(CELLFLAG_FORMULA))


def count_formulas(doc) -> int:
    """How many cells in the document hold a formula."""
    return sum(sheet_formula_count(sheet) for sheet in sheets(doc))


def error_cells(doc):
    """Every formula cell whose result is an error.

    Returns dicts: ``cell`` (Sheet!A1), ``error`` (#NAME?, ...), ``code``
    (LibreOffice error number) and ``formula``.
    """
    found = []
    for sheet in sheets(doc):
        ranges = sheet.queryFormulaCells(FORMULARESULT_ERROR)
        enumeration = ranges.getCells().createEnumeration()
        while enumeration.hasMoreElements():
            cell = enumeration.nextElement()
            addr = cell.CellAddress
            code = cell.getError()
            display = (cell.getString() or "").strip()
            name = display if display.startswith("#") else ERROR_NAMES.get(code, f"Err:{code}")
            found.append(
                {
                    "cell": cell_name(sheet.Name, addr.Column, addr.Row),
                    "error": name,
                    "code": code,
                    "formula": cell.getFormula(),
                }
            )
    return found
