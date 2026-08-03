#!/usr/bin/env python3
"""
Recalculate every formula in a spreadsheet with LibreOffice and report the
cells that end up as errors.

Works on any format Calc round-trips -- .ods, .fods, .xlsx, .xlsm, .xltx,
.ots -- and always saves in the file's own format.  The file is rewritten in
place; the report goes to stdout as JSON.

    python3 scripts/recalc.py book.ods [timeout_seconds] [--force]
"""

import json
import os
import sys
import threading
from pathlib import Path

import uno

from office.calc import (
    SofficeError,
    count_formulas,
    error_cells,
    kill,
    load_document,
    shutdown,
)

MAX_LOCATIONS = 100
MAX_DETAILS = 20

SUPPORTED = {".ods", ".fods", ".ots", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".sxc"}


def _external_links(doc):
    """Every other document this one pulls values from, and whether those
    values are still available: the linked file itself, or failing that the
    copy of its data cached inside this file."""
    state = []
    links = getattr(doc, "ExternalDocLinks", None)
    for url in links.ElementNames if links else ():
        link = links.getByName(url)
        cached_rows = sum(len(link.getByIndex(i).getAllRows()) for i in range(link.Count))
        try:
            path = uno.fileUrlToSystemPath(url)
            exists = Path(path).exists()
        except Exception:  # a non-file URL (http, smb): treat as unresolvable here
            path, exists = url, False
        state.append(
            {"url": url, "path": path, "resolved": exists, "cached_rows": cached_rows}
        )
    return state


def _summarize(errors):
    summary = {}
    for entry in errors:
        bucket = summary.setdefault(entry["error"], {"count": 0, "locations": []})
        bucket["count"] += 1
        bucket["locations"].append(entry["cell"])
    for bucket in summary.values():
        extra = len(bucket["locations"]) - MAX_LOCATIONS
        if extra > 0:
            bucket["locations"] = bucket["locations"][:MAX_LOCATIONS]
            bucket["locations_truncated"] = extra
    return summary


def recalc(filename, force=False):
    path = Path(filename)
    if not path.exists():
        return {"error": f"{filename} does not exist"}
    if path.suffix.lower() not in SUPPORTED:
        return {
            "error": (
                f"{path.suffix or path.name} is not a spreadsheet format that stores formulas "
                f"(supported: {', '.join(sorted(SUPPORTED))}). A .csv holds values only."
            )
        }
    if not os.access(path, os.W_OK):
        return {"error": f"{filename} is not writable; recalculation rewrites the file in place"}

    try:
        doc = load_document(path)
    except SofficeError as e:
        return {"error": str(e)}

    try:
        before = path.stat().st_mtime_ns

        links = _external_links(doc)
        lost = [link for link in links if not link["resolved"] and not link["cached_rows"]]
        if lost and not force:
            return {
                "error": (
                    f"Refusing to recalculate: this file pulls values from {len(lost)} document(s) "
                    "that are missing from disk, and it no longer holds a cached copy of their "
                    "data -- so every formula reading them would silently recalculate to 0. "
                    "(Rewriting a workbook with pandas, or with a tool that drops the link cache, "
                    "does exactly this.) Restore the linked file next to this one, or pass --force "
                    "to write the zeros anyway. The file on disk is unchanged."
                ),
                "external_links": lost,
            }

        doc.calculateAll()
        errors = error_cells(doc)
        total_formulas = count_formulas(doc)

        doc.setModified(True)
        doc.store()

        if path.stat().st_mtime_ns == before:
            return {
                "error": (
                    "LibreOffice reported no error but never rewrote the file, so nothing was "
                    "saved. Check the file is not open elsewhere, then retry."
                )
            }
    except SofficeError as e:
        return {"error": str(e)}
    except Exception as e:  # UNO exceptions do not share a common base class
        return {"error": f"{type(e).__name__}: {e}"}
    finally:
        try:
            doc.close(False)
        except Exception:
            pass

    result = {
        "status": "success" if not errors else "errors_found",
        "file": str(path),
        "total_formulas": total_formulas,
        "total_errors": len(errors),
    }
    if errors:
        result["error_summary"] = _summarize(errors)
        result["errors"] = errors[:MAX_DETAILS]
        if len(errors) > MAX_DETAILS:
            result["errors_truncated"] = len(errors) - MAX_DETAILS
    if links:
        result["external_links"] = links
    return result


def run_with_timeout(fn, timeout):
    """Run fn in a worker thread; a hung LibreOffice cannot wedge the script."""
    box = {}
    worker = threading.Thread(target=lambda: box.setdefault("result", fn()), daemon=True)
    worker.start()
    worker.join(timeout)
    if worker.is_alive():
        kill()  # a wedged soffice would otherwise outlive us
        return {
            "error": (
                f"LibreOffice did not finish within {timeout}s; the file was NOT modified. "
                "Re-run with a longer timeout, e.g. recalc.py file.ods 120"
            )
        }
    return box.get("result", {"error": "recalculation produced no result"})


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv[1:]

    if not args:
        print(__doc__.strip())
        print("\nJSON out: status (success|errors_found), total_formulas, total_errors,")
        print("error_summary per error type, errors (first 20, with the offending formula).")
        print("On failure the JSON has an 'error' key and no 'status', and the exit code is 1.")
        sys.exit(1)

    timeout = int(args[1]) if len(args) > 1 else 60
    result = run_with_timeout(lambda: recalc(args[0], force=force), timeout)
    print(json.dumps(result, indent=2))

    shutdown()
    sys.exit(1 if "error" in result else 0)


if __name__ == "__main__":
    main()
