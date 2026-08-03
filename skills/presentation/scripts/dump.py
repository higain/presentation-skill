#!/usr/bin/env python3
"""
Read any presentation to markdown -- .odp and .pptx alike.

    python3 scripts/dump.py deck.odp
    python3 scripts/dump.py deck.pptx -o deck.md
    python3 scripts/dump.py deck.odp --no-notes --no-layout

Why this and not markitdown: markitdown has no ODF converter and refuses .odp
outright, and even on .pptx it reports text only -- no layout, no master, no
hidden-slide flag, and speaker notes only sometimes.  Those are exactly the
things you need to decide which template slide to reuse.

Output is one section per slide, in slide order.  Text is reported in z-order,
which is the order shapes were added, not top-to-bottom reading order.
"""

import argparse
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.impress import (  # noqa: E402
    PRESENTATION_EXTS,
    SofficeError,
    is_empty_placeholder,
    layout_slots,
    load_document,
    masters,
    notes_text,
    placeholder_kind,
    shape_text,
    shapes,
    slides,
)

# Shapes that hold text but say nothing useful when empty.
PLACEHOLDER_TYPES = {
    "com.sun.star.presentation.TitleTextShape": "title",
    "com.sun.star.presentation.SubtitleShape": "subtitle",
    "com.sun.star.presentation.OutlinerShape": "content",
    "com.sun.star.presentation.OutlineTextShape": "content",
    "com.sun.star.presentation.NotesShape": "notes",
}


def _table_rows(shape):
    """Rows of a table shape as lists of cell strings, or None."""
    model = getattr(shape, "Model", None)
    if model is None:
        return None
    try:
        rows, cols = model.Rows.Count, model.Columns.Count
    except AttributeError:
        return None
    out = []
    for r in range(rows):
        out.append([model.getCellByPosition(c, r).getString().strip() for c in range(cols)])
    return out


def _render_table(rows) -> list[str]:
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return lines


def _walk(shape, out: list[str], slots: list[str], depth: int = 0):
    """Emit one shape's content, recursing into groups.

    Empty placeholders are collected into ``slots`` instead of being dropped: on
    a template deck they are the whole point, since they say what the layout
    offers to fill in.
    """
    kind = shape.getShapeType()

    if kind == "com.sun.star.drawing.GroupShape":
        for i in range(shape.getCount()):
            _walk(shape.getByIndex(i), out, slots, depth + 1)
        return

    if kind == "com.sun.star.drawing.TableShape":
        rows = _table_rows(shape)
        if rows:
            out.extend(_render_table(rows))
        return

    if kind == "com.sun.star.drawing.GraphicObjectShape":
        title = (getattr(shape, "Title", "") or "").strip()
        out.append(f"[image{': ' + title if title else ''}]")
        return

    if kind == "com.sun.star.drawing.OLE2Shape":
        # Charts and embedded spreadsheets arrive as OLE objects.
        name = (getattr(shape, "CLSID", "") or "").strip()
        out.append(f"[embedded object{': ' + name if name else ''}]")
        return

    text = shape_text(shape).strip()
    label = PLACEHOLDER_TYPES.get(kind)
    if not text:
        # An empty placeholder is a slot the layout offers -- catch it by the
        # IsEmptyPresentationObject signal, not just the known text-shape types,
        # so a custom placeholder (e.g. an action title inherited from the
        # master) is reported instead of the slide reading as blank.
        if label:
            slots.append(label)
        elif is_empty_placeholder(shape):
            slots.append(placeholder_kind(shape))
        return
    prefix = f"**{label}:** " if label else ""
    for i, line in enumerate(text.splitlines()):
        line = line.strip()
        if line:
            out.append(f"{prefix if i == 0 else ''}{line}")


def dump(path: Path, notes: bool = True, layout: bool = True) -> str:
    doc = load_document(path, read_only=True)
    try:
        out = [f"# {path.name}", ""]
        pages = slides(doc)
        master_names = [m.Name for m in masters(doc)]
        out.append(
            f"{len(pages)} slide(s), {len(master_names)} master(s): "
            + ", ".join(master_names)
        )
        out.append("")

        for number, page in enumerate(pages, start=1):
            heading = f"## Slide {number}"
            if page.Name and page.Name != f"page{number}":
                heading += f" — {page.Name}"
            if not getattr(page, "Visible", True):
                heading += "  *(hidden)*"
            out.append(heading)

            if layout:
                out.append(
                    f"*layout {page.Layout}, master «{page.MasterPage.Name}»*"
                )
            out.append("")

            body: list[str] = []
            slots: list[str] = []
            for shape in shapes(page):
                _walk(shape, body, slots)
            # Slots the layout offers but that never land on the slide's own
            # page ("Click to add Title"): read them off the master.
            slots.extend(s["kind"] for s in layout_slots(page))
            if slots:
                counts = Counter(slots)
                body.append(
                    "*empty slots: "
                    + ", ".join(
                        f"{name}{f' ×{n}' if n > 1 else ''}" for name, n in counts.items()
                    )
                    + "*"
                )
            out.extend(body or ["*(no content)*"])
            out.append("")

            if notes:
                text = notes_text(page)
                if text:
                    out.append("**Notes:**")
                    out.extend(f"> {line}" for line in text.splitlines())
                    out.append("")

        return "\n".join(out).rstrip() + "\n"
    finally:
        doc.close(False)


def main():
    parser = argparse.ArgumentParser(
        description="Dump a presentation (.odp/.pptx/...) to markdown"
    )
    parser.add_argument("input", help="Presentation to read")
    parser.add_argument("-o", "--output", help="Write markdown here (default: stdout)")
    parser.add_argument(
        "--no-notes", action="store_true", help="Leave out speaker notes"
    )
    parser.add_argument(
        "--no-layout", action="store_true", help="Leave out layout/master lines"
    )
    args = parser.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"Error: {path} does not exist", file=sys.stderr)
        return 2
    if path.suffix.lower() not in PRESENTATION_EXTS:
        print(
            f"Error: {path} is not a presentation; expected one of: "
            f"{', '.join(sorted(PRESENTATION_EXTS))}",
            file=sys.stderr,
        )
        return 2

    try:
        text = dump(path, notes=not args.no_notes, layout=not args.no_layout)
    except (SofficeError, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
