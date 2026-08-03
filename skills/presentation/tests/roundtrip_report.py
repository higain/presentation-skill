#!/usr/bin/env python3
"""
Prove what a LibreOffice round trip does and does not preserve in a deck.

    python3 tests/roundtrip_report.py build fixture.odp   # make a rich fixture
    python3 tests/roundtrip_report.py report deck.odp     # dump its properties
    python3 tests/roundtrip_report.py check fixture.odp   # build/report/diff all

``check`` round-trips the fixture three ways and diffs a property report before
and after each:

    .odp -> .odp     LibreOffice's own format, expected exact
    .pptx -> .pptx   OOXML in, OOXML out
    .odp -> .pptx    the export filter, i.e. what this skill does by default

Rerun this after a LibreOffice upgrade: the claims in SKILL.md are only true for
the version that produced them.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import uno  # noqa: E402

from office.impress import (  # noqa: E402
    add_image,
    add_slide,
    add_textbox,
    impress_document,
    load_document,
    masters,
    mm,
    new_document,
    notes_text,
    placeholder,
    pt,
    save_as,
    set_notes,
    shape_text,
    shapes,
    slides,
)

# --- property extraction --------------------------------------------------

# Character properties worth checking; these are the ones that silently change.
CHAR_PROPS = ("CharFontName", "CharHeight", "CharWeight", "CharPosture", "CharColor")
# Slide transition properties.
PAGE_PROPS = ("Change", "Effect", "Speed", "Duration", "TransitionType", "Visible")


def _get(obj, name, default=None):
    try:
        value = getattr(obj, name)
    except Exception:
        return default
    if hasattr(value, "value"):  # UNO enums
        return str(value.value)
    if isinstance(value, (int, float, str, bool)) or value is None:
        return value
    return str(value)


def _q(value: int) -> int:
    """Quantise a 1/100 mm measurement to 0.5 mm.

    Every format conversion re-rounds geometry (1/100 mm <-> EMU), so raw
    coordinates differ by a unit or two on any round trip.  That is noise; at
    this resolution only a real move shows up.
    """
    return int(round(value / 50.0))


def _first_portion(shape):
    """The first non-empty text portion of a shape, or None."""
    try:
        paragraphs = shape.createEnumeration()
    except Exception:
        return None
    while paragraphs.hasMoreElements():
        para = paragraphs.nextElement()
        try:
            portions = para.createEnumeration()
        except Exception:
            continue
        while portions.hasMoreElements():
            portion = portions.nextElement()
            if getattr(portion, "String", ""):
                return portion
    return None


def _shape_report(shape) -> dict:
    kind = shape.getShapeType()
    out = {
        "type": kind,
        "pos": [_q(shape.Position.X), _q(shape.Position.Y)],
        "size": [_q(shape.Size.Width), _q(shape.Size.Height)],
        "text": shape_text(shape).strip(),
    }
    # Read character formatting from the first text portion, not from the shape.
    # A .pptx import leaves shape-level character properties at their defaults
    # and carries the real values on the portions, so comparing shape-level
    # properties across formats reports losses that are not there.
    portion = _first_portion(shape)
    for name in CHAR_PROPS:
        value = _get(portion, name) if portion is not None else None
        if value is None:
            value = _get(shape, name)
        if value is not None:
            out[name] = value
    for name in ("FillColor", "FillStyle", "Effect", "Name"):
        value = _get(shape, name)
        if value is not None:
            out[name] = value

    if kind == "com.sun.star.drawing.GroupShape":
        out["children"] = [_shape_report(shape.getByIndex(i)) for i in range(shape.getCount())]
    elif kind == "com.sun.star.drawing.TableShape":
        model = getattr(shape, "Model", None)
        if model is not None:
            rows, cols = model.Rows.Count, model.Columns.Count
            out["table"] = [
                [model.getCellByPosition(c, r).getString() for c in range(cols)]
                for r in range(rows)
            ]
    elif kind == "com.sun.star.drawing.GraphicObjectShape":
        graphic = _get(shape, "Graphic")
        out["has_graphic"] = graphic is not None
    elif kind == "com.sun.star.drawing.OLE2Shape":
        out["clsid"] = _get(shape, "CLSID", "")
    return out


def report(path) -> dict:
    """Every property this test cares about, as a plain dict."""
    doc = load_document(path, read_only=True)
    try:
        props = doc.getDocumentProperties()
        out = {
            "slide_count": len(slides(doc)),
            "master_names": [m.Name for m in masters(doc)],
            "doc_title": props.Title,
            "doc_subject": props.Subject,
            "doc_keywords": list(props.Keywords),
            "slides": [],
        }
        for page in slides(doc):
            entry = {
                "name": page.Name,
                "layout": page.Layout,
                "master": page.MasterPage.Name,
                "size": [_q(page.Width), _q(page.Height)],
                "notes": notes_text(page),
                "shapes": [_shape_report(s) for s in shapes(page)],
            }
            for name in PAGE_PROPS:
                value = _get(page, name)
                if value is not None:
                    entry[name] = value
            out["slides"].append(entry)
        return out
    finally:
        doc.close(False)


# --- fixture --------------------------------------------------------------

TEMPLATE = Path("/usr/lib/libreoffice/share/template/common/presnt/Candy.otp")


def build(target: Path) -> Path:
    """A deck exercising everything the format supports that we claim about."""
    if TEMPLATE.exists():
        # Start from a real designed template: 11 masters, layouts, images.
        # It is root-owned and not writable, so work on a copy.
        seed = target.with_name("_seed.otp")
        seed.write_bytes(TEMPLATE.read_bytes())
        doc = load_document(seed)
    else:
        doc = new_document()

    try:
        pages = slides(doc)

        # 1. Placeholder text on the title slide, with explicit character formatting.
        title = placeholder(pages[0], "Title")
        if title is not None:
            title.setString("Quartalsbericht Q3")
            title.CharHeight = 40.0
            title.CharWeight = 150.0  # BOLD
            title.CharColor = 0x1E2761
        subtitle = placeholder(pages[0], "Subtitle")
        if subtitle is not None:
            subtitle.setString("Vertrieb & Marketing")

        # 2. Speaker notes.
        set_notes(pages[0], "Mit der Umsatzzahl aufmachen.\nZweite Notizzeile.")
        set_notes(pages[1], "Agenda kurz halten.")

        # 3. A slide transition.
        pages[1].Change = 1  # advance automatically
        pages[1].Duration = 3
        pages[1].Effect = uno.Enum(
            "com.sun.star.presentation.FadeEffect", "FADE_FROM_LEFT"
        )

        # 4. A hidden slide.
        pages[2].Visible = False

        # 5. A free text box with its own formatting.
        page = add_slide(doc, layout=20)
        box = add_textbox(doc, page, mm(20), mm(20), mm(200), mm(30), "Freies Textfeld")
        box.CharFontName = "Liberation Sans"
        box.CharHeight = 28.0
        box.CharColor = 0xB85042

        # 6. A shape animation (the legacy per-shape effect, which both
        #    formats can carry).
        box.Effect = uno.Enum("com.sun.star.presentation.AnimationEffect", "FADE_FROM_LEFT")

        # 7. A table.
        table = doc.createInstance("com.sun.star.drawing.TableShape")
        page.add(table)
        table.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", mm(20), mm(60)))
        table.setSize(uno.createUnoStruct("com.sun.star.awt.Size", mm(200), mm(40)))
        model = table.Model
        model.Rows.insertByIndex(model.Rows.Count, 3 - model.Rows.Count)
        model.Columns.insertByIndex(model.Columns.Count, 3 - model.Columns.Count)
        data = [["Region", "Umsatz", "Δ"], ["Nord", "1.234,50", "+5,2%"], ["Süd", "987,10", "-1,4%"]]
        for r, row in enumerate(data):
            for c, value in enumerate(row):
                model.getCellByPosition(c, r).setString(value)

        # 8. An embedded image.
        logo = _make_png()
        if logo is not None:
            add_image(doc, page, logo, mm(20), mm(110), mm(40), mm(40))

        # 9. A native chart (an OLE object, as in a real deck).
        chart = doc.createInstance("com.sun.star.drawing.OLE2Shape")
        page.add(chart)
        chart.setPosition(uno.createUnoStruct("com.sun.star.awt.Point", mm(80), mm(110)))
        chart.setSize(uno.createUnoStruct("com.sun.star.awt.Size", mm(120), mm(60)))
        chart.CLSID = "12DCAE26-281F-416F-A234-C3086127382E"  # chart

        # 10. Document properties.
        props = doc.getDocumentProperties()
        props.Title = "Quartalsbericht"
        props.Subject = "Q3"
        props.Keywords = ("Umsatz", "Vertrieb")

        doc.setModified(True)
        save_as(doc, target)
    finally:
        doc.close(False)
    return target


def _make_png():
    """A small PNG on disk, or None if Pillow is unavailable."""
    try:
        from PIL import Image
    except ImportError:
        return None
    path = Path.cwd() / "_fixture_logo.png"
    Image.new("RGB", (256, 256), (40, 90, 160)).save(path)
    return path


# --- driver ---------------------------------------------------------------


def _diff(before: dict, after: dict, label: str) -> bool:
    b = json.dumps(before, indent=2, ensure_ascii=False, sort_keys=True).splitlines()
    a = json.dumps(after, indent=2, ensure_ascii=False, sort_keys=True).splitlines()
    import difflib

    delta = [
        line
        for line in difflib.unified_diff(b, a, "before", "after", lineterm="", n=1)
        if line.startswith(("+", "-")) and not line.startswith(("+++", "---"))
    ]
    print(f"\n=== {label} ===")
    if not delta:
        print("IDENTICAL - nothing changed")
        return True
    print(f"{len(delta)} differing line(s):")
    for line in delta[:60]:
        print("  " + line)
    if len(delta) > 60:
        print(f"  ... and {len(delta) - 60} more")
    return False


def check(fixture: Path):
    build(fixture)
    print(f"Built {fixture} ({fixture.stat().st_size:,} bytes)")

    # Settle the fixture first.  A freshly authored deck is normalised on its
    # next load -- a table shape grows to the natural height of its rows -- and
    # that one-time adjustment is not round-trip loss.  Baselining after it
    # keeps the comparison about what a round trip actually costs.
    with impress_document(fixture):
        pass
    base = report(fixture)

    results = {}

    # .odp -> .odp: open and save in place.
    odp_copy = fixture.with_name("rt_odp.odp")
    odp_copy.write_bytes(fixture.read_bytes())
    with impress_document(odp_copy):
        pass
    results["odp -> odp (save in place)"] = _diff(
        base, report(odp_copy), "odp -> odp (save in place)"
    )

    # .odp -> .pptx: the default creation path.
    pptx = fixture.with_name("rt_export.pptx")
    with impress_document(fixture, save=False) as doc:
        save_as(doc, pptx)
    results["odp -> pptx (export filter)"] = _diff(
        base, report(pptx), "odp -> pptx (export filter)"
    )

    # .pptx -> .pptx: open the exported deck and save it again.
    pptx_rt = fixture.with_name("rt_pptx.pptx")
    pptx_rt.write_bytes(pptx.read_bytes())
    pptx_base = report(pptx_rt)
    with impress_document(pptx_rt):
        pass
    results["pptx -> pptx (save in place)"] = _diff(
        pptx_base, report(pptx_rt), "pptx -> pptx (save in place)"
    )

    print("\n=== summary ===")
    for label, ok in results.items():
        print(f"  {'PRESERVED' if ok else 'CHANGED  '}  {label}")
    return 0


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    command, target = sys.argv[1], Path(sys.argv[2])
    if command == "build":
        print(f"Built {build(target)}")
    elif command == "report":
        print(json.dumps(report(target), indent=2, ensure_ascii=False, sort_keys=True))
    elif command == "check":
        return check(target)
    else:
        print(f"Unknown command {command!r}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
