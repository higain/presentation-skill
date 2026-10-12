#!/usr/bin/env python3
"""Name a template's designs from a worksheet.

    python3 scripts/name_designs.py templates/acme/catalog            # worksheet
    python3 scripts/name_designs.py templates/acme/catalog --import   # read it back

The catalog's judged half -- what each design is *for* -- is normally written
by the vision-agent looking at the renders. **A published skill cannot require
a GPU**, so this is the other way: measurements fill in what measurements can,
and the vision-agent (or, failing that, a person) fills the rest from the
renders in about five minutes for twenty designs. A person is better at it than
a 31B model was, and much better than a text-only model, which will answer
confidently about images it never received.

Without `--import` it writes `designs.md` beside the renders: one block per
design, the measured facts already filled, blanks where judgement goes.
`--import` reads it back into the `entries/*.json` and `decision_table.json`
that `catalog_build.py` expects.

**Every measured field is filled and never asked about.** Slot counts, body
capacity, image placeholders, grouped text, the title budget -- all of it is
arithmetic, all of it is already in `structure.json`, and asking a person to
retype it is how a worksheet gets abandoned half done.
"""

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine  # noqa: E402

engine()

from office.vision import CONTENT_TYPES, last_json  # noqa: E402

#: What a vision-agent answer may fill, beyond the three blanks a person fills.
#: ``content_shape``, ``max_items`` and ``looks_best`` are the fields a
#: text-only builder can act on: "3 numbers, each with a short label" matches
#: content, "a KPI dashboard" does not.
JUDGED = ("name", "content_types", "purpose", "use_when", "avoid_when",
          "content_shape", "max_items", "looks_best")


def read_answer(folder: Path) -> dict:
    """The vision-agent's JSON answer for one design, or ``{}``.

    ``vision_brief.py designs`` writes ``brief.md`` beside each render; the
    build model saves the vision-agent's reply beside it as ``answer.md``.
    """
    path = folder / "answer.md"
    if not path.is_file():
        return {}
    answer = last_json(path.read_text(encoding="utf-8")) or {}
    out = {}
    for key in JUDGED:
        value = answer.get(key)
        if isinstance(value, list):
            value = ", ".join(str(v) for v in value)
        if value not in (None, ""):
            out[key] = str(value).replace("\n", " ").strip()
    imagery = answer.get("imagery")
    if isinstance(imagery, list):
        out["imagery"] = [i for i in imagery if isinstance(i, dict)]
    return out

#: What a design is for, as far as a measurement can tell.  Deliberately few:
#: these are the ones that follow from shape counts alone, and everything finer
#: -- "an effort estimate", "a customer reference" -- is what the person adds.
def _guess(structure):
    """``(name, content_types)`` from the shape census alone."""
    slots = structure.get("slots") or []
    texts = [slot for slot in slots if slot.get("sample_text")]
    images = [slot for slot in slots
              if "Graphic" in str(slot.get("type", ""))
              or slot.get("placeholder") == "graphic"]
    tables = [slot for slot in slots if "Table" in str(slot.get("type", ""))]
    columns = len((structure.get("reading_order") or {}).get("columns") or [])
    biggest = max((max((font.get("size") or 0)
                       for font in slot.get("fonts") or [{}])
                   for slot in slots), default=0)

    if tables:
        return "Table slide", ["table"]
    if len(texts) <= 2 and biggest >= 28:
        return "Section divider", ["section-divider"]
    if columns >= 3:
        return f"{columns}-column layout", ["comparison", "cards"]
    if columns == 2:
        return "Two-column layout", ["comparison"]
    if images:
        return f"Content with {len(images)} image slot(s)", ["image"]
    if len(texts) >= 6:
        return "Dense content slide", ["bulleted-content", "matrix"]
    return "Content slide", ["bulleted-content"]


def _capacity(structure):
    slots = structure.get("slots") or []
    bodies = [slot for slot in slots
              if "Outliner" in str(slot.get("type", ""))
              or slot.get("placeholder") == "content"]
    empty = structure.get("empty_slots") or []
    images = [slot for slot in slots
              if "Graphic" in str(slot.get("type", ""))
              or slot.get("placeholder") == "graphic"]
    return {
        "headline": any("Title" in str(slot.get("type", "")) for slot in slots),
        "body_slots": len(bodies) + len([e for e in empty
                                         if e.get("placeholder") == "content"]),
        "max_bullets_per_body": None,          # a person reads this off a render
        "images": len(images),
        "supports_chart": False,
    }


BLOCK = re.compile(
    r"^##\s*Design\s+(\d+).*?$(.*?)(?=^##\s|\Z)", re.M | re.S)
CANVAS_BLOCK = re.compile(r"^##\s*Canvas\s*$(.*?)(?=^##\s|\Z)", re.M | re.S)
FIELD = re.compile(r"^-\s*\*\*(?P<key>[a-z_ ]+)\*\*\s*:\s*(?P<value>.*)$", re.M)


def worksheet(catalog: Path):
    """Write ``designs.md``: measured facts filled, judgement blank."""
    manifest = json.loads((catalog / "manifest.json").read_text())
    lines = [
        f"# Designs — {manifest['file']}", "",
        "One block per design. **The measured lines are already right — do not "
        "edit them.** Blocks marked *proposed by the vision-agent* were filled "
        "from its answer: review them with the user and correct what is "
        "wrong. Fill any blank field from the render beside the block; "
        "`content_shape` and `use_when` are what a builder actually matches "
        "content against.", "",
        "`name` is what the decision table will call it. `content_types` is a "
        "comma-separated list from: " + ", ".join(CONTENT_TYPES) + ". "
        "Anything else is fine; the table just repeats it.", "",
        "Then: `name_designs.py " + str(catalog) + " --import`", "",
        "## Canvas", "",
        "The design a diagram or a chart is drawn on, stripped back to its "
        "chrome. Leave both blank and `catalog_build.py` picks the plainest "
        "roomy design on the content master; fill them in when the probe "
        "deck shows that pick drawing over the design's own furniture, or "
        "looking like a cover rather than a content slide. `keep` is the "
        "shape indices to leave standing -- the title, a topic label -- as "
        "`probe.py` numbers them.", "",
        "- **design**: ",
        "- **keep**: ", "",
    ]
    for row in manifest["slides"]:
        number = row["number"]
        structure = json.loads(
            (catalog / f"slide-{number:02d}" / "structure.json").read_text())
        name, kinds = _guess(structure)
        capacity = _capacity(structure)
        answer = read_answer(catalog / f"slide-{number:02d}")
        render = catalog / f"slide-{number:02d}" / "render.png"
        keys = structure.get("retext_keys") or []
        budgets = structure.get("title_budgets") or []
        lines += [
            f"## Design {number}", "",
            f"![render]({render.relative_to(catalog)})", "",
            *(["*Proposed by the vision-agent — review.*", ""] if answer
              else []),
            f"- **name**: {answer.get('name') or name}",
            f"- **content_types**: "
            f"{answer.get('content_types') or ', '.join(kinds)}",
            *(f"- **{key}**: {answer.get(key, '')}" for key in JUDGED[2:]),
            "",
            f"*Measured — master `{row['master']}`, {structure.get('shape_count', 0)} "
            f"shapes, {capacity['body_slots']} body slot(s), "
            f"{capacity['images']} image slot(s), {len(keys)} retext key(s)"
            + (f", title budget {budgets[0]['chars_per_line']} chars on one "
               f"line at {budgets[0]['pt']:g}pt"
               if budgets and isinstance(budgets[0], dict)
               and budgets[0].get("chars_per_line") else "")
            + (", text inside groups" if row.get("grouped_text") else "")
            + ".*", "",
        ]
    out = catalog / "designs.md"
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


def read_worksheet(catalog: Path):
    """``{number: {field: value}}`` from a filled ``designs.md``."""
    path = catalog / "designs.md"
    if not path.is_file():
        raise SystemExit(f"no {path} — run without --import first, fill it in, "
                         f"then come back")
    text = path.read_text(encoding="utf-8")
    out = {}
    for match in BLOCK.finditer(text):
        number = int(match.group(1))
        fields = {field.group("key").strip(): field.group("value").strip()
                  for field in FIELD.finditer(match.group(2))}
        out[number] = fields
    return out


def read_canvas(catalog: Path):
    """``{"design": N, "keep": [...]}`` from the worksheet, or ``None``."""
    text = (catalog / "designs.md").read_text(encoding="utf-8")
    match = CANVAS_BLOCK.search(text)
    if not match:
        return None
    fields = {field.group("key").strip(): field.group("value").strip()
              for field in FIELD.finditer(match.group(1))}
    digits = re.findall(r"\d+", fields.get("design") or "")
    if not digits:
        return None
    block = {"design": int(digits[0])}
    keep = [int(k) for k in re.findall(r"\d+", fields.get("keep") or "")]
    if keep:
        block["keep"] = keep
    return block


def write_entries(catalog: Path, filled):
    """The judged half, from the worksheet plus the measurements."""
    manifest = json.loads((catalog / "manifest.json").read_text())
    entries_dir = catalog / "entries"
    entries_dir.mkdir(exist_ok=True)
    written, blank = 0, []
    rows = []
    for row in manifest["slides"]:
        number = row["number"]
        structure = json.loads(
            (catalog / f"slide-{number:02d}" / "structure.json").read_text())
        guess_name, guess_kinds = _guess(structure)
        fields = filled.get(number, {})
        answer = read_answer(catalog / f"slide-{number:02d}")
        kinds = [kind.strip() for kind
                 in (fields.get("content_types") or "").split(",") if kind.strip()]
        entry = {
            "slide": number,
            "layout": structure.get("layout"),
            "master": structure.get("master"),
            "written_by": "name_designs.py + the vision-agent",
            "name": fields.get("name") or guess_name,
            "purpose": fields.get("purpose") or "",
            "use_when": fields.get("use_when") or "",
            "avoid_when": fields.get("avoid_when") or "",
            "content_shape": fields.get("content_shape") or "",
            "looks_best": fields.get("looks_best") or "",
            "notes": "",
            "content_types": kinds or guess_kinds,
            "capacity": _capacity(structure),
            "imagery": answer.get("imagery") or [],
        }
        items = re.findall(r"\d+", fields.get("max_items") or "")
        if items:
            entry["capacity"]["max_items"] = int(items[0])
        if not entry["purpose"] or not entry["use_when"]:
            blank.append(number)
        (entries_dir / f"slide-{number:02d}.json").write_text(
            json.dumps(entry, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")
        written += 1
        rows.append((number, entry))
    return written, blank, rows


#: Decision-table rows, in the order a build script reads them.  Each is a
#: content type and the question it answers; the first design claiming that
#: type wins, and the canvas row is appended by `catalog_build.py` itself.
TABLE_ORDER = [
    ("deck opener", ("cover", "section-divider")),
    ("section start", ("section-divider",)),
    ("agenda", ("agenda", "bulleted-content")),
    ("a comparison, two things side by side", ("comparison",)),
    ("parallel items of equal weight", ("cards",)),
    ("a sequence of steps", ("process",)),
    ("a plan over time, with dates or phases", ("timeline", "process")),
    ("a repeating cycle", ("cycle",)),
    ("numbers that carry the point", ("kpis",)),
    ("a chart of real data", ("chart",)),
    ("a table of real rows and columns", ("table",)),
    ("a 2x2 or a grid of positions", ("matrix",)),
    ("who is involved, and who reports to whom", ("org-chart", "team")),
    ("a picture doing the arguing", ("image",)),
    ("a quotation or a customer voice", ("quote",)),
    ("one idea, stated large", ("one-big-idea",)),
    ("body text and bullets", ("bulleted-content",)),
    ("the last slide", ("closing",)),
]


def write_decision_table(catalog: Path, rows, canvas=None):
    """A table derived from what each design says it carries."""
    by_kind = {}
    for number, entry in rows:
        for kind in entry["content_types"]:
            by_kind.setdefault(kind, []).append((number, entry["name"]))

    table, used = [], set()
    for when, kinds in TABLE_ORDER:
        kind = next((k for k in kinds if by_kind.get(k)), kinds[0])
        options = by_kind.get(kind) or []
        pick = next((opt for opt in options if opt[0] not in used), None)
        if pick is None:
            pick = options[0] if options else None
        if pick is None:
            continue
        used.add(pick[0])
        table.append({
            "when": when,
            "use_slide": pick[0],
            "because": f"{pick[1]} — carries {kind}."
                       + (f" Alternatives: "
                          + ", ".join(str(n) for n, _ in options if n != pick[0])
                          if len(options) > 1 else ""),
        })

    spread = sorted({number for number, _ in rows})
    variety = (
        "Vary the design per slide. This template offers "
        f"{len(spread)}; a deck that uses one of them is a deck drawn from "
        "scratch on a template it never used.")
    payload = {"decision_table": table, "variety_rule": variety,
               "written_by": "name_designs.py — measured, not judged"}
    if canvas:
        payload["canvas"] = canvas
    (catalog / "decision_table.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8")
    return table


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("catalog", type=Path,
                    help="The directory catalog_extract.py wrote")
    ap.add_argument("--import", dest="do_import", action="store_true",
                    help="Read a filled designs.md into entries/")
    args = ap.parse_args()

    if not (args.catalog / "manifest.json").is_file():
        ap.error(f"no manifest.json in {args.catalog} — point this at the "
                 f"directory catalog_extract.py wrote")

    if not args.do_import:
        out = worksheet(args.catalog)
        print(f"Wrote {out}")
        answered = sum(1 for folder in args.catalog.glob("slide-*")
                       if (folder / "answer.md").is_file())
        if answered:
            print(f"\n{answered} design(s) prefilled from the vision-agent's "
                  f"answers. Review them with the user; the measured lines "
                  f"are already right.\n")
        else:
            print("\nNo answers yet. Write the briefs first -- "
                  "`vision_brief.py designs <catalog>` in the presentation "
                  "skill -- and give each to the vision-agent, one design per "
                  "call. Or fill the blanks by hand with the user.\n")
        print(f"Then: python3 {sys.argv[0]} {args.catalog} --import")
        return 0

    filled = read_worksheet(args.catalog)
    written, blank, rows = write_entries(args.catalog, filled)
    canvas = read_canvas(args.catalog)
    table = write_decision_table(args.catalog, rows, canvas)
    print(f"Wrote {written} entries and a {len(table)}-row decision table"
          + (f", canvas set by hand to design {canvas['design']}." if canvas
             else "."))
    if blank:
        print(f"\n! {len(blank)} design(s) have no purpose or use_when: "
              + ", ".join(str(n) for n in blank)
              + "\n  They still build, and the decision table will never "
                "choose them for the right reason. Fill them in designs.md "
                "and re-import.")
    print("\nNext: assemble the catalog.\n"
          f"  catalog_build.py {args.catalog} -o {args.catalog.parent}/catalog.md")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
