#!/usr/bin/env python3
"""
Recover the judged half of a catalog from a catalog that was written by hand.

    python3 scripts/catalog_import.py references/company-template.md catalog/

**A migration aid, run once per template.**  Stage B's output --
``catalog/entries/slide-NN.json`` and ``catalog/decision_table.json`` -- is what
lets ``catalog_build.py`` rebuild the catalog.  A template catalogued before
that split existed, or one whose ``catalog/`` was treated as scratch and
deleted, has all of that judgement in one place only: the shipped markdown.
Which means the catalog cannot be regenerated, so the measured half can never
be refreshed, so it silently falls behind the template it describes.  (The
house catalog had drifted twelve retext keys behind its own asset, and was
missing every empty-slot and grouped-text line the extractor produces.)

This reads the markdown back into the JSON stage B would have written, so the
pipeline closes:

    catalog_extract.py  ->  structure   (measured, fresh)
    catalog_import.py   ->  entries     (judged, recovered from the old file)
    catalog_build.py    ->  the catalog (both halves, plus --overrides prose)

It reads only the judged fields -- name, purpose, content types, capacity, use
when, avoid when, imagery, notes, and the decision table.  Retext keys, empty
slots, group counts and title budgets are deliberately ignored: those come from
the template itself, and the hand-written file is exactly where they went stale.

Anything the markdown does not carry comes out absent rather than guessed, and
the run prints what it could not parse.  Check the result with:

    scripts/catalog_build.py catalog/ -o references/company-template.md --check
"""

import argparse
import json
import re
import sys
from pathlib import Path

CAPACITY = re.compile(
    r"headline\s*(?P<headline>[✓✗])\s*·\s*body slots\s*(?P<body>\d+)\s*·\s*"
    r"max bullets\s*(?P<bullets>\d+)\s*·\s*images\s*(?P<images>\d+)\s*·\s*"
    r"chart\s*(?P<chart>[✓✗])")

#: ``- **Field:** value``, with any wrapped continuation lines.
FIELD = re.compile(r"^- \*\*(?P<name>[^:*]+?):?\*\*[ :]*(?P<value>.*?)"
                   r"(?=^- \*\*|\Z)", re.M | re.S)

NOT_CATALOGUED = "_(not catalogued)_"


def _clean(value):
    value = " ".join(value.split())
    return "" if value == NOT_CATALOGUED else value


def _capacity(value):
    m = CAPACITY.search(value)
    if not m:
        return None
    return {"headline": m.group("headline") == "✓",
            "body_slots": int(m.group("body")),
            "max_bullets_per_body": int(m.group("bullets")),
            "images": int(m.group("images")),
            "supports_chart": m.group("chart") == "✓"}


def _imagery(value):
    """``[{depicts, role}]`` from ``a bird (icon); the logo (logo)``."""
    value = _clean(value)
    # "none", or the generator's own placeholder for pictures nobody described
    if not value or value == "none" or value.startswith("_"):
        return []
    # Drop the generator's disagreement note before splitting.
    value = re.sub(r"\s*_\[\d+ picture shape\(s\) in the file\]_\s*$", "", value)
    out = []
    for part in re.split(r"\s*;\s*", value):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(?P<depicts>.+?)\s*\((?P<role>[a-z-]+)\)$", part)
        if m:
            out.append({"depicts": m.group("depicts").strip(),
                        "role": m.group("role")})
        else:
            out.append({"depicts": part, "role": "unknown"})
    return out


def parse_entries(text):
    """``{slide: entry}`` for every ``### Slide N — Name`` block."""
    entries = {}
    # Stop at the next entry *or the next section*.  Without the ``^## ``
    # boundary the last entry's final field swallowed every appendix section
    # in the file -- 150 lines of prose collapsed into one design's Notes.
    for m in re.finditer(
            r"^### Slide (\d+) — (.+?)$(.*?)(?=^### Slide |^## |\Z)",
            text, re.M | re.S):
        n, name, body = int(m.group(1)), m.group(2).strip(), m.group(3)
        fields = {f.group("name").strip().lower(): f.group("value")
                  for f in FIELD.finditer(body)}
        entry = {"name": name}
        for key, field in (("purpose", "purpose"), ("use_when", "use when"),
                           ("avoid_when", "avoid when"), ("notes", "notes")):
            entry[key] = _clean(fields.get(field, ""))
        entry["content_types"] = re.findall(r"`([^`]+)`",
                                            fields.get("content types", ""))
        cap = _capacity(fields.get("capacity", ""))
        if cap:
            entry["capacity"] = cap
        entry["imagery"] = _imagery(fields.get("imagery", ""))
        entries[n] = entry
    return entries


def parse_decision_table(text):
    """``([rule], variety_rule)`` from the decision-table section."""
    section = re.search(r"^## Decision table.*?$(.*?)(?=^## )", text,
                        re.M | re.S)
    rules = []
    if section:
        for line in section.group(1).splitlines():
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cells) < 4 or cells[0].startswith("---"):
                continue
            when, slide, _name, because = cells[0], cells[1], cells[2], cells[3]
            if when.lower().startswith("when the content"):
                continue
            m = re.fullmatch(r"`?(\d+)`?", slide)
            if not m:
                continue                      # the canvas row, or a ⚠ row
            # The canvas row is generated from the measurements, not judged.
            if "**Canvas**" in _name or "canvas" == _name.strip().lower():
                continue
            rules.append({"when": re.sub(r"\*\*", "", when).strip(),
                          "use_slide": int(m.group(1)),
                          "because": because})
    variety = re.search(r"\*\*Variety rule:\*\*\s*(.+)", text)
    return rules, (variety.group(1).strip() if variety else None)


def main():
    ap = argparse.ArgumentParser(
        description="Recover catalog/entries + decision_table.json from a "
                    "hand-written catalog markdown file")
    ap.add_argument("markdown", help="The existing catalog, e.g. "
                                     "references/company-template.md")
    ap.add_argument("catalog", help="The directory catalog_extract.py wrote")
    ap.add_argument("--force", action="store_true",
                    help="Overwrite entries that already exist")
    args = ap.parse_args()

    catalog = Path(args.catalog)
    manifest_path = catalog / "manifest.json"
    if not manifest_path.exists():
        ap.error(f"{manifest_path} not found — run catalog_extract.py first, "
                 f"so the recovered entries can be stamped with each design's "
                 f"real layout and master")
    manifest = json.loads(manifest_path.read_text())
    rows = {row["number"]: row for row in manifest["slides"]}

    text = Path(args.markdown).read_text(encoding="utf-8")
    entries = parse_entries(text)
    rules, variety = parse_decision_table(text)

    out_dir = catalog / "entries"
    out_dir.mkdir(exist_ok=True)
    written, skipped = 0, []
    for n, entry in sorted(entries.items()):
        if n not in rows:
            skipped.append(n)
            continue
        dest = out_dir / f"slide-{n:02d}.json"
        if dest.exists() and not args.force:
            print(f"slide {n:02d}: exists, left alone (--force to replace)")
            continue
        # Identity is stamped from the manifest, exactly as catalog_gemma does.
        stamped = {"slide": n, "layout": rows[n]["layout"],
                   "master": rows[n]["master"], "recovered_from":
                   str(Path(args.markdown).name), **entry}
        dest.write_text(json.dumps(stamped, indent=2, ensure_ascii=False))
        written += 1

    table_path = catalog / "decision_table.json"
    if rules and (not table_path.exists() or args.force):
        table_path.write_text(json.dumps(
            {"decision_table": rules, "variety_rule": variety},
            indent=2, ensure_ascii=False))
        print(f"decision_table.json: {len(rules)} rules recovered")
    elif table_path.exists():
        print("decision_table.json: exists, left alone (--force to replace)")

    print(f"Recovered {written} entr{'y' if written == 1 else 'ies'} into "
          f"{out_dir}/")
    missing = sorted(set(rows) - set(entries))
    if missing:
        print(f"note: no entry in the markdown for design(s) {missing} — they "
              f"will read _(not catalogued)_ until the model pass runs for them")
    if skipped:
        print(f"WARNING: the markdown describes design(s) {skipped} that this "
              f"template does not have", file=sys.stderr)
    for n, entry in sorted(entries.items()):
        if n in rows and "capacity" not in entry:
            print(f"WARNING: slide {n:02d}: no capacity line could be parsed",
                  file=sys.stderr)
    print("\nNext: rebuild and compare —\n"
          f"  scripts/catalog_build.py {catalog} -o {args.markdown} --check")


if __name__ == "__main__":
    main()
