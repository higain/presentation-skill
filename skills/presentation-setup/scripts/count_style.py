#!/usr/bin/env python3
"""Count a template's colours and type sizes. Never choose them.

    python3 scripts/count_style.py brand.pptx           # the table to show a person
    python3 scripts/count_style.py brand.pptx --json    # the same thing as data

Every colour it prints is one the template already fills a shape or sets text
in, and every size is one it types at. That is the whole point: four separate
runs against a corporate template invented their own hex -- 0xE8730C for an
orange, 0x6B3FA0 for a purple -- and neither appears anywhere in the deck they
were supposedly matching. A colour that is counted cannot be invented.

**It stops short of naming roles.** Which of six brand colours is *the accent*
is a decision about what the company means, not a fact about the file: the
most-used colour is usually a grey, and the accent is often used twice. The
counts, the swatches and the contrast ratios go to a person, who answers in
five seconds what no heuristic gets right.

What it does decide is what to leave out, because those are facts:

* **LibreOffice's default shape fill** (0x729FCF) is what an unstyled shape
  gets. One template carries it on 308 shapes -- scaffolding the designer never
  filled -- and it is not a brand colour on any template anywhere.
* **Highlighter yellows** (0xFFFF00, 0xF5EB05) mark fill-in fields for whoever
  builds the deck. They are instructions, not design.
* **Colours on one shape in the whole file** are a stray, not a palette entry.
  The threshold is a floor, not a judgement: they are still listed, under the
  fold, in case one of them is the logo.
"""

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine  # noqa: E402

engine()

from office.impress import load_document, slides  # noqa: E402
from office.draw import contrast_ratio  # noqa: E402
from office.template import read_text, walk_shapes  # noqa: E402

#: Colours that are never a brand's, whatever their count.  See the docstring.
EXCLUDED = {
    0x729FCF: "LibreOffice's default shape fill -- an unstyled shape, not a "
              "brand colour",
    0xFFFF00: "highlighter yellow -- marks a fill-in field for whoever builds "
              "the deck",
    0xF5EB05: "highlighter yellow -- marks a fill-in field for whoever builds "
              "the deck",
}

#: Below this many uses a colour is a stray rather than a palette entry.
MIN_USES = 2

#: Text sizes below this are furniture -- think-cell stubs, corner marks,
#: registration text -- not a role in anybody's type scale.
MIN_SIZE_PT = 5.0


def _hex(value):
    return f"#{int(value) & 0xFFFFFF:06X}"


def survey(path):
    """``{"fills": Counter, "text": Counter, "sizes": Counter, ...}``.

    Fills are counted per shape, text colours and sizes per **character**: a
    12pt body paragraph and a 12pt page number are not equally strong evidence
    that 12pt is a role, and counting runs rather than characters says they
    are.
    """
    doc = load_document(path, read_only=True)
    fills, text_colours, sizes, fonts = Counter(), Counter(), Counter(), Counter()
    try:
        pages = slides(doc)
        for page in pages:
            for _address, shape in walk_shapes(page):
                style = str(getattr(shape, "FillStyle", ""))
                if "SOLID" in style:
                    colour = getattr(shape, "FillColor", None)
                    if colour is not None:
                        fills[int(colour) & 0xFFFFFF] += 1
                try:
                    paragraphs = read_text(shape)
                except Exception:
                    continue
                for body, cprops, _pprops in paragraphs:
                    weight = len(body.strip())
                    if not weight:
                        continue
                    colour = cprops.get("CharColor")
                    if colour is not None and int(colour) >= 0:
                        text_colours[int(colour) & 0xFFFFFF] += weight
                    size = cprops.get("CharHeight")
                    if size and size >= MIN_SIZE_PT:
                        sizes[round(float(size), 1)] += weight
                    name = cprops.get("CharFontName")
                    if name:
                        fonts[name] += weight
    finally:
        doc.close(False)
    return {"slides": len(pages), "fills": fills, "text": text_colours,
            "sizes": sizes, "fonts": fonts}


def candidates(found):
    """``[{colour, hex, fills, text, ratio_white, excluded}]``, strongest first.

    One row per colour, with how it is *used*: a colour that only ever fills
    and never sets type is a colour that can never be a word, and that single
    fact decides more about how a deck looks than any name given to it later.
    """
    fills, text = found["fills"], found["text"]
    rows = []
    for colour in set(fills) | set(text):
        rows.append({
            "colour": colour,
            "hex": _hex(colour),
            "fills": fills.get(colour, 0),
            "text": text.get(colour, 0),
            "ratio_white": round(contrast_ratio(colour, 0xFFFFFF), 2),
            "excluded": EXCLUDED.get(colour),
        })
    rows.sort(key=lambda row: (row["fills"] + row["text"]), reverse=True)
    return rows


def scale(found, roles=7):
    """The type sizes worth calling roles, largest first.

    A corporate master types at more sizes than it has roles -- rounding, a
    stray 11.5, one slide somebody nudged -- so this keeps the ones carrying
    real text and drops the tail. `office.draw` names exactly seven roles,
    which is the one number it is safe to assume.

    **Weight alone gets this wrong at the top.** A display size exists to be
    used once a deck: a 20-character cover title at 60pt weighs nothing against
    7,000 characters of 14pt body, so ranking by characters drops the largest
    size in the template every time -- which is how a first draft of this
    produced a scale with no `display` in it at all. The two largest sizes are
    kept whatever their weight, and the rest of the slots go by weight.
    """
    sizes = found["sizes"]
    if not sizes:
        return []
    biggest = sorted(sizes, reverse=True)[:2]
    strongest = [size for size, _ in sizes.most_common()
                 if size not in biggest][:max(roles - len(biggest), 0)]
    return sorted(set(biggest) | set(strongest), reverse=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("template")
    ap.add_argument("--json", action="store_true", help="Emit data, not a table")
    args = ap.parse_args()

    found = survey(args.template)
    rows = candidates(found)
    sizes = scale(found)

    if args.json:
        print(json.dumps({
            "slides": found["slides"],
            "colours": rows,
            "sizes": sizes,
            "all_sizes": {str(k): v for k, v in sorted(found["sizes"].items(),
                                                       reverse=True)},
            "fonts": found["fonts"].most_common(5),
        }, indent=2))
        return 0

    print(f"# {args.template} — {found['slides']} slides\n")
    print("## Colours the template actually uses\n")
    print("Ranked by use. `fills` counts shapes, `text` counts characters.")
    print("A colour with no text count can never be a word on a slide.\n")
    print("| colour | fills | text | on white | note |")
    print("|---|---:|---:|---:|---|")
    listed = 0
    for row in rows:
        if row["excluded"] or row["fills"] + row["text"] < MIN_USES:
            continue
        listed += 1
        readable = "type ok" if row["ratio_white"] >= 4.5 else (
            "headline only" if row["ratio_white"] >= 3.0 else "fill only")
        print(f"| `{row['hex']}` | {row['fills']} | {row['text']} | "
              f"{row['ratio_white']}:1 | {readable} |")
    if not listed:
        print("| — | | | | this template fills nothing and sets no colour |")

    strays = [row for row in rows
              if not row["excluded"] and row["fills"] + row["text"] < MIN_USES]
    if strays:
        print(f"\n**{len(strays)} colour(s) used once**, listed in case one is "
              f"the logo: "
              + ", ".join(f"`{row['hex']}`" for row in strays[:12]))

    dropped = [row for row in rows if row["excluded"]]
    if dropped:
        print("\n**Left out on purpose:**\n")
        for row in dropped:
            print(f"- `{row['hex']}` ({row['fills'] + row['text']} uses) — "
                  f"{row['excluded']}")

    print("\n## Type sizes\n")
    print("Ranked by how much text is set at them, then ordered large to "
          "small. `office.draw` has seven roles: display, title, subtitle, "
          "body, label, caption, footnote.\n")
    total = sum(found["sizes"].values()) or 1
    for size, count in sorted(found["sizes"].items(), reverse=True):
        share = 100 * count / total
        mark = "**" if size in sizes else ""
        note = "" if size in sizes else "   ·  not shortlisted"
        print(f"- {mark}{size:g}pt{mark} — {count} characters "
              f"({share:.0f}% of the deck's text){note}")
    print("\n(Bold are the seven the scale needs. The largest two are kept "
          "whatever their weight: a display size is used once a deck and would "
          "otherwise never make the list.)")
    if found["fonts"]:
        print("\n**Fonts:** " + ", ".join(
            f"{name} ({count})" for name, count in found["fonts"].most_common(3)))

    print("\n---\n\n**Nothing here is a decision yet.** Which colour is the "
          "accent, which is the alert, and which size is `body` are questions "
          "for the person whose brand it is. Ask them, then write the answers "
          "with `make_pack.py`.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
