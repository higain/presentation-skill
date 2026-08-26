#!/usr/bin/env python3
"""Write a template pack from the answers a person gave.

    python3 scripts/make_pack.py --name acme --template brand.pptx \\
        --colour accent=#009EE0 --colour alert=#E2001A --colour ok=#009036 \\
        --colour muted=#D9D9D9 --colour neutral=#4A4A4A \\
        --series '#009EE0,#F29400,#97BF0D' \\
        --size display=60 --size title=16 --size body=14 --size caption=10 \\
        --note Example --note-prefix 'Instructions: 1.'

Creates `templates/<name>/` in the presentation skill, copies the template in,
and writes `profile.json`. Re-running updates the profile and leaves the
catalog alone -- roles get revised after somebody looks at a real deck, and
that must not cost the catalog.

**Every value here came from a person answering a question about counted
evidence.** This script chooses nothing: it validates that the colours are in
the template, that the roles it needs are present, and that the scale
descends -- then writes what it was told. The counting is `count_style.py`, the
asking is the skill, and keeping the three apart is what stops a setup run from
quietly inventing a brand.
"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _skill import engine, presentation_skill, templates_dir  # noqa: E402

engine()

from office.draw import contrast_ratio  # noqa: E402

#: Roles the drawing primitives and the accent packs both rely on.  A pack
#: missing one of these still works -- `D.color` raises with the list -- but
#: `references/accents.md` cannot be followed, so the setup skill asks for all
#: five and this warns rather than fails if it did not get them.
WANTED_ROLES = ("accent", "alert", "ok", "muted", "neutral")

#: The seven `office.draw` knows.  A profile may define fewer; `size_for` says
#: which are missing when something asks for one.
ROLE_SIZES = ("display", "title", "subtitle", "body", "label", "caption",
              "footnote")

HEX = re.compile(r"^#?[0-9A-Fa-f]{6}$")


def _colour(value):
    if not HEX.match(value.strip()):
        raise argparse.ArgumentTypeError(
            f"{value!r} is not a #RRGGBB colour. Take it from count_style.py; "
            f"do not type one from memory.")
    return int(value.strip().lstrip("#"), 16)


def _pair(value, cast=str):
    if "=" not in value:
        raise argparse.ArgumentTypeError(f"{value!r} should be name=value")
    name, _, raw = value.partition("=")
    return name.strip(), cast(raw.strip())


def _name_for(colour):
    """A readable key for a colour, from its own hue and lightness.

    The keys are how a build script says `D.color("cyan")`, so they want to be
    recognisable rather than correct -- `c3` is a key nobody will ever type on
    purpose. Roles are what the documentation uses; these are for the person
    reading a palette table and checking it against their brand book.
    """
    red, green, blue = (colour >> 16) & 255, (colour >> 8) & 255, colour & 255
    high, low = max(red, green, blue), min(red, green, blue)
    # HSL lightness, not value: a saturated cyan has a *value* of 0.88 and is
    # not a light colour, and naming it "cyan_light" beside a genuinely pale
    # one is the kind of key nobody trusts twice.
    light = (high + low) / 510
    if high - low < 24:                       # no hue worth naming
        if light > 0.93:
            return "white"
        if light < 0.12:
            return "black"
        return "grey_light" if light > 0.72 else (
            "grey" if light > 0.45 else "anthracite")
    if high == red:
        hue = 60 * (((green - blue) / (high - low)) % 6)
    elif high == green:
        hue = 60 * (((blue - red) / (high - low)) + 2)
    else:
        hue = 60 * (((red - green) / (high - low)) + 4)
    for edge, name in ((15, "red"), (45, "orange"), (70, "yellow"),
                       (95, "lime"), (150, "green"), (185, "teal"),
                       (205, "cyan"), (250, "blue"), (290, "violet"),
                       (330, "magenta"), (360, "red")):
        if hue < edge:
            return f"{name}_light" if light > 0.75 else name
    return "colour"


def build_profile(args, existing=None):
    """The profile dict, with every key the engine reads."""
    roles = dict(args.colour or ())
    series = list(args.series or ())
    scale = dict(args.size or ())

    names, palette = {}, {}
    for colour in list(roles.values()) + series:
        if colour in names:
            continue
        key = _name_for(colour)
        suffix = 2
        while key in palette and palette[key] != colour:
            key, suffix = f"{_name_for(colour)}_{suffix}", suffix + 1
        names[colour] = key
        palette[key] = colour

    profile = {
        "_comment": [
            "Written by presentation-setup from counted evidence and a "
            "person's answers.",
            "",
            "Every colour is one this template already fills a shape or sets "
            "text in, and every size is one it types at. Nothing here was "
            "chosen by a script: count_style.py counted, somebody named.",
            "",
            "'exact' matches a shape's whole text, case-insensitively and "
            "with whitespace collapsed; 'prefixes' matches the start of it, "
            "for a multi-sentence instruction block.",
            "",
            "Re-run make_pack.py to revise the roles. The catalog is not "
            "touched by it.",
        ],
        "name": args.name,
        "asset": "template" + Path(args.template).suffix.lower(),
        "annotations": {
            "exact": list(args.note or ()),
            "prefixes": list(args.note_prefix or ()),
        },
        "palette": {
            "_order": [names[c] for c in
                       list(dict.fromkeys(series + list(roles.values())))],
            **{key: f"#{value:06X}" for key, value in palette.items()},
            "_roles": {
                **{role: names[colour] for role, colour in roles.items()},
                "series": [names[colour] for colour in series],
            },
        },
        "type_scale": {role: float(size) for role, size in scale.items()},
    }
    if existing and not scale:
        profile["type_scale"] = existing.get("type_scale", {})
    return profile


def warnings(profile):
    """Everything worth saying before the pack is used in anger."""
    out = []
    roles = profile["palette"]["_roles"]
    palette = profile["palette"]
    missing = [role for role in WANTED_ROLES if role not in roles]
    if missing:
        out.append(
            f"no {', '.join(missing)} role — `references/accents.md` needs all "
            f"of {', '.join(WANTED_ROLES)}; a pack without them still builds, "
            f"but the accent packs cannot be followed")
    if not roles.get("series"):
        out.append("no series — D.series(n) falls back to palette order, so a "
                   "chart's colours are whatever order the palette is in")

    accent = palette.get(roles.get("accent", ""))
    if accent:
        ratio = contrast_ratio(int(str(accent).lstrip("#"), 16), 0xFFFFFF)
        if ratio < 3.0:
            out.append(
                f"the accent is {ratio:.1f}:1 on white — **fill only**, it can "
                f"never be a word on a slide. That is normal for a brand "
                f"colour and it is what the Ember pack is for")
        elif ratio < 4.5:
            out.append(
                f"the accent is {ratio:.1f}:1 on white — headline sizes only, "
                f"never body text")

    scale = profile["type_scale"]
    missing_sizes = [role for role in ROLE_SIZES if role not in scale]
    if missing_sizes:
        out.append(f"no size for {', '.join(missing_sizes)} — `role=` raises "
                   f"for those and names the ones you do have")
    values = [scale[role] for role in ROLE_SIZES if role in scale]
    if values != sorted(values, reverse=True):
        out.append("the type scale does not descend in role order — check "
                   "display > title > subtitle > body > label > caption > "
                   "footnote")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--name", required=True, help="Pack name, e.g. acme")
    ap.add_argument("--template", required=True, help="The template file")
    ap.add_argument("--colour", action="append", metavar="ROLE=#RRGGBB",
                    type=lambda v: _pair(v, _colour),
                    help="A role: accent, alert, ok, muted, neutral")
    ap.add_argument("--series", type=lambda v: [_colour(part)
                                                for part in v.split(",")],
                    help="Figure colours in order, comma separated")
    ap.add_argument("--size", action="append", metavar="ROLE=PT",
                    type=lambda v: _pair(v, float),
                    help="A type role: " + ", ".join(ROLE_SIZES))
    ap.add_argument("--note", action="append", metavar="TEXT",
                    help="An authoring note matched on the whole string")
    ap.add_argument("--note-prefix", action="append", metavar="TEXT",
                    help="An authoring note matched on its opening words")
    ap.add_argument("--default", action="store_true",
                    help="Make this the pack used when no template is named")
    args = ap.parse_args()

    template = Path(args.template)
    if not template.is_file():
        ap.error(f"no template at {template}")
    if not re.fullmatch(r"[a-z0-9][a-z0-9_-]*", args.name):
        ap.error("--name: lowercase letters, digits, dash and underscore — it "
                 "is a directory name and a build script types it")

    root = templates_dir() / args.name
    root.mkdir(parents=True, exist_ok=True)
    asset = root / ("template" + template.suffix.lower())
    if asset.resolve() != template.resolve():
        shutil.copy2(template, asset)

    profile_path = root / "profile.json"
    existing = None
    if profile_path.is_file():
        try:
            existing = json.loads(profile_path.read_text(encoding="utf-8"))
        except ValueError:
            existing = None

    profile = build_profile(args, existing)
    profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False)
                            + "\n", encoding="utf-8")

    if args.default:
        (templates_dir() / "default").write_text(args.name + "\n",
                                                 encoding="utf-8")

    print(f"Wrote {root}/")
    print(f"  {asset.name}   the template")
    print(f"  profile.json   {len(profile['palette']) - 2} colours, "
          f"{len(profile['type_scale'])} type roles, "
          f"{len(profile['annotations']['exact'])} notes")
    for line in warnings(profile):
        print(f"  ! {line}")
    print("\nNext: extract the catalog, or the pack has a palette and no idea "
          "which design carries what.")
    print(f"  python3 {presentation_skill()}/scripts/catalog_extract.py "
          f"{asset} -o {root}/catalog")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
