#!/usr/bin/env python3
"""Write a template pack from the answers a person gave.

    python3 scripts/make_pack.py --name acme --template brand.pptx \\
        --colour accent=#RRGGBB --colour alert=#RRGGBB --colour ok=#RRGGBB \\
        --colour muted=#RRGGBB --colour neutral=#RRGGBB \\
        --series '#RRGGBB,#RRGGBB,#RRGGBB' \\
        --size display=60 --size title=16 --size body=14 --size caption=10 \\
        --note Example --note-prefix 'Instructions: 1.'

Every #RRGGBB above is a placeholder on purpose. The first person to use this
script pasted an example's real hex values into a pack for a different
template, and three of that pack's colours were another company's.

Creates `templates/<name>/` in the presentation skill, copies the template in,
and writes `profile.json`. Re-running **merges**: whatever the command line
does not mention -- the palette, the type scale, the notes -- is kept from the
profile already there, so the notes stage cannot wipe the palette stage and a
role can be revised on its own. The catalog is never touched -- roles get
revised after somebody looks at a real deck, and that must not cost it.

**Every value here came from a person answering a question about counted
evidence.** This script chooses nothing. It refuses a colour the template does
not contain, warns when a role is missing or the scale does not descend, and
otherwise writes what it was told. The counting is `count_style.py`, the
asking is the skill, and keeping the three apart is what stops a setup run from
quietly inventing a brand.
"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))
from _skill import engine, presentation_skill, templates_dir  # noqa: E402

engine()

from office.colors import parse_hex  # noqa: E402
from office.draw import contrast_ratio  # noqa: E402

from count_style import survey  # noqa: E402

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


def unseen(template, wanted):
    """Colours asked for that the template does not contain.

    The one check that has to be mechanical. Every other guard in this skill is
    advice a person can overrule; this one catches the exact failure the whole
    pack idea exists to prevent -- a deck built in colours the company does not
    own -- and it caught it the first time somebody used this script, because
    the example command in SKILL.md carried real hex values from another
    template and they were pasted verbatim.

    Compared against what `count_style.py` counts on the slides. A colour that
    lives only in the master or a theme is invisible to that and legitimate,
    which is what ``--allow-unseen`` is for; it is rare enough to be worth
    typing.
    """
    found = survey(template)
    present = set(found["fills"]) | set(found["text"])
    return [colour for colour in wanted if colour not in present], present


def _nearest(colour, present):
    """The counted colour closest to one that is not there."""
    def distance(other):
        return sum((((colour >> shift) & 255) - ((other >> shift) & 255)) ** 2
                   for shift in (16, 8, 0))
    return min(present, key=distance) if present else None


def build_profile(args, existing=None):
    """The profile dict, with every key the engine reads.

    ``existing`` is the profile already on disk, if any. Each of the three
    groups -- colours (roles and series together), type scale, annotations --
    is taken from the command line when any of it was given there, and from
    ``existing`` otherwise. The first version of this script rebuilt the
    profile from the command line alone, and the documented notes-only re-run
    of stage 3 silently emptied the palette stage 2 had just written.
    """
    existing = existing or {}
    old_palette = existing.get("palette") or {}
    old_roles = old_palette.get("_roles") or {}

    _parse = parse_hex

    if args.colour or args.series:
        roles = dict(args.colour or ())
        series = list(args.series or ())
        if not args.colour or not args.series:
            # Half given: keep the other half from the profile on disk.
            kept_roles = {role: _parse(old_palette.get(key))
                          for role, key in old_roles.items()
                          if role != "series"}
            kept_series = [_parse(old_palette.get(key))
                           for key in old_roles.get("series") or ()]
            if not args.colour:
                roles = {r: c for r, c in kept_roles.items() if c is not None}
            if not args.series:
                series = [c for c in kept_series if c is not None]
    else:
        roles = {role: _parse(old_palette.get(key))
                 for role, key in old_roles.items() if role != "series"}
        roles = {r: c for r, c in roles.items() if c is not None}
        series = [c for c in (_parse(old_palette.get(key))
                              for key in old_roles.get("series") or ())
                  if c is not None]

    if args.size:
        scale = dict(args.size)
    else:
        scale = {role: float(size) for role, size
                 in (existing.get("type_scale") or {}).items()}

    old_notes = existing.get("annotations") or {}
    if args.note or args.note_prefix:
        notes = list(args.note or ())
        prefixes = list(args.note_prefix or ())
    else:
        notes = list(old_notes.get("exact") or ())
        prefixes = list(old_notes.get("prefixes") or ())

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
            "exact": notes,
            "prefixes": prefixes,
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
        ratio = contrast_ratio(parse_hex(accent), 0xFFFFFF)
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
    ap.add_argument("--allow-unseen", action="store_true",
                    help="Accept a colour the template's slides do not show "
                         "(it lives only in the master)")
    ap.add_argument("--default", action="store_true",
                    help="Make this the pack used when no template is named")
    args = ap.parse_args()

    template = Path(args.template)
    if not template.is_file():
        ap.error(f"no template at {template}")
    if not re.fullmatch(r"[a-z0-9][a-z0-9_-]*", args.name):
        ap.error("--name: lowercase letters, digits, dash and underscore — it "
                 "is a directory name and a build script types it")

    asked = [colour for _role, colour in (args.colour or ())] + list(args.series or ())
    if asked and not args.allow_unseen:
        missing, present = unseen(template, asked)
        if missing:
            lines = []
            for colour in dict.fromkeys(missing):
                near = _nearest(colour, present)
                lines.append(f"    #{colour:06X}"
                             + (f"  — nearest colour the template does use: "
                                f"#{near:06X}" if near is not None else ""))
            ap.error(
                "these colours are not in %s:\n%s\n"
                "  Take them from count_style.py rather than from an example "
                "command — a pack in colours the company does not own is the "
                "one failure this skill exists to prevent.\n"
                "  If a colour really is the brand's and lives only in the "
                "master, pass --allow-unseen."
                % (template, "\n".join(lines)))

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
    kept = []
    if existing:
        if not (args.colour or args.series):
            kept.append("palette")
        if not args.size:
            kept.append("type scale")
        if not (args.note or args.note_prefix):
            kept.append("notes")
    print(f"  profile.json   {len(profile['palette']) - 2} colours, "
          f"{len(profile['type_scale'])} type roles, "
          f"{len(profile['annotations']['exact'])} notes"
          + (f"  (kept from the existing profile: {', '.join(kept)})"
             if kept else ""))
    for line in warnings(profile):
        print(f"  ! {line}")
    print("\nNext: extract the catalog, or the pack has a palette and no idea "
          "which design carries what.")
    print(f"  python3 {presentation_skill()}/scripts/catalog_extract.py "
          f"{asset} -o {root}/catalog")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
