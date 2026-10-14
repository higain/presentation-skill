#!/usr/bin/env python3
"""Find an icon to place with ``D.icon``.

    scripts/icons.py search shield            # names matching every word
    scripts/icons.py list [--template acme]   # the pack's own icons
    scripts/icons.py sets                     # which icon sets are installed

Pack icons are listed first: they are the brand's own.  Search in English --
names like 'lock', 'coffee-cup', 'user'.  A shadowed name is reachable as
'ep:lock'.  Nothing fits?  Take the closest icon that exists; never install
an icon set.
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).absolute().parent))

from office import icons  # noqa: E402
from office.pack import pack_named, default_pack  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("command", choices=("search", "list", "sets"))
    ap.add_argument("words", nargs="*")
    ap.add_argument("--template", help="pack name (default: the default pack)")
    args = ap.parse_args()
    pack = pack_named(args.template) if args.template else default_pack()

    if args.command == "sets":
        for label, folder in icons.sources(pack):
            count = (len(icons.pack_index(pack)) if label == "pack"
                     else len(list(folder.glob("*.svg"))))
            print(f"{label:10} {count:5}  {folder}")
        return 0
    if args.command == "list":
        index = icons.pack_index(pack)
        if not index:
            print(f"{pack.name if pack else 'no pack'}: no icons of its own "
                  "(presentation-setup harvest_icons.py collects them).")
        for name, entry in sorted(index.items()):
            print(f"{name:28} {', '.join(entry.get('keywords', []))}")
        return 0
    if not args.words:
        ap.error("search needs a word")
    hits = icons.search(" ".join(args.words), pack)
    for name, label in hits:
        print(f"{name:32} {label}   (\"{label}:{name}\")" if label != "pack"
              else f"{name:32} {label}")
    if not hits:
        print("nothing found -- try a broader English word, or take the "
              "closest icon that exists")
    return 0


if __name__ == "__main__":
    sys.exit(main())
