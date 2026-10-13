"""Icons: the pack's own, then an installed open icon set.

A slide that wants a pictogram -- a shield beside "security", a cup on the
break slide -- had no way to get one.  One run drew 45 "icons" out of
rectangles and lines for a 58-slide deck: a page of build script, and on the
slide they read as "II", "#" and three dots.  The original deck had line icons
in circles; nothing in the skill could place one.

Where an icon comes from, first match wins:

1. **the pack** -- ``templates/<name>/icons/``, harvested from the template by
   ``presentation-setup`` (``harvest_icons.py``).  These are the brand's own
   pictograms and the only ones guaranteed to match its style.
2. **``$PRESENTATION_ICONS``** -- extra directories of ``.svg`` files,
   ``os.pathsep``-separated: an icon set the company licenses.
3. **Element Plus** (``ep``) -- 293 monochrome icons, MIT, shipped with the
   skill in ``icons/ep/`` (licence beside them).  Always there.
4. **Bootstrap Icons** -- ``apt install bootstrap-icons``, MIT, about 2000
   monochrome SVGs under ``/usr/share/bootstrap-icons/svg``, if installed.

``"ep:lock"`` or ``"bootstrap:lock"`` picks a set explicitly; a bare name
takes the first set that has it.

Every icon is drawn in **one palette colour**: the SVG's fills and strokes are
rewritten to it, so an icon is never in a colour the brand does not use.  An
icon with several colours (a pack icon the harvest marked ``"mono": false``)
is placed as it is.

``scripts/icons.py search <word>`` lists what exists; ``D.icon`` places one.
"""

from __future__ import annotations

import atexit
import difflib
import json
import os
import re
import shutil
import tempfile
from pathlib import Path

ICON_DIRNAME = "icons"
INDEX_NAME = "index.json"
EXTRA_ENV = "PRESENTATION_ICONS"

#: Sets shipped inside the skill, by label.
BUNDLED_SETS = (
    ("ep", Path(__file__).resolve().parent.parent.parent / "icons" / "ep"),
)

#: Installed icon sets, by label.  Debian paths; missing ones are skipped.
SYSTEM_SETS = (
    ("bootstrap", Path("/usr/share/bootstrap-icons/svg")),
)

INSTALL_HINT = ("Bootstrap Icons are not installed: "
                "`sudo apt install bootstrap-icons` adds about 2000 (MIT).")


def _pack_dir(pack):
    if pack is None:
        from .pack import active_pack
        pack = active_pack()
    return None if pack is None else Path(pack.root) / ICON_DIRNAME


def pack_index(pack=None):
    """``{name: entry}`` from the pack's ``icons/index.json``, or ``{}``."""
    folder = _pack_dir(pack)
    if folder is None:
        return {}
    try:
        data = json.loads((folder / INDEX_NAME).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data.get("icons", {})


def sources(pack=None):
    """``[(label, directory)]`` in lookup order; only directories that exist."""
    out = []
    folder = _pack_dir(pack)
    if folder is not None and folder.is_dir():
        out.append(("pack", folder))
    for i, extra in enumerate(p for p in os.environ.get(EXTRA_ENV, "")
                              .split(os.pathsep) if p):
        if Path(extra).is_dir():
            out.append((f"extra{i + 1}" if i else "extra", Path(extra)))
    for label, folder in BUNDLED_SETS + SYSTEM_SETS:
        if folder.is_dir():
            out.append((label, folder))
    return out


def catalog(pack=None, every=False):
    """``{name: (label, path, entry)}`` -- every icon reachable, first wins.

    ``entry`` is the pack index entry (keywords, ``mono``) or ``{}``.
    ``every=True`` adds each icon under ``"label:name"`` as well, so a name
    another set shadows can still be reached.
    """
    out = {}

    def put(name, value):
        out.setdefault(name, value)
        if every:
            out[f"{value[0]}:{name}"] = value

    index = pack_index(pack)
    for label, folder in sources(pack):
        if label == "pack":
            for name, entry in sorted(index.items()):
                path = folder / entry.get("file", f"{name}.svg")
                if path.is_file():
                    put(name, (label, path, entry))
            continue
        for path in sorted(folder.glob("*.svg")):
            put(path.stem, (label, path, {}))
    return out


def search(query, pack=None, limit=40):
    """Icons whose name or keywords contain every word of ``query``.

    Pack icons first: they are the brand's.  An icon another set shadows is
    listed too, with its set: ``("lock", "ep")`` is placed as ``"ep:lock"``.
    """
    words = [w for w in re.split(r"[\s,]+", query.lower()) if w]
    hits, seen = [], set()
    for key, (label, path, entry) in catalog(pack, every=True).items():
        name = key.split(":", 1)[-1]
        if (name, label) in seen:
            continue
        seen.add((name, label))
        hay = " ".join([name, *entry.get("keywords", ())]).lower()
        if all(w in hay for w in words):
            hits.append((label != "pack", len(name), name, label))
    return [(name, label) for _, _, name, label in sorted(hits)[:limit]]


def find(name, pack=None):
    """``(label, path, entry)`` for ``name``; raises with suggestions."""
    found = catalog(pack, every=":" in name)
    if name in found:
        return found[name]
    found = catalog(pack)
    close = difflib.get_close_matches(name, list(found), n=8, cutoff=0.6)
    partial = [n for n, _ in search(name.replace("-", " "), pack, limit=8)]
    suggest = list(dict.fromkeys(close + partial))[:10]
    labels = sorted({label for _, (label, _, _) in found.items()})
    raise ValueError(
        "icon(): no icon %r.\n  %s\n  Sets searched: %s.\n  %s"
        % (name,
           ("Close: " + ", ".join(suggest)) if suggest else "Nothing close.",
           ", ".join(labels) or "none",
           "Find one with `scripts/icons.py search <word>` -- English words, "
           "names like 'lock', 'coffee-cup', 'user'."
           + ("" if "bootstrap" in labels else "\n  " + INSTALL_HINT)))


# --------------------------------------------------------------------------
# colour
# --------------------------------------------------------------------------

_COLOUR = re.compile(
    r"#[0-9a-fA-F]{6}\b|#[0-9a-fA-F]{3}\b|rgb\([^)]*\)|currentColor")
_PAINT_ATTR = re.compile(r'\b(fill|stroke)\s*=\s*"([^"]*)"')
_PAINT_CSS = re.compile(r"\b(fill|stroke)\s*:\s*([^;\"'}]+)")
_NO_PAINT = {"none", "transparent"}


def colours(svg):
    """The distinct paints an SVG uses, ``currentColor`` counted as one."""
    found = set()
    for _, value in _PAINT_ATTR.findall(svg) + _PAINT_CSS.findall(svg):
        value = value.strip()
        if value.lower() not in _NO_PAINT and not value.startswith("url("):
            found.add(value.lower())
    return found


def recolour(svg, rgb):
    """``svg`` with every fill and stroke set to ``rgb`` (an int, 0xRRGGBB).

    Strips ``<style>`` blocks and ``class`` attributes -- Office icons carry
    their theme colour there -- and gives the root a ``fill`` so paths that
    set none (black by default) take the colour too.  ``none`` stays none.
    """
    hexc = "#%06X" % int(rgb)
    svg = re.sub(r"<style\b.*?</style>", "", svg, flags=re.S)
    svg = re.sub(r'\sclass="[^"]*"', "", svg)

    def attr(m):
        value = m.group(2).strip()
        if value.lower() in _NO_PAINT or value.startswith("url("):
            return m.group(0)
        return f'{m.group(1)}="{hexc}"'

    def css(m):
        value = m.group(2).strip()
        if value.lower() in _NO_PAINT or value.startswith("url("):
            return m.group(0)
        return f"{m.group(1)}:{hexc}"

    svg = _PAINT_ATTR.sub(attr, svg)
    svg = _PAINT_CSS.sub(css, svg)
    svg = svg.replace("currentColor", hexc)
    root = re.search(r"<svg\b[^>]*>", svg)
    if root and not re.search(r"\sfill\s*=", root.group(0)):
        tag = root.group(0)
        end = -2 if tag.endswith("/>") else -1
        svg = (svg[:root.start()] + tag[:end] + f' fill="{hexc}"' + tag[end:]
               + svg[root.end():])
    return svg


def aspect(svg):
    """Width over height from the ``viewBox`` (or width/height), else 1."""
    m = re.search(r'viewBox\s*=\s*"\s*([-\d.]+)[\s,]+([-\d.]+)[\s,]+'
                  r'([\d.]+)[\s,]+([\d.]+)', svg)
    try:
        if m:
            return float(m.group(3)) / float(m.group(4))
        w = re.search(r'<svg\b[^>]*\swidth="([\d.]+)', svg)
        h = re.search(r'<svg\b[^>]*\sheight="([\d.]+)', svg)
        return float(w.group(1)) / float(h.group(1))
    except (AttributeError, ValueError, ZeroDivisionError):
        return 1.0


_TMP = None


def write_temp(svg, name):
    """``svg`` written to a file LibreOffice can load; removed at exit."""
    global _TMP
    if _TMP is None:
        _TMP = Path(tempfile.mkdtemp(prefix="icons-"))
        atexit.register(shutil.rmtree, _TMP, True)
    safe = re.sub(r"[^\w-]", "_", name)
    path = Path(tempfile.mkstemp(prefix=safe + "-", suffix=".svg",
                                 dir=_TMP)[1])
    path.write_text(svg, encoding="utf-8")
    return path
