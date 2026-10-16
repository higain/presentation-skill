#!/usr/bin/env python3
"""Add slides from finished decks to a pack as new designs.

    python3 scripts/add_designs.py candidates acme deck.pptx [more.pptx ...]
    python3 scripts/add_designs.py texts acme deck.pptx --slides 5,12,21
    python3 scripts/add_designs.py append acme add-designs/deck.texts.json

A company's best slides are rarely in its master. They are in project decks:
someone built a risk-to-countermeasure mapping for one workshop, on the
company master, and it never made it back. This takes such slides and makes
them designs of the pack, so the catalog and its decision table can offer them.

**candidates** compares every slide of the decks with the pack's designs by
geometry -- where the top-level shapes sit -- and groups the slides that are
new into candidates: six phase slides of one workshop are one candidate, not
six. It reads XML only; no LibreOffice, no model. A person picks from the list.

**texts** writes ``add-designs/<deck>.texts.json``: every paragraph of the
chosen slides with a blank ``new`` beside it. The slide's project wording is a
customer's; it is replaced with sample text of about the same length before
the slide enters the pack, because the pack's wording becomes the retext keys
and the budgets every later deck is measured against. Alt texts are listed
prefilled with their own text, to be checked.

**append** copies the slides into the pack's template at the end, so no
existing design number moves. It works on the ``.pptx`` parts directly: the
slide, its pictures, charts and embedded objects are copied, its layout is
rebound to the template's layout of the same name, speaker notes and comments
are dropped. LibreOffice never saves the template -- every ``.pptx`` save
through it appends one more master. The template before the change is kept in
``<pack>/history/``.

``.pptx`` and ``.potx`` packs only. Nothing leaves the machine.
"""

import argparse
import hashlib
import json
import posixpath
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).absolute().parent))
from _skill import engine  # noqa: E402

engine()

from office.pack import pack_named  # noqa: E402

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A, P, R = (f"{{{NS[k]}}}" for k in ("a", "p", "r"))
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
SLIDE_TYPE = DOC_REL + "slide"
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

#: Relationships of a slide that do not travel into the template.  Notes and
#: comments are the project's conversation, not the design; a link to another
#: slide points into a deck the template is not.
DROPPED = {"notesSlide", "comments", "slide", "commentAuthors"}

#: Top-level shapes of two slides are the same shape when their boxes overlap
#: this much (intersection over union).
SAME_BOX = 0.6
#: A slide matches a design when this share of the shapes on both pair up.
#: A copy of a design with its text replaced scores 1.0; one where somebody
#: moved a box or added a caption scores about 0.8.
SAME_SLIDE = 0.75

OUT_DIR = "add-designs"


# --------------------------------------------------------------- package I/O

class Package:
    """A ``.pptx`` as a dict of parts, with the few lookups this needs."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as archive:
            self.order = archive.namelist()
            self.parts = {name: archive.read(name) for name in self.order}
        self.types = etree.fromstring(self.parts["[Content_Types].xml"])

    def xml(self, name):
        return etree.fromstring(self.parts[name])

    def content_type(self, name):
        for node in self.types:
            if (node.tag == f"{{{CT_NS}}}Override"
                    and node.get("PartName") == "/" + name):
                return node.get("ContentType"), True
        ext = name.rsplit(".", 1)[-1].lower()
        for node in self.types:
            if (node.tag == f"{{{CT_NS}}}Default"
                    and node.get("Extension", "").lower() == ext):
                return node.get("ContentType"), False
        return None, False

    def rels(self, name):
        """``[(id, type, target part or URL, external)]`` for one part."""
        folder, base = posixpath.split(name)
        rels_name = posixpath.join(folder, "_rels", base + ".rels")
        if rels_name not in self.parts:
            return []
        out = []
        for rel in etree.fromstring(self.parts[rels_name]):
            external = rel.get("TargetMode") == "External"
            target = rel.get("Target")
            if not external:
                target = (target.lstrip("/") if target.startswith("/")
                          else posixpath.normpath(posixpath.join(folder, target)))
            out.append((rel.get("Id"), rel.get("Type"), target, external))
        return out

    def slide_parts(self):
        """Slide part names in presentation order."""
        pres = self.xml("ppt/presentation.xml")
        by_id = {rid: target for rid, kind, target, _ in self.rels("ppt/presentation.xml")
                 if kind == SLIDE_TYPE}
        lst = pres.find(f"{P}sldIdLst")
        return [] if lst is None else [by_id[s.get(f"{R}id")] for s in lst]

    def slide_size(self):
        size = self.xml("ppt/presentation.xml").find(f"{P}sldSz")
        return int(size.get("cx")), int(size.get("cy"))

    def layout_of(self, slide):
        for _, kind, target, _ in self.rels(slide):
            if kind.endswith("/slideLayout"):
                return target
        return None

    def layout_name(self, layout):
        node = self.xml(layout).find(f"{P}cSld")
        return node.get("name", "") if node is not None else ""

    def master_layout_names(self, layout):
        """The names of every layout on the master ``layout`` belongs to."""
        master = next((t for _, k, t, _ in self.rels(layout)
                       if k.endswith("/slideMaster")), None)
        if master is None:
            return set()
        return {self.layout_name(t) for _, k, t, _ in self.rels(master)
                if k.endswith("/slideLayout")}


def _rel_target(source_part, target_part):
    return posixpath.relpath(target_part, posixpath.dirname(source_part))


def _free_name(names, wanted):
    """``wanted`` if no part has it, else the next free number in its stem."""
    if wanted not in names:
        return wanted
    folder, base = posixpath.split(wanted)
    stem, dot, ext = base.rpartition(".")
    prefix = re.sub(r"\d+$", "", stem)
    pattern = re.compile(re.escape(posixpath.join(folder, prefix)) + r"(\d+)\."
                         + re.escape(ext) + "$")
    numbers = [int(m.group(1)) for m in map(pattern.match, names) if m]
    return posixpath.join(folder, f"{prefix}{max(numbers, default=0) + 1}.{ext}")


# ------------------------------------------------------------------ geometry

SHAPES = {f"{P}{tag}" for tag in ("sp", "grpSp", "pic", "graphicFrame",
                                   "cxnSp", "contentPart")}
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
#: Smaller than this share of the slide in both directions is a marker an
#: add-in hides on the slide (think-cell leaves several), not a shape.
TINY = 0.005


def _top_shapes(tree):
    for node in tree:
        if node.tag == f"{MC}AlternateContent":
            branch = next(iter(node), None)
            node = next((c for c in branch if c.tag in SHAPES), None) \
                if branch is not None else None
            if node is None:
                continue
        if node.tag in SHAPES:
            yield node


def _xfrm(shape):
    """``(off, ext)`` of a shape's own transform, or ``(None, None)``."""
    for path in (f"{P}spPr/{A}xfrm", f"{P}grpSpPr/{A}xfrm", f"{P}xfrm"):
        x = shape.find(path)
        if x is not None and x.find(f"{A}off") is not None:
            return x.find(f"{A}off"), x.find(f"{A}ext")
    return None, None


def _placeholder_box(pkg, part, ph):
    """Where the layout, else its master, puts a placeholder."""
    kind, idx = ph.get("type", "body"), ph.get("idx")
    for _ in range(2):
        nxt = next((t for _, k, t, _ in pkg.rels(part)
                    if k.endswith(("/slideLayout", "/slideMaster"))), None)
        if nxt is None:
            return None, None
        part = nxt
        fallback = (None, None)
        for shape in pkg.xml(part).iter(f"{P}sp"):
            other = shape.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
            if other is None:
                continue
            off, ext = _xfrm(shape)
            if off is None:
                continue
            if idx is not None and other.get("idx") == idx:
                return off, ext
            if other.get("type", "body") == kind and fallback[0] is None:
                fallback = (off, ext)
        if fallback[0] is not None:
            return fallback
    return None, None


def _boxes(pkg, slide):
    """Normalised boxes of the slide's top-level shapes.

    A placeholder without its own transform sits where its layout puts it,
    and is counted there: a divider is often nothing else.
    """
    width, height = pkg.slide_size()
    tree = pkg.xml(slide).find(f"{P}cSld/{P}spTree")
    out = []
    for shape in _top_shapes(tree):
        off, ext = _xfrm(shape)
        if off is None:
            ph = shape.find(f"./*/{P}nvPr/{P}ph")
            if ph is None:
                continue
            off, ext = _placeholder_box(pkg, slide, ph)
            if off is None:
                continue
        w, h = int(ext.get("cx")) / width, int(ext.get("cy")) / height
        if w < TINY and h < TINY:
            continue
        x, y = int(off.get("x")) / width, int(off.get("y")) / height
        out.append((x, y, x + w, y + h))
    return out


def _iou(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    union = ((a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1])
             - inter)
    if union <= 0:
        # Two lines or two points: the same when they nearly coincide.
        return 1.0 if all(abs(p - q) < 0.01 for p, q in zip(a, b)) else 0.0
    return inter / union


def similarity(one, two):
    """Share of shapes that pair up, greedily, by overlap. 0..1."""
    if not one and not two:
        return 1.0
    pairs = sorted(((_iou(a, b), i, j) for i, a in enumerate(one)
                    for j, b in enumerate(two)), reverse=True)
    used_a, used_b, matched = set(), set(), 0
    for score, i, j in pairs:
        if score < SAME_BOX:
            break
        if i in used_a or j in used_b:
            continue
        used_a.add(i)
        used_b.add(j)
        matched += 1
    return matched / max(len(one), len(two))


def _title(pkg, slide):
    for shape in pkg.xml(slide).iter(f"{P}sp"):
        ph = shape.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            return " ".join("".join(t.text or "" for t in p.iter(f"{A}t"))
                            for p in shape.iter(f"{A}p")).strip()
    return ""


# ---------------------------------------------------------------- candidates

def candidates(pack, decks, outdir):
    template = Package(pack.template)
    size = template.slide_size()
    designs = {n: _boxes(template, s)
               for n, s in enumerate(template.slide_parts(), start=1)}
    names = {template.layout_name(template.layout_of(s))
             for s in template.slide_parts()} | {
        template.layout_name(n) for n in template.parts
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)}

    known, foreign, new = [], [], []
    for deck in decks:
        pkg = Package(deck)
        if _size_differs(pkg.slide_size(), size):
            raise SystemExit(f"{deck}: slide size {pkg.slide_size()} differs "
                             f"from the template's {size} -- not built on "
                             f"this master")
        for number, slide in enumerate(pkg.slide_parts(), start=1):
            layout = pkg.layout_name(pkg.layout_of(slide))
            row = {"deck": str(Path(deck).resolve()), "slide": number,
                   "layout": layout, "title": _title(pkg, slide)}
            if layout not in names:
                foreign.append(row)
                continue
            boxes = _boxes(pkg, slide)
            best, score = max(((n, similarity(boxes, b)) for n, b in designs.items()),
                              key=lambda t: t[1], default=(None, 0.0))
            row.update(shapes=len(boxes), closest=best, match=round(score, 2))
            if score >= SAME_SLIDE:
                known.append(row)
            else:
                row["boxes"] = boxes
                new.append(row)

    # Group the new slides: single linkage, so a series whose members each
    # look like their neighbour lands in one group.
    groups = []
    for row in new:
        home = [g for g in groups
                if any(similarity(row["boxes"], other["boxes"]) >= SAME_SLIDE
                       for other in g)]
        merged = [row]
        for g in home:
            merged += g
            groups.remove(g)
        groups.append(merged)
    groups.sort(key=lambda g: (g[0]["deck"], min(r["slide"] for r in g)))
    for g in groups:
        g.sort(key=lambda r: (r["deck"], r["slide"]))
        for r in g:
            r.pop("boxes", None)

    outdir.mkdir(parents=True, exist_ok=True)
    data = {"pack": pack.name, "decks": [str(Path(d).resolve()) for d in decks],
            "candidates": [{"id": _letter(i), "slides": g}
                           for i, g in enumerate(groups)],
            "in_catalog": known, "other_master": foreign}
    (outdir / "candidates.json").write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    md = outdir / "candidates.md"
    md.write_text(_candidates_md(data), encoding="utf-8")
    return data, md


def _size_differs(one, two):
    return any(abs(a - b) > 0.005 * b for a, b in zip(one, two))


def _letter(i):
    out = ""
    i += 1
    while i:
        i, rem = divmod(i - 1, 26)
        out = chr(65 + rem) + out
    return out


def _short(deck):
    return Path(deck).name


def _candidates_md(data):
    lines = [f"# Design candidates — pack `{data['pack']}`", "",
             "Slides whose arrangement no design of the pack has. Slides that "
             "look alike are one candidate: pick **one slide per candidate** "
             "-- the cleanest, fullest example -- or skip the candidate.", "",
             "*Nearest design* is the closest the pack has, by how many shapes "
             "sit in the same place: at 50-75 % the slide is probably that "
             "design, reworked -- look at both before adding it.", "",
             "| Candidate | Deck | Slides | Layout | Shapes | Nearest design "
             "| First title |",
             "|---|---|---|---|---|---|---|"]
    for cand in data["candidates"]:
        first = cand["slides"][0]
        decks = sorted({_short(r["deck"]) for r in cand["slides"]})
        slides = ", ".join(str(r["slide"]) for r in cand["slides"])
        title = first["title"].replace("|", "/")[:70]
        near = (f"{first['closest']} ({first['match']:.0%})"
                if first.get("closest") else "—")
        lines.append(f"| **{cand['id']}** | {', '.join(decks)} | {slides} | "
                     f"{first['layout']} | {first['shapes']} | {near} | "
                     f"{title} |")
    if data["in_catalog"]:
        lines += ["", "## Already a design", "",
                  "| Deck | Slide | Closest design | Match |", "|---|---|---|---|"]
        for r in data["in_catalog"]:
            lines.append(f"| {_short(r['deck'])} | {r['slide']} | "
                         f"{r['closest']} | {r['match']:.0%} |")
    if data["other_master"]:
        lines += ["", "## Not on this master", "",
                  "Their layout is not in the template; they cannot become "
                  "designs of this pack.", "",
                  "| Deck | Slide | Layout |", "|---|---|---|"]
        for r in data["other_master"]:
            lines.append(f"| {_short(r['deck'])} | {r['slide']} | {r['layout']} |")
    return "\n".join(lines) + "\n"


# --------------------------------------------------------------------- texts

#: Placeholders that are the master's furniture, not the slide's content:
#: their text is a page property every deck sets for itself.
CHROME_PH = {"ftr", "sldNum", "dt", "hdr"}


def _paragraphs(root):
    """Every text paragraph of a slide, in document order."""
    return [p for p in root.iter(f"{A}p")
            if p.getparent() is not None
            and p.getparent().tag.endswith("}txBody")]


def _is_chrome(p):
    node = p
    while node is not None and node.tag != f"{P}sp":
        node = node.getparent()
    if node is None:
        return False
    ph = node.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
    return ph is not None and ph.get("type") in CHROME_PH


def _shape_name(node):
    while node is not None:
        if node.tag in (f"{P}sp", f"{P}graphicFrame", f"{P}grpSp"):
            c = node.find(f".//{P}cNvPr")
            if c is not None:
                return c.get("name", "")
        node = node.getparent()
    return ""


def _runs(p):
    return [r for r in p if r.tag == f"{A}r"]


def _ptext(p):
    out = []
    for node in p:
        if node.tag in (f"{A}r", f"{A}fld"):
            out.append("".join(t.text or "" for t in node.iter(f"{A}t")))
        elif node.tag == f"{A}br":
            out.append("\n")
    return "".join(out)


def _warnings(pkg, slide):
    found = []
    pictures = 0
    for _, kind, target, external in pkg.rels(slide):
        short = kind.rsplit("/", 1)[-1]
        if short in ("image", "media", "video", "audio"):
            pictures += 1
        elif short == "chart":
            found.append(f"a chart ({target}) -- its data travels with it")
        elif short in ("oleObject", "package"):
            found.append(f"an embedded object ({target}) -- its content "
                         f"travels with it")
        elif short.startswith("diagram"):
            if short == "diagramData":
                found.append("a SmartArt -- its text is not listed here")
    if pictures:
        found.append(f"{pictures} picture(s) -- they travel as they are; "
                     f"a customer photo or logo has to be replaced by hand")
    return found


def texts(pack, deck, numbers, outdir):
    template = Package(pack.template)
    pkg = Package(deck)
    if _size_differs(pkg.slide_size(), template.slide_size()):
        raise SystemExit(f"{deck}: slide size differs from the template's")
    order = pkg.slide_parts()
    layouts = _template_layouts(template)
    slides = []
    for number in numbers:
        if not 1 <= number <= len(order):
            raise SystemExit(f"{deck} has {len(order)} slides, not {number}")
        part = order[number - 1]
        layout = pkg.layout_name(pkg.layout_of(part))
        if layout not in layouts:
            raise SystemExit(f"slide {number}: layout {layout!r} is not in the "
                             f"template -- not built on this master")
        root = pkg.xml(part)
        items = []
        for i, p in enumerate(_paragraphs(root)):
            text = _ptext(p)
            if not text.strip() or _is_chrome(p):
                continue
            item = {"id": f"p{i}", "shape": _shape_name(p), "text": text,
                    "chars": len(text), "new": None}
            if p.find(f"{A}fld") is not None:
                item["field"] = True
            runs = _runs(p)
            if len(runs) > 1:
                item["runs"] = ["".join(t.text or "" for t in r.iter(f"{A}t"))
                                for r in runs]
            items.append(item)
        for i, c in enumerate(root.iter(f"{P}cNvPr")):
            for attr in ("descr", "title"):
                if c.get(attr):
                    items.append({"id": f"alt{i}.{attr}", "shape": c.get("name", ""),
                                  "kind": "alt text", "text": c.get(attr),
                                  "new": c.get(attr), "check": True})
        slides.append({"slide": number, "title": _title(pkg, part),
                       "layout": layout, "items": items,
                       "carries": _warnings(pkg, part)})

    outdir.mkdir(parents=True, exist_ok=True)
    data = {
        "pack": pack.name,
        "deck": str(Path(deck).resolve()),
        "deck_sha1": hashlib.sha1(Path(deck).read_bytes()).hexdigest(),
        "how": "Fill every `new` with sample text: same language, about the "
               "same length (`chars`), generic -- no customer, product, person "
               "or number from the project. A string replaces the whole "
               "paragraph (\\n is a line break). Where `runs` is listed, a "
               "list of the same length keeps each run's formatting. `field` "
               "paragraphs keep their field (a page number, a date). Alt texts "
               "are prefilled: check them. Then show the user and run append.",
        "slides": slides,
    }
    out = outdir / f"{Path(deck).stem}.texts.json"
    out.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n",
                   encoding="utf-8")
    return out, data


def _template_layouts(template):
    """``{layout name: [(layout part, names on its master)]}``."""
    out = {}
    for name in template.parts:
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", name):
            out.setdefault(template.layout_name(name), []).append(
                (name, template.master_layout_names(name)))
    return out


def _replace(p, new):
    runs = _runs(p)
    if isinstance(new, list):
        if len(new) != len(runs):
            raise ValueError(f"{len(new)} strings for {len(runs)} runs")
        for run, text in zip(runs, new):
            run.find(f"{A}t").text = text
        return
    for br in p.findall(f"{A}br"):
        p.remove(br)
    if not runs:
        return                                # a field-only paragraph
    first = runs[0]
    for run in runs[1:]:
        p.remove(run)
    lines = str(new).split("\n")
    first.find(f"{A}t").text = lines[0]
    anchor = first
    for line in lines[1:]:
        br = etree.Element(f"{A}br")
        rpr = first.find(f"{A}rPr")
        if rpr is not None:
            br.append(_copy(rpr))
        run = _copy(first)
        run.find(f"{A}t").text = line
        anchor.addnext(br)
        br.addnext(run)
        anchor = run


def _copy(node):
    return etree.fromstring(etree.tostring(node))


# -------------------------------------------------------------------- append

class Appender:
    """Copies slides of one source package into the template's parts."""

    def __init__(self, template, source):
        self.t, self.s = template, source
        self.copied = {}                       # source part -> template part
        self.added_types = []

    def _add_type(self, part, source_part):
        ctype, override = self.s.content_type(source_part)
        if ctype is None:
            raise SystemExit(f"{source_part}: no content type in the source")
        if override:
            node = etree.SubElement(self.t.types, f"{{{CT_NS}}}Override")
            node.set("PartName", "/" + part)
            node.set("ContentType", ctype)
            return
        mine, _ = self.t.content_type(part)
        if mine is None:
            node = etree.Element(f"{{{CT_NS}}}Default")
            node.set("Extension", part.rsplit(".", 1)[-1])
            node.set("ContentType", ctype)
            self.t.types.insert(0, node)
        elif mine != ctype:
            node = etree.SubElement(self.t.types, f"{{{CT_NS}}}Override")
            node.set("PartName", "/" + part)
            node.set("ContentType", ctype)

    def copy_part(self, source_part):
        """Copy one part, and everything it relates to, into the template."""
        if source_part in self.copied:
            return self.copied[source_part]
        data = self.s.parts[source_part]
        rels = self.s.rels(source_part)
        if not rels:
            # A leaf -- a picture, an embedding.  The template may hold the
            # very same bytes already (a brand icon); reuse that part.
            digest = hashlib.sha1(data).digest()
            folder = posixpath.dirname(source_part)
            for name, existing in self.t.parts.items():
                if (posixpath.dirname(name) == folder
                        and len(existing) == len(data)
                        and hashlib.sha1(existing).digest() == digest):
                    self.copied[source_part] = name
                    return name
        part = _free_name(self.t.parts, source_part)
        self.copied[source_part] = part
        self.t.parts[part] = data
        self._add_type(part, source_part)
        if rels:
            self._write_rels(part, source_part, rels, drop=set())
        return part

    def _write_rels(self, part, source_part, rels, drop, layout=None):
        root = etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})
        for rid, kind, target, external in rels:
            if rid in drop:
                continue
            node = etree.SubElement(root, f"{{{RELS_NS}}}Relationship")
            node.set("Id", rid)
            node.set("Type", kind)
            if external:
                node.set("Target", target)
                node.set("TargetMode", "External")
            elif kind.endswith("/slideLayout") and layout:
                node.set("Target", _rel_target(part, layout))
            else:
                node.set("Target", _rel_target(part, self.copy_part(target)))
        folder, base = posixpath.split(part)
        self.t.parts[posixpath.join(folder, "_rels", base + ".rels")] = \
            etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                           standalone=True)

    def append_slide(self, source_slide, replacements):
        layouts = _template_layouts(self.t)
        source_layout = self.s.layout_of(source_slide)
        name = self.s.layout_name(source_layout)
        options = layouts.get(name)
        if not options:
            raise SystemExit(f"layout {name!r} is not in the template")
        # Two layouts may share a name (a master copied in twice).  Prefer the
        # one on a master like the source's, then the one the template's own
        # designs use most -- that is the one its catalog was measured on.
        siblings = self.s.master_layout_names(source_layout)
        usage = {}
        for slide in self.t.slide_parts():
            used = self.t.layout_of(slide)
            usage[used] = usage.get(used, 0) + 1
        layout = max(options, key=lambda o: (len(o[1] & siblings),
                                             usage.get(o[0], 0)))[0]

        root = self.s.xml(source_slide)
        root.attrib.pop("show", None)          # a design is never hidden
        paragraphs = _paragraphs(root)
        names = list(root.iter(f"{P}cNvPr"))
        for item in replacements:
            ident, new = item["id"], item["new"]
            if ident.startswith("alt"):
                index, attr = ident[3:].split(".")
                node = names[int(index)]
                if node.get(attr) != item["text"]:
                    raise SystemExit(f"{ident}: the deck changed since `texts`")
                if new:
                    node.set(attr, new)
                else:
                    del node.attrib[attr]
                continue
            p = paragraphs[int(ident[1:])]
            if _ptext(p) != item["text"]:
                raise SystemExit(f"{ident}: the deck changed since `texts` -- "
                                 f"run texts again")
            _replace(p, new)

        rels = self.s.rels(source_slide)
        drop = {rid for rid, kind, _, ext in rels
                if not ext and kind.rsplit("/", 1)[-1] in DROPPED}
        if drop:
            _strip_references(root, drop)
        part = _free_name(self.t.parts, "ppt/slides/slide1.xml")
        self.t.parts[part] = etree.tostring(root, xml_declaration=True,
                                            encoding="UTF-8", standalone=True)
        node = etree.SubElement(self.t.types, f"{{{CT_NS}}}Override")
        node.set("PartName", "/" + part)
        node.set("ContentType", SLIDE_CT)
        self._write_rels(part, source_slide, rels, drop, layout=layout)
        self._register(part)
        return part

    def _register(self, part):
        pres_name = "ppt/presentation.xml"
        rels_name = "ppt/_rels/presentation.xml.rels"
        rels = etree.fromstring(self.t.parts[rels_name])
        used = {r.get("Id") for r in rels}
        n = 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        rel = etree.SubElement(rels, f"{{{RELS_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", SLIDE_TYPE)
        rel.set("Target", _rel_target(pres_name, part))
        self.t.parts[rels_name] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True)

        pres = etree.fromstring(self.t.parts[pres_name])
        lst = pres.find(f"{P}sldIdLst")
        if lst is None:
            lst = etree.Element(f"{P}sldIdLst")
            pres.find(f"{P}sldMasterIdLst").addnext(lst)
        ids = [int(s.get("id")) for s in lst] or [255]
        entry = etree.SubElement(lst, f"{P}sldId")
        entry.set("id", str(max(ids) + 1))
        entry.set(f"{R}id", rid)
        self.t.parts[pres_name] = etree.tostring(
            pres, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, path):
        self.t.parts["[Content_Types].xml"] = etree.tostring(
            self.t.types, xml_declaration=True, encoding="UTF-8",
            standalone=True)
        app = self.t.parts.get("docProps/app.xml")
        if app:
            count = len(self.t.slide_parts())
            self.t.parts["docProps/app.xml"] = re.sub(
                rb"<Slides>\d+</Slides>", f"<Slides>{count}</Slides>".encode(), app)
        staged = path.with_name(path.name + ".tmp")
        names = [n for n in self.t.order if n in self.t.parts]
        names += [n for n in self.t.parts if n not in set(names)]
        with zipfile.ZipFile(staged, "w", zipfile.ZIP_DEFLATED) as out:
            for name in names:
                out.writestr(name, self.t.parts[name])
        staged.replace(path)


def _strip_references(root, rids):
    """Remove what points at a dropped relationship: a hyperlink to another
    slide becomes no hyperlink."""
    for node in list(root.iter()):
        for key, value in list(node.attrib.items()):
            if key.startswith(R) and value in rids:
                if node.tag in (f"{A}hlinkClick", f"{A}hlinkHover"):
                    node.getparent().remove(node)
                    break
                del node.attrib[key]


def _clashes(data):
    """Pairs whose new text is equal, or one inside the other, where the
    original's was not.  Duplicates the original had are the design's own."""
    out = []
    for s in data["slides"]:
        items = [i for i in s["items"]
                 if not i["id"].startswith("alt") and isinstance(i["new"], str)
                 and i["new"].strip()]
        for n, one in enumerate(items):
            for two in items[n + 1:]:
                a, b = one["new"].strip(), two["new"].strip()
                was_a, was_b = one["text"].strip(), two["text"].strip()
                if (a in b or b in a) and not (was_a in was_b or was_b in was_a):
                    out.append(f"slide {s['slide']}: {one['id']} {a!r} / "
                               f"{two['id']} {b!r}")
    return out


def _ledger_path(pack):
    return pack.catalog_dir / "added.json"


def append(pack, text_files, again=False):
    if pack.template is None or pack.template.suffix.lower() not in (".pptx", ".potx"):
        raise SystemExit(f"pack {pack.name}: adding designs needs a .pptx or "
                         f".potx template, this pack has {pack.template}")
    jobs = []
    for path in text_files:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        if data.get("pack") != pack.name:
            raise SystemExit(f"{path} was written for pack {data.get('pack')!r}")
        deck = Path(data["deck"])
        if hashlib.sha1(deck.read_bytes()).hexdigest() != data["deck_sha1"]:
            raise SystemExit(f"{deck} changed since {path} was written -- run "
                             f"texts again")
        missing = [f"slide {s['slide']} {i['id']}" for s in data["slides"]
                   for i in s["items"] if i.get("new") is None]
        if missing:
            raise SystemExit(f"{path}: no sample text yet for "
                             + ", ".join(missing[:8])
                             + (f" and {len(missing) - 8} more" if len(missing) > 8 else "")
                             + " -- fill every `new` first")
        clashes = _clashes(data)
        if clashes:
            raise SystemExit(f"{path}: sample text that the original kept "
                             f"apart now collides -- a retext key is a "
                             f"substring, so it would hit both:\n  "
                             + "\n  ".join(clashes[:10])
                             + "\nGive each paragraph its own wording.")
        jobs.append((path, data))

    ledger_path = _ledger_path(pack)
    ledger = (json.loads(ledger_path.read_text(encoding="utf-8"))
              if ledger_path.is_file() else {"added": []})
    seen = {(e["deck_sha1"], e["slide"]): e["design"] for e in ledger["added"]}
    for _, data in jobs:
        for s in data["slides"]:
            hit = seen.get((data["deck_sha1"], s["slide"]))
            if hit and not again:
                raise SystemExit(f"{Path(data['deck']).name} slide {s['slide']} "
                                 f"is design {hit} already (--again adds it twice)")

    template = Package(pack.template)
    first = len(template.slide_parts()) + 1
    added = []
    for path, data in jobs:
        source = Package(data["deck"])
        appender = Appender(template, source)
        order = source.slide_parts()
        for s in data["slides"]:
            appender.append_slide(order[s["slide"] - 1], s["items"])
            added.append({"design": first + len(added),
                          "deck": Path(data["deck"]).name,
                          "deck_sha1": data["deck_sha1"], "slide": s["slide"],
                          "added": time.strftime("%Y-%m-%d")})

    history = pack.root / "history"
    history.mkdir(exist_ok=True)
    backup = history / (f"{pack.template.stem}-{time.strftime('%Y%m%d-%H%M%S')}"
                        f"{pack.template.suffix}")
    shutil.copy2(pack.template, backup)
    appender.save(pack.template)
    ledger["added"] += added
    ledger_path.parent.mkdir(parents=True, exist_ok=True)
    ledger_path.write_text(json.dumps(ledger, indent=2, ensure_ascii=False) + "\n",
                           encoding="utf-8")
    return added, backup


# ---------------------------------------------------------------------- main

def _numbers(spec):
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += range(int(a), int(b) + 1)
        elif part:
            out.append(int(part))
    return out


def _pack(name):
    pack = pack_named(name)
    if pack is None or pack.template is None:
        raise SystemExit(f"no installed pack {name!r} with a template -- run "
                         f"the setup stages first")
    return pack


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("candidates", help="which slides are new designs")
    c.add_argument("pack")
    c.add_argument("decks", nargs="+", type=Path)
    c.add_argument("-o", "--outdir", type=Path, default=Path(OUT_DIR))
    t = sub.add_parser("texts", help="list the chosen slides' text to neutralise")
    t.add_argument("pack")
    t.add_argument("deck", type=Path)
    t.add_argument("--slides", required=True, help="e.g. 5,12,20-21")
    t.add_argument("-o", "--outdir", type=Path, default=Path(OUT_DIR))
    a = sub.add_parser("append", help="add the slides to the pack's template")
    a.add_argument("pack")
    a.add_argument("texts", nargs="+", type=Path)
    a.add_argument("--again", action="store_true",
                   help="add a slide that is a design already")
    args = ap.parse_args()
    pack = _pack(args.pack)

    if args.cmd == "candidates":
        for deck in args.decks:
            if deck.suffix.lower() not in (".pptx", ".potx"):
                ap.error(f"{deck}: .pptx only -- convert an .odp with "
                         f"convert.py first")
        data, md = candidates(pack, args.decks, args.outdir)
        print(f"Wrote {md}")
        print(f"{len(data['candidates'])} candidate(s), "
              f"{len(data['in_catalog'])} slide(s) already a design, "
              f"{len(data['other_master'])} on another master.")
        print("Show the table to the user; for each candidate they want, take "
              "one slide. Then:\n  add_designs.py texts "
              f"{pack.name} <deck> --slides <numbers>")
        return 0

    if args.cmd == "texts":
        out, data = texts(pack, args.deck, _numbers(args.slides), args.outdir)
        count = sum(len(s["items"]) for s in data["slides"])
        print(f"Wrote {out}: {count} text(s) on {len(data['slides'])} slide(s).")
        for s in data["slides"]:
            for line in s["carries"]:
                print(f"  slide {s['slide']} carries {line}")
        print("Fill every `new` with neutral sample text of about the same "
              "length, show the user, then:\n  add_designs.py append "
              f"{pack.name} {out}")
        return 0

    added, backup = append(pack, args.texts, again=args.again)
    print(f"Added {len(added)} design(s) to {pack.template}: "
          + ", ".join(f"{a['design']} (from {a['deck']} slide {a['slide']})"
                      for a in added))
    print(f"The template before: {backup}")
    print("Next: re-extract the catalog, brief the vision-agent on the new "
          "designs, and fill the worksheet for them only:\n"
          f"  catalog_extract.py {pack.template} -o {pack.catalog_dir}\n"
          f"  vision_brief.py designs {pack.catalog_dir}\n"
          f"  name_designs.py {pack.catalog_dir} --only "
          + ",".join(str(a["design"]) for a in added))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
