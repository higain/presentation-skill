"""Per-slide QA results kept between runs: re-check only the slides that changed.

A 58-slide deck spent nine minutes in QA on every gate run, and an agent fixing
one slide ran the gate eight times -- an hour re-measuring 57 slides nobody had
touched.  The per-slide checks (overflow, collision, off-slide, contrast,
occlusion, buried text, type size, titles, paint metrics) depend only on the
slide, its layout and the template, so their findings are stored under a
fingerprint of exactly that and reused while it holds.

The deck-wide checks -- template wording left behind, footers, design variety,
prose, approved titles -- always run: they are cheap and they compare slides
with each other.

The fingerprint is the slide's XML with the parts a save renumbers stripped
(shape ids, relationship ids), plus the content of everything it links to
except its speaker notes.  It is salted with the QA engine's own source and the
template file, so a change to either invalidates every entry.  .pptx only; an
.odp keeps all its slides in one part and is always checked in full.
"""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from pathlib import Path

CACHE_SUFFIX = "-qa-cache.json"

#: Code whose change can change a finding.
_ENGINE = ("qa/*.py", "template.py", "draw.py", "impress.py", "colors.py")


def cache_path(deck) -> Path:
    deck = Path(deck)
    return deck.with_name(deck.stem + CACHE_SUFFIX)


def engine_key(template=None) -> str:
    """Hash of the QA code and the template: what every entry is valid for."""
    here = Path(__file__).resolve().parent.parent
    digest = hashlib.sha256()
    for pattern in _ENGINE:
        for path in sorted(here.glob(pattern)):
            digest.update(path.name.encode())
            digest.update(path.read_bytes())
    if template:
        try:
            digest.update(Path(template).read_bytes())
        except OSError:
            pass
    return digest.hexdigest()


_IDS = re.compile(rb'\s(?:id|r:id|r:embed|r:link|r:pict)="[^"]*"')


def _resolve(base, target):
    parts = base.rsplit("/", 1)[0].split("/")
    for seg in target.split("/"):
        if seg == "..":
            parts.pop()
        elif seg and seg != ".":
            parts.append(seg)
    return "/".join(parts)


def slide_fingerprints(deck):
    """``[hex]`` per slide in presentation order, or ``None`` if not a .pptx."""
    try:
        z = zipfile.ZipFile(deck)
    except (OSError, zipfile.BadZipFile):
        return None
    with z:
        names = set(z.namelist())
        if "ppt/presentation.xml" not in names:
            return None

        def rels(part):
            path = "/".join(part.split("/")[:-1]) + "/_rels/" + part.split("/")[-1] + ".rels"
            if path not in names:
                return []
            xml = z.read(path).decode("utf-8", "replace")
            out = []
            for rel in re.findall(r"<Relationship\b[^>]*>", xml):
                rid = re.search(r'\bId="([^"]+)"', rel)
                target = re.search(r'\bTarget="([^"]+)"', rel)
                kind = re.search(r'\bType="([^"]+)"', rel)
                external = 'TargetMode="External"' in rel
                if rid and target:
                    out.append((rid.group(1), target.group(1),
                                kind.group(1) if kind else "", external))
            return out

        pres_rels = {rid: target for rid, target, _, _ in rels("ppt/presentation.xml")}
        order = re.findall(rb'<p:sldId\b[^>]*\br:id="([^"]+)"',
                           z.read("ppt/presentation.xml"))
        prints = []
        for rid in order:
            part = _resolve("ppt/presentation.xml", pres_rels[rid.decode()])
            digest = hashlib.sha256(_IDS.sub(b"", z.read(part)))
            linked = []
            for _, target, kind, external in rels(part):
                if kind.endswith("/notesSlide"):
                    continue
                if external:
                    linked.append("ext:" + target)
                    continue
                path = _resolve(part, target)
                if path in names:
                    linked.append(hashlib.sha256(
                        _IDS.sub(b"", z.read(path))).hexdigest())
            for item in sorted(linked):
                digest.update(item.encode())
            prints.append(digest.hexdigest())
        return prints


def load(deck, engine):
    """``{fingerprint: entry}`` saved for this deck and engine, else ``{}``."""
    try:
        raw = json.loads(cache_path(deck).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if raw.get("engine") != engine:
        return {}
    return raw.get("slides") or {}


def save(deck, engine, entries):
    try:
        cache_path(deck).write_text(json.dumps(
            {"engine": engine, "slides": entries}, ensure_ascii=False),
            encoding="utf-8")
    except OSError:
        pass


def pack_finding(number, text, key):
    """A finding of slide ``number`` in slide-independent form, or ``None``."""
    prefix = f"slide {number}"
    if not text.startswith(prefix) or text[len(prefix):][:1] not in (" ", ":"):
        return None
    if key is not None:
        key = list(key)
        if len(key) > 1 and key[1] == number:
            key[1] = "#"
    return [text[len(prefix):], key]


def unpack_finding(number, packed):
    """``(text, key)`` for slide ``number`` from :func:`pack_finding`'s form."""
    rest, key = packed
    if key is not None:
        key = tuple(number if (i == 1 and k == "#") else k
                    for i, k in enumerate(key))
    return f"slide {number}{rest}", key
