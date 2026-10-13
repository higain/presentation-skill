"""What the template itself already does, measured once and cached, so it is not reported as a defect."""

from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path

from ..impress import load_document, slides
from ..template import colliding, footer_text, offslide_baseline
from .colour import _contrast_findings, _occlusion_findings
from .text import MIN_LEFTOVER, _page_texts

class _Baseline:
    """What the template already does, so a deck is not blamed for inheriting it.

    Same principle as ``offslide_baseline``: a brand template that sets white
    on orange, or stacks two shapes on its own design 17, hands both to every
    deck built from it.  Reporting them is reporting work nobody did, and a
    check that does that gets ignored -- including on the lines that matter.

    Colour pairs are baselined whatever their ratio.  A pairing the template
    ships is a house style decision you inherit, not a defect this build
    introduced; a pairing it does not ship -- black on a navy divider -- is
    yours, and still fires.
    """

    def __init__(self, strings=(), bleeds=(), contrast=(), occlusion=(),
                 collisions=(), footers=(), by_design=None):
        self.strings = set(strings)
        # The same wording, per template slide -- which is per design, since
        # design N is template slide N.  See ``leftover_strings``.
        self.by_design = {int(k): set(v) for k, v in (by_design or {}).items()}
        self.bleeds = set(bleeds)
        self.contrast = set(contrast)
        self.occlusion = set(occlusion)
        self.collisions = set(collisions)
        self.footers = set(footers)

    def leftover_strings(self, design=None):
        """The template wording a slide built on ``design`` can still carry.

        Only its own design's.  Text a build drew itself is content, even when
        another design happens to hold the same word: one run drew a phase bar
        ("Design", "Testen", "Bereitstellen") on 28 canvas slides, the words
        were the sample labels of an unrelated design 18, and 112 findings --
        each one an allowlist entry -- buried the report.  A label the slide's
        own design ships is still caught, and that is the defect this check
        exists for.  No recorded design (a deck edited, not built) falls back
        to the whole template.
        """
        if design is None or design not in self.by_design:
            return self.strings
        return self.by_design[design]


def _collision_key(reason):
    """The two texts a collision names, order-independent."""
    return frozenset(re.findall(r"'((?:[^']|\\')*)'", reason)[:2])


#: Where a computed baseline is remembered between runs.  Override with
#: ``$PRESENTATION_CACHE_DIR``; set ``$PRESENTATION_NO_CACHE`` to skip it.
CACHE_DIR_ENV = "PRESENTATION_CACHE_DIR"
NO_CACHE_ENV = "PRESENTATION_NO_CACHE"


def _cache_dir():
    override = os.environ.get(CACHE_DIR_ENV)
    if override:
        return Path(override)
    base = os.environ.get("XDG_CACHE_HOME") or (Path.home() / ".cache")
    return Path(base) / "presentation-skill"


def _baseline_key(path):
    """A cache key over the template *and* the code that measures it.

    Hashing the checker sources as well as the template means editing a rule in
    ``qa.py`` or ``office/template.py`` invalidates every cached baseline by
    itself.  A cache that can serve an answer the current code would not
    produce is worse than no cache: it hides exactly the finding someone just
    taught the checker to make.
    """
    h = hashlib.sha256()
    h.update(Path(path).read_bytes())
    # Every engine module, not a hand-kept list: the baseline also depends on
    # draw.contrast_ratio and impress, and a list is how those were missed.
    office = Path(__file__).resolve().parent.parent
    for src in sorted(office.rglob("*.py")):
        try:
            h.update(src.read_bytes())
        except OSError:                     # running from a zip or read-only
            return None
    return h.hexdigest()[:32]


def _baseline_load(cache_file):
    try:
        d = json.loads(cache_file.read_text())
        return _Baseline(
            strings=d["strings"],
            bleeds={(b[0], b[1]) for b in d["bleeds"]},
            contrast={(c[0], c[1]) for c in d["contrast"]},
            occlusion={(o[0], o[1]) for o in d["occlusion"]},
            collisions={frozenset(c) for c in d["collisions"]},
            footers=d.get("footers", ()),
            by_design=d["by_design"],
        )
    except (OSError, ValueError, KeyError, TypeError, IndexError):
        return None                         # a cache we cannot read is a miss


def _baseline_save(cache_file, base):
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        # Written via a temp file: two make_deck runs in parallel must never
        # leave a half-written cache for a third to read.
        tmp = cache_file.with_suffix(f".{os.getpid()}.tmp")
        tmp.write_text(json.dumps({
            "strings": sorted(base.strings),
            "bleeds": [list(b) for b in base.bleeds],
            "contrast": [list(c) for c in base.contrast],
            "occlusion": [list(o) for o in base.occlusion],
            "collisions": [sorted(c) for c in base.collisions],
            "footers": sorted(base.footers),
            "by_design": {str(k): sorted(v)
                          for k, v in base.by_design.items()},
        }))
        tmp.replace(cache_file)
    except OSError:
        pass                                # a cache that cannot write is fine


def _template_baseline(path):
    """Everything the template already does, in one document load.

    **Cached on disk**, because it is a pure function of the template and it is
    expensive: measured on the 21-design house template it is 37-43 s, against
    29 s to check a 10-slide deck -- well over half of a QA run, paid again on
    every build iteration.  The key covers the template bytes and the checker
    sources, so any change to either recomputes it.
    """
    if not path:
        return _Baseline()

    key = None if os.environ.get(NO_CACHE_ENV) else _baseline_key(path)
    cache_file = _cache_dir() / f"baseline-{key}.json" if key else None
    if cache_file is not None:
        cached = _baseline_load(cache_file)
        if cached is not None:
            return cached

    doc = load_document(path, read_only=True)
    try:
        strings, contrast, occlusion, collisions = set(), set(), set(), set()
        footers, by_design = set(), {}
        for number, page in enumerate(slides(doc), start=1):
            own = by_design.setdefault(number, set())
            for _, text in _page_texts(page):
                if len(text) >= MIN_LEFTOVER and any(c.isalpha() for c in text):
                    strings.add(text)
                    own.add(text)
            foot = footer_text(page)
            if foot.strip():
                footers.add(foot.strip())
            contrast.update(pair for pair, *_ in _contrast_findings(page))
            occlusion.update(pair for pair, *_ in _occlusion_findings(page))
            collisions.update(_collision_key(reason)
                              for _, reason in colliding(doc, page))
    finally:
        doc.close(False)

    base = _Baseline(strings, offslide_baseline(path), contrast, occlusion,
                     collisions, footers, by_design)
    if cache_file is not None:
        _baseline_save(cache_file, base)
    return base
