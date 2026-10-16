"""Per-deck acknowledgements: findings recorded as intentional, with a reason."""

from __future__ import annotations

import json
from pathlib import Path

from ..impress import PRESENTATION_EXTS

#: Sidecar file, beside the deck, that lets a build acknowledge a finding as
#: intentional and get a clean report without hacking content.  See
#: ``load_allowlist`` and the QA section of SKILL.md.  A deck has no right to a
#: clean report just because its author said so: an acknowledgement only counts
#: when it matches a real finding by kind + slide + content, and carries a
#: non-empty reason.  Stale acknowledgements -- ones that match nothing in the
#: deck -- are themselves reported, so the list cannot silently grow into a
#: blanket "pass everything" file.
ALLOWLIST_SUFFIX = "-allowlist.json"


#: The finding kinds an allowlist entry can name.  Kept explicit so a typo in a
#: deck's allowlist is a hard error rather than a silent no-op.
ALLOWLIST_KINDS = {"leftover-wording", "stray-image", "footer", "balance",
                   "no-title", "variety", "alignment", "vision"}


class AllowlistError(ValueError):
    """An allowlist file is malformed, so the deck cannot be gated on it."""


def _allowlist_path(deck):
    """``<deck>-allowlist.json`` beside the deck, or ``None``."""
    deck = Path(deck)
    if deck.suffix.lower() in PRESENTATION_EXTS:
        return deck.with_name(deck.stem + ALLOWLIST_SUFFIX)
    return None


def load_allowlist(deck):
    """``{kind: {slide: {content: reason}}}`` for ``<deck>-allowlist.json``.

    ``[]`` (an empty list, a missing file, a file that is an empty JSON array)
    is the same as no allowlist.  A file that exists but is malformed raises
    :class:`AllowlistError`, because an allowlist that cannot be read must not
    silently let every finding through nor silently gate every one.

    Each entry:

    .. code-block:: json

        {"slide": 16, "kind": "leftover-wording",
         "content": "Testen",
         "reason": "SDLC phase label on the lifecycle circle -- correct content"}

    ``slide`` is the 1-based output-slide number.  ``content`` is the exact
    text (for ``leftover-wording`` and ``footer``) or the media filename (for
    ``stray-image``) that the finding names.  Matching is exact on all three of
    slide + kind + content; there is no wildcard, so an entry cannot
    accidentally cover a finding the author had not seen.
    """
    out = {}
    for entry in _entries(deck):
        out.setdefault(entry["kind"], {}).setdefault(
            entry["slide"], {})[entry["content"]] = entry["reason"].strip()
    return out


def vision_overrules(deck):
    """The ``vision`` entries of ``<deck>-allowlist.json``, as dicts.

    Each has ``slide``, ``content`` (a question such as ``"Q7"``, or
    ``"seen"`` for a slide a person looked at in place of the vision-agent),
    ``reason`` and ``approved_by`` -- empty until the user has approved it.
    An overrule counts only once approved: the build model that wants the
    gate to pass is the one writing the entry, and one run overruled a
    reviewer who was right ("the probe shows the title is there") while the
    title sat invisible under a picture.
    """
    return [{"slide": e["slide"], "content": e["content"],
             "reason": e["reason"].strip(),
             "approved_by": str(e.get("approved_by") or "").strip()}
            for e in _entries(deck) if e["kind"] == "vision"]


def _entries(deck):
    """The validated entries of ``<deck>-allowlist.json``; see ``load_allowlist``."""
    path = _allowlist_path(deck)
    if path is None or not path.exists():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        raise AllowlistError(f"{path.name}: not readable JSON ({e})") from e

    if isinstance(raw, dict) and "acknowledgements" in raw:
        entries = raw["acknowledgements"]
    elif isinstance(raw, list):
        entries = raw
    else:
        raise AllowlistError(
            f"{path.name}: expected a list, or an object with an "
            f"'acknowledgements' list")

    for i, entry in enumerate(entries):
        if not isinstance(entry, dict):
            raise AllowlistError(
                f"{path.name}: entry {i} is not an object")
        kind = entry.get("kind")
        slide = entry.get("slide")
        content = entry.get("content")
        reason = entry.get("reason", "").strip()
        if kind not in ALLOWLIST_KINDS:
            raise AllowlistError(
                f"{path.name}: entry {i} kind {kind!r} not one of "
                f"{sorted(ALLOWLIST_KINDS)}")
        if not isinstance(slide, int) or slide < 1:
            raise AllowlistError(
                f"{path.name}: entry {i} needs a positive integer 'slide'")
        if not content or not isinstance(content, str):
            raise AllowlistError(
                f"{path.name}: entry {i} needs a non-empty 'content' string")
        if not reason:
            raise AllowlistError(
                f"{path.name}: entry {i} needs a non-empty 'reason' -- an "
                f"acknowledgement with no reason is a way to ignore the check")
    return entries


def _acknowledged(allowlist, kind, slide, content):
    """The reason if ``allowlist`` covers this finding, else ``None``."""
    return allowlist.get(kind, {}).get(slide, {}).get(content)


def stale_acknowledgements(allowlist, findings):
    """``[message]`` for allowlist entries that match nothing in the deck.

    ``findings`` is the list of :class:`Finding`-shaped ``(kind, slide,
    content)`` tuples actually produced.  An entry that matches none of them
    means the finding it was written for no longer exists (good -- it was
    fixed) or was never there (the author guessed).  Either way it should be
    removed, so it is reported rather than silently left to accumulate.
    """
    present = {(kind, slide, content)
               for kind, slide, content in findings}
    out = []
    for kind, by_slide in allowlist.items():
        if kind == "vision":
            continue        # answers, not QA findings: the gate matches these
        for slide, by_content in by_slide.items():
            for content in by_content:
                if (kind, slide, content) not in present:
                    out.append(
                        f"allowlist {kind!r} slide {slide} {content!r} matches "
                        f"nothing in this deck -- it was fixed, or never "
                        f"there. Remove it from the allowlist")
    return out
