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
                   "no-title", "variety", "vision"}


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
    path = _allowlist_path(deck)
    if path is None or not path.exists():
        return {}
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

    out = {}
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
        out.setdefault(kind, {}).setdefault(slide, {})[content] = reason
    return out


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
        for slide, by_content in by_slide.items():
            for content in by_content:
                if (kind, slide, content) not in present:
                    out.append(
                        f"allowlist {kind!r} slide {slide} {content!r} matches "
                        f"nothing in this deck -- it was fixed, or never "
                        f"there. Remove it from the allowlist")
    return out
