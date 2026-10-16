"""Iterate a shape's paragraphs and runs over UNO -- optionally without ever
releasing the proxies while LibreOffice runs.

Why: LibreOffice <= 25.2.x (at least; upstream fix pending) crashes with
"free(): invalid pointer" when a remote client drops paragraph/portion proxies
while it enumerates the same shape again.  The release arrives as an
asynchronous URP message that soffice executes on its own thread without the
SolarMutex; the next createEnumeration() can pick up a paragraph object that
release is deleting, and frees it a second time.

With PRESENTATION_PIN_TEXT_PROXIES=1 (make_deck.py --pin-text-proxies) every
enumeration, paragraph and run handed out here is kept alive until the
LibreOffice process has stopped (impress._Server.stop calls release_pins()).
No proxy is released into a live soffice, so neither race can happen.  Cost:
the proxies of one soffice lifetime stay in memory -- a few MB per deck.
"""

import os

PIN_ENV = "PRESENTATION_PIN_TEXT_PROXIES"

_PINS = []


def pinning() -> bool:
    return os.environ.get(PIN_ENV, "").lower() in ("1", "true", "yes", "on")


def pin(obj):
    """Keep ``obj`` alive until release_pins(); returns it.  No-op when off."""
    if obj is not None and pinning():
        _PINS.append(obj)
    return obj


def release_pins() -> int:
    """Drop all pinned proxies.  Only call once soffice is gone."""
    n = len(_PINS)
    _PINS.clear()
    return n


def enumerate_text(obj):
    """Yield the elements of ``obj.createEnumeration()`` (paragraphs of a
    shape/text, or runs of a paragraph), pinning the enumeration and every
    element when pinning is on.

    A generator: an ``obj`` without ``createEnumeration`` raises on the first
    ``next()``, not on the call, so wrap the whole loop in any ``try``."""
    enum = pin(obj.createEnumeration())
    while enum.hasMoreElements():
        yield pin(enum.nextElement())
