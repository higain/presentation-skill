"""Resources a deck would fetch from elsewhere when LibreOffice opens it.

A slide can carry an image, a video or an embedded object by URL instead of
inside the file.  LibreOffice fetches it on load -- for an untrusted deck,
that is the author learning who opened it and when, or a request into the
intranet.  ``office.soffice`` runs LibreOffice without a network, which is the
real defence; this is the second one, for a machine that cannot isolate:
it names what a deck would fetch so a caller can refuse it.

Hyperlinks are not listed: clicking one is a person's decision, and opening
the file follows none of them.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

#: A target that leaves the machine.  ``file:`` is included because a UNC or
#: an SMB path resolves over the network too.
REMOTE = re.compile(r"^\s*(?:[a-z][a-z0-9+.-]*:)?//|^\s*(?:https?|ftp|smb|file):",
                    re.I)

_OOXML_REL = re.compile(
    rb'<Relationship\b[^>]*?Type="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?'
    rb'TargetMode="External"|<Relationship\b[^>]*?TargetMode="External"'
    rb'[^>]*?Type="([^"]+)"[^>]*?Target="([^"]+)"', re.S)
_ODF_HREF = re.compile(
    rb'<(?:draw:image|draw:object|draw:object-ole|draw:plugin|draw:applet)\b'
    rb'[^>]*?xlink:href="([^"]+)"', re.S)

#: Archives larger than this are not scanned: the scan reads XML parts only,
#: and a deck whose XML alone is this big is not a deck.
MAX_XML_BYTES = 200 * 1024 * 1024


def external_resources(path) -> list[tuple[str, str]]:
    """``[(part, url)]`` for everything opening ``path`` would fetch."""
    path = Path(path)
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return []                       # .ppt, .fodp: not a zip, not scanned
    found = []
    total = 0
    with archive:
        for info in archive.infolist():
            name = info.filename
            if not (name.endswith(".rels") or name.endswith(".xml")):
                continue
            total += info.file_size
            if total > MAX_XML_BYTES:
                break
            data = archive.read(name)
            if name.endswith(".rels"):
                for m in _OOXML_REL.finditer(data):
                    kind = (m.group(1) or m.group(3)).decode(errors="replace")
                    target = (m.group(2) or m.group(4)).decode(errors="replace")
                    if kind.rsplit("/", 1)[-1] == "hyperlink":
                        continue
                    if REMOTE.search(target):
                        found.append((name, target))
            else:
                for m in _ODF_HREF.finditer(data):
                    target = m.group(1).decode(errors="replace")
                    if REMOTE.search(target):
                        found.append((name, target))
    return found
