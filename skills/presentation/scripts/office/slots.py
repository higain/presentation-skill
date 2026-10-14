"""How many LibreOffice instances this skill runs at once, machine-wide.

Every build, check and render starts its own headless LibreOffice, and a big
deck is built by several builder subagents at the same time.  A field run
started ten of them on a 4-core, 8 GB machine: ten LibreOffice instances
rendering at once, and the machine went down with the run.  Telling the agent
to start fewer is the first defence; this is the one that holds when it does
not listen.

A **slot** is one of ``limit()`` lock files.  A process takes one before its
first LibreOffice and holds it until it exits -- the kernel drops an
``flock`` with the process, so a killed build cannot leave a slot taken.  A
process that finds every slot taken waits and says so.

One slot per *job*, not per instance: ``make_deck.py`` takes it up front and
its children (the build script, the renderer) inherit it through
:data:`HELD_ENV` instead of each waiting for a slot of their own -- which, with
the parent holding one and waiting on them, could wait forever.
"""

import fcntl
import os
import sys
import tempfile
import time
from pathlib import Path

#: Overrides the computed limit: a positive integer.
LIMIT_ENV = "PRESENTATION_MAX_OFFICE"

#: Set while this process -- or the one that started it -- holds a slot.
HELD_ENV = "PRESENTATION_OFFICE_SLOT"

_held = None


def limit() -> int:
    """How many may run at once: half the cores, one per 3 GB of RAM, at least one.

    Measured on the machine that went down: 4 cores and 7.8 GB give 2.  One
    instance rendering a 60-slide deck peaks around a gigabyte and keeps a core
    busy; the rest of the machine is running the agents.
    """
    raw = os.environ.get(LIMIT_ENV, "").strip()
    if raw.isdigit() and int(raw) > 0:
        return int(raw)
    cores = os.cpu_count() or 2
    try:
        gigabytes = os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES") / 2**30
    except (ValueError, OSError):
        gigabytes = 8
    return max(1, min(cores // 2, int(gigabytes // 3)))


def _directory() -> Path:
    path = Path(tempfile.gettempdir()) / f"presentation-office-slots-{os.getuid()}"
    path.mkdir(mode=0o700, exist_ok=True)
    return path


def acquire() -> None:
    """Take a slot for this process, waiting for one if all are taken.

    Does nothing if this process or its parent already holds one.
    """
    global _held
    if _held is not None or os.environ.get(HELD_ENV):
        return
    try:
        directory = _directory()
    except OSError:
        return                       # no shared tmp: run unlimited, as before
    count, said = limit(), False
    while True:
        for i in range(count):
            fd = os.open(directory / f"slot-{i}", os.O_RDWR | os.O_CREAT, 0o600)
            try:
                fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                os.close(fd)
                continue
            _held = fd
            os.environ[HELD_ENV] = str(os.getpid())
            if said:
                print("[office] slot free, starting.", file=sys.stderr, flush=True)
            return
        if not said:
            print(f"[office] waiting: {count} LibreOffice job(s) already run on "
                  f"this machine, the limit ({LIMIT_ENV} changes it). This one "
                  f"starts when one of them finishes -- do not start another.",
                  file=sys.stderr, flush=True)
            said = True
        time.sleep(1)
