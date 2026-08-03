"""Shared setup for everything under ``office`` -- currently, where temp goes.

Importing any ``office`` module runs this first, so every entry point
(``qa.py``, ``render.py``, ``thumbnail.py``, ``make_deck.py``, a build script)
gets the same answer without having to know about it.

**Why this exists.**  This skill writes a lot of temporary data: a disposable
LibreOffice profile per run, a PDF between a deck and its PNGs, an unzipped
``.pptx`` for the validator.  All of it goes through ``tempfile``, which means
``/tmp`` unless told otherwise -- and on a locked-down machine ``/tmp`` is not
writable.  Python's own fallback would quietly scatter the files across the
working directory, but **LibreOffice is a separate process with no such
fallback**: it aborts with "User installation could not be completed" and
converts nothing, which surfaces as every render and every QA run failing at
once for no visible reason.

So: pick a writable directory once, point both ``tempfile`` and the
environment at it, and child processes inherit it.

Order of preference, first that works:

1. ``$PRESENTATION_TMPDIR`` -- an explicit choice, and errors are loud.
2. Whatever ``tempfile`` would use anyway (``$TMPDIR``, else ``/tmp``), if a
   real write succeeds there.  The common case changes nothing.
3. ``.presentation-tmp/`` in the working directory.

Set ``PRESENTATION_TMPDIR`` to put it somewhere specific::

    PRESENTATION_TMPDIR=./build/tmp python3 scripts/make_deck.py build.py out.pptx
"""

import atexit
import os
import tempfile
from pathlib import Path

ENV_OVERRIDE = "PRESENTATION_TMPDIR"
LOCAL_DIRNAME = ".presentation-tmp"

#: The directory temporary files are going to, once resolved.
TEMP_DIR = None


def _writable(path) -> bool:
    """Whether a directory can actually be created and written inside ``path``.

    Attempts a real write.  ``os.access`` reports the permission bits and is
    wrong about read-only mounts, full filesystems, ACLs and container
    overlays -- all of which are how ``/tmp`` is usually taken away.
    """
    try:
        base = Path(path)
        base.mkdir(parents=True, exist_ok=True)
        probe = tempfile.mkdtemp(prefix=".probe-", dir=str(base))
        os.rmdir(probe)
        return True
    except OSError:
        return False


def use_temp_dir(path) -> str:
    """Send ``tempfile`` **and every child process** to ``path``.

    Setting ``tempfile.tempdir`` covers this interpreter; the environment
    variables cover LibreOffice, which is spawned with a copy of ``os.environ``
    and does its own temp handling.  Both halves are needed.
    """
    resolved = str(Path(path).expanduser().resolve())
    Path(resolved).mkdir(parents=True, exist_ok=True)
    tempfile.tempdir = resolved
    for var in ("TMPDIR", "TMP", "TEMP"):
        os.environ[var] = resolved
    return resolved


def _default_temp_dir():
    """What ``tempfile`` would pick on its own, without committing to it."""
    try:
        return tempfile.gettempdir()
    except (OSError, FileNotFoundError):
        return None


def _resolve() -> str:
    override = os.environ.get(ENV_OVERRIDE)
    if override:
        # Explicit instruction: honour it, and let a bad path fail loudly here
        # rather than as a puzzling LibreOffice error twenty seconds later.
        if not _writable(override):
            raise OSError(
                f"{ENV_OVERRIDE}={override!r} is not writable. Point it at a "
                f"directory this process can create files in, or unset it to "
                f"fall back to {LOCAL_DIRNAME}/ in the working directory."
            )
        return use_temp_dir(override)

    inherited = _default_temp_dir()
    if inherited and _writable(inherited):
        return use_temp_dir(inherited)      # normal machine; no change in effect

    local = Path.cwd() / LOCAL_DIRNAME
    if not _writable(local):
        raise OSError(
            f"No writable temporary directory: {inherited or '/tmp'} is not "
            f"writable and neither is {local}. This skill needs somewhere to "
            f"put a LibreOffice profile and an intermediate PDF. Set "
            f"{ENV_OVERRIDE} to a writable path."
        )
    return use_temp_dir(local)


def _cleanup():
    """Remove our own temp directory if it is empty.

    Only ever removes an empty directory, and only the one we made: the
    individual ``TemporaryDirectory`` contexts clean themselves up, so what is
    left is either nothing -- which should go -- or somebody's output, which
    must not.
    """
    if TEMP_DIR and Path(TEMP_DIR).name == LOCAL_DIRNAME:
        try:
            os.rmdir(TEMP_DIR)
        except OSError:
            pass


TEMP_DIR = _resolve()
atexit.register(_cleanup)
