"""
Headless LibreOffice process management.

Two ways in:

    run_soffice([...])   one-shot command line (``--convert-to`` and friends)
    office.impress       a live UNO bridge for reading and editing documents

Every invocation gets its own throwaway user profile, so nothing here collides
with a LibreOffice the user happens to have open, or with a second script
running at the same time.

**LibreOffice runs without a network.**  A deck can link an image by URL, and
LibreOffice fetches it on load -- a beacon that tells the deck's author who
opened it and when, or a request into the intranet.  Measured: one render of
such a deck sent six requests.  Every soffice started here runs inside its
own network namespace (``unshare``, else ``bwrap``), where there is nothing to
connect to; the UNO bridge is a named pipe, which is a file and reaches into
the namespace where a TCP port could not.  ``PRESENTATION_NETWORK=allow``
turns the isolation off; without either tool a warning says so once.
"""

import contextlib
import functools
import os
import secrets
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Iterable
from pathlib import Path

SOFFICE_MISSING = (
    "soffice not found on PATH; LibreOffice is required "
    "(apt install libreoffice-impress python3-uno)"
)


def soffice_path() -> str:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise FileNotFoundError(SOFFICE_MISSING)
    return exe


#: Network-less launchers, most preferred first.  Each keeps the caller's own
#: uid (the UNO pipe is named after it) and needs no root.
_ISOLATORS = (
    ["unshare", "--user", "--map-current-user", "--net"],
    ["bwrap", "--dev-bind", "/", "/", "--unshare-net", "--die-with-parent"],
)

NETWORK_ENV = "PRESENTATION_NETWORK"


@functools.lru_cache(maxsize=1)
def isolation_prefix() -> tuple:
    """The command prefix that runs a program with no network, or ``()``.

    Probed once per process by actually running ``true`` under it: a tool can
    be installed and still refused by the kernel (unprivileged user namespaces
    off, a container without the capability).
    """
    if os.environ.get(NETWORK_ENV, "").lower() == "allow":
        return ()
    for prefix in _ISOLATORS:
        if shutil.which(prefix[0]) is None:
            continue
        try:
            ok = subprocess.run(prefix + ["true"], capture_output=True,
                                timeout=10).returncode == 0
        except (OSError, subprocess.TimeoutExpired):
            ok = False
        if ok:
            return tuple(prefix)
    print("[presentation] WARNING: LibreOffice runs WITH network access -- "
          "neither `unshare` nor `bwrap` can create a network namespace "
          "here. An untrusted deck can make LibreOffice fetch URLs. Enable "
          "unprivileged user namespaces, or install bubblewrap.",
          file=sys.stderr)
    return ()


def isolated(argv) -> list:
    """``argv`` run without network, when this machine allows it."""
    return list(isolation_prefix()) + [str(a) for a in argv]


def pipe_name() -> str:
    """An unguessable name for a UNO pipe.

    The pipe is a socket file in ``/tmp`` named after it, so the name is the
    only thing between another local user and a bridge that can run code as
    this one.  A TCP port was worse: any local process could connect, and the
    port was free to be taken between choosing it and soffice binding it.
    """
    return f"lo-{secrets.token_hex(16)}"


#: A throwaway profile older than this belongs to a run that was killed
#: (SIGKILL, OOM) before it could clean up.
STALE_PROFILE_SECONDS = 12 * 3600


@functools.lru_cache(maxsize=1)
def sweep_stale_profiles() -> int:
    """Delete this user's abandoned ``lo-*profile-*`` directories.  Once per process.

    They hold LibreOffice's recent-documents list, which names the
    confidential template a build used.  ``atexit`` and the signal handlers
    remove them on every ordinary exit; a SIGKILL from an agent's timeout
    leaves them behind.
    """
    import time
    removed, cutoff = 0, time.time() - STALE_PROFILE_SECONDS
    root = Path(tempfile.gettempdir())
    for d in list(root.glob("lo-profile-*")) + list(root.glob("lo-uno-profile-*")):
        try:
            if d.owner() == Path.home().owner() and d.stat().st_mtime < cutoff:
                shutil.rmtree(d, ignore_errors=True)
                removed += 1
        except (OSError, KeyError):
            continue
    return removed


def get_soffice_env() -> dict:
    """Environment for a headless run: no X11, no Java probing noise."""
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    env.pop("DISPLAY", None)
    return env


@contextlib.contextmanager
def profile_dir():
    """A disposable ``-env:UserInstallation`` directory."""
    with tempfile.TemporaryDirectory(prefix="lo-profile-", ignore_cleanup_errors=True) as d:
        yield Path(d)


def profile_arg(path) -> str:
    return f"-env:UserInstallation={Path(path).absolute().as_uri()}"


def run_soffice(args: Iterable[str], timeout: int = 120, **kwargs) -> subprocess.CompletedProcess:
    """Run soffice once and wait for it.

    A private profile is added unless the caller supplied one.  Without it a
    non-root or concurrent run aborts with "User installation could not be
    completed" and converts nothing.
    """
    args = [str(a) for a in args]
    with contextlib.ExitStack() as stack:
        if not any(a.startswith("-env:UserInstallation") for a in args):
            args = [profile_arg(stack.enter_context(profile_dir()))] + args
        kwargs.setdefault("capture_output", True)
        kwargs.setdefault("text", True)
        return subprocess.run(
            isolated([soffice_path()] + args), env=get_soffice_env(),
            timeout=timeout, **kwargs
        )


if __name__ == "__main__":
    sys.exit(run_soffice(sys.argv[1:], capture_output=False).returncode)
