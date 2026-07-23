"""
Headless LibreOffice process management.

Two ways in:

    run_soffice([...])   one-shot command line (``--convert-to`` and friends)
    office.calc          a live UNO bridge for reading and editing documents

Every invocation gets its own throwaway user profile, so nothing here collides
with a LibreOffice the user happens to have open, or with a second script
running at the same time.
"""

import contextlib
import os
import shutil
import socket
import subprocess
import tempfile
from collections.abc import Iterable
from pathlib import Path

SOFFICE_MISSING = (
    "soffice not found on PATH; LibreOffice is required "
    "(apt install libreoffice-calc python3-uno)"
)


def soffice_path() -> str:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise FileNotFoundError(SOFFICE_MISSING)
    return exe


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
            [soffice_path()] + args, env=get_soffice_env(), timeout=timeout, **kwargs
        )


def free_port() -> int:
    """An unused loopback port for the UNO bridge."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


if __name__ == "__main__":
    import sys

    sys.exit(run_soffice(sys.argv[1:], capture_output=False).returncode)
