"""Find the presentation skill and put its engine on the path.

This skill writes a template pack for the `presentation` skill and drives that
skill's own scripts to do it -- extraction, rendering, catalog assembly. It
deliberately duplicates none of them: a second copy of `catalog_extract.py`
would be a second copy to keep in step, and the two would disagree about what a
retext key is inside a month.

Looked for beside this skill first, then where skills are installed. If neither
has it, that is the whole prerequisite and the message says so.
"""

import sys
from pathlib import Path

# absolute(), not resolve(): installed skills are often symlinks, and a sandboxed
# agent may not reach the target that paths in its tracebacks would then name.
HERE = Path(__file__).absolute().parent.parent

#: Where the presentation skill might be, most specific first.
CANDIDATES = (
    HERE.parent / "presentation",                      # same skills/ directory
    HERE.resolve().parent / "presentation",            # beside the symlink's target
    Path.home() / ".claude" / "skills" / "presentation",
    Path("/usr/share/claude/skills/presentation"),
)


def presentation_skill() -> Path:
    """The presentation skill's root directory."""
    for candidate in CANDIDATES:
        if (candidate / "scripts" / "office").is_dir():
            return candidate
    raise SystemExit(
        "presentation-setup needs the `presentation` skill: it builds that "
        "skill's template packs and uses its engine to do it.\n"
        "  Looked in:\n    "
        + "\n    ".join(str(c) for c in CANDIDATES)
        + "\n  Install it, or run this from a checkout where both are siblings."
    )


def engine():
    """Put the presentation skill's ``scripts/`` on ``sys.path``; return root."""
    root = presentation_skill()
    scripts = str(root / "scripts")
    if scripts not in sys.path:
        sys.path.insert(0, scripts)
    return root


def templates_dir() -> Path:
    """Where packs are installed."""
    return presentation_skill() / "templates"
