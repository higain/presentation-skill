"""Template packs: everything this skill knows about one company's template.

A **pack** is a directory under ``templates/`` holding one template and the
facts measured or decided about it::

    templates/
      README.md            what a pack is -- ships with the skill
      acme/
        template.pptx      the asset itself; build_deck duplicates slides out of it
        profile.json       palette, type scale, authoring annotations
        catalog.md         which design carries which content (generated)
        overrides.md       the prose no script can measure (hand-written)
        catalog/           the extraction dossiers and the judged entries

The skill ships **no packs**.  A company makes its own by pointing the setup
skill at its template, and a company with two brands has two directories rather
than a fork.

Why a directory and not the old ``assets/`` + ``references/`` pair: the paths
used to be singular by construction -- one ``company-template.pptx``, one
``company-template.md`` -- so a second template had nowhere to live and the
palette a build drew with came from whichever one was installed, whatever
template it was actually building on.  Grouping by pack is what makes
"which template is this deck on" answerable at all.

It also makes the privacy boundary a path instead of a rule someone has to
remember: ``templates/*/`` is ignored by git, so a brand template cannot be
committed by forgetting a step.  ``templates/README.md`` is the exception, and
is the only thing in here a published skill contains.
"""

import json
from contextlib import contextmanager
from pathlib import Path

#: The skill root -- ``scripts/office/pack.py`` is three levels down.
SKILL = Path(__file__).resolve().parent.parent.parent
TEMPLATES = SKILL / "templates"

#: A pack's template file, whatever format the company works in.
TEMPLATE_EXTS = (".pptx", ".potx", ".odp", ".otp")

#: Names a pack directory always uses, so nothing has to be configured.
PROFILE_NAME = "profile.json"
CATALOG_NAME = "catalog.md"
OVERRIDES_NAME = "overrides.md"
CATALOG_DIR = "catalog"

#: ``{profile path: (mtime, parsed)}``.
_PROFILE_CACHE = {}

#: The pack a build is running against, innermost last.  See :func:`active`.
_ACTIVE = []

#: One line naming the pack to use when a deck names no template.  Absent is
#: the normal case: with exactly one pack installed, that one is the default,
#: and this file only exists to break a tie.
DEFAULT_MARKER = "default"


class Pack:
    """One template and its facts.  Attributes are paths; none need exist."""

    def __init__(self, root):
        self.root = Path(root)
        self.name = self.root.name

    @property
    def template(self):
        """The asset, or ``None`` if the directory holds no template file."""
        for ext in TEMPLATE_EXTS:
            candidate = self.root / f"template{ext}"
            if candidate.is_file():
                return candidate
        # A pack built by hand may have kept the template's own filename.
        for candidate in sorted(self.root.iterdir()) if self.root.is_dir() else ():
            if candidate.suffix.lower() in TEMPLATE_EXTS and candidate.is_file():
                return candidate
        return None

    @property
    def profile_path(self):
        return self.root / PROFILE_NAME

    @property
    def catalog(self):
        return self.root / CATALOG_NAME

    @property
    def overrides(self):
        return self.root / OVERRIDES_NAME

    @property
    def catalog_dir(self):
        return self.root / CATALOG_DIR

    def profile(self):
        """The parsed ``profile.json``, or ``{}``.

        Missing and malformed read the same on purpose: a pack half-built by an
        interrupted setup run should degrade to "no palette" rather than raise
        in the middle of somebody's build.  ``palette()`` says so when asked
        for a colour it does not have.

        Cached on path and mtime.  ``D.color()`` is called once per drawn
        shape, and a setup run rewriting the profile between two calls should
        be seen rather than cached over.
        """
        path = self.profile_path
        try:
            stamp = path.stat().st_mtime_ns
        except OSError:
            return {}
        key = str(path)
        hit = _PROFILE_CACHE.get(key)
        if hit is not None and hit[0] == stamp:
            return hit[1]
        try:
            parsed = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            parsed = {}
        _PROFILE_CACHE[key] = (stamp, parsed)
        return parsed

    def __repr__(self):
        return f"Pack({self.name!r})"

    def __eq__(self, other):
        return isinstance(other, Pack) and self.root == other.root

    def __hash__(self):
        return hash(self.root)


def packs():
    """Every installed pack, by name.  ``[]`` on a freshly published skill."""
    if not TEMPLATES.is_dir():
        return []
    found = [Pack(child) for child in sorted(TEMPLATES.iterdir())
             if child.is_dir() and not child.name.startswith(".")]
    return [pack for pack in found if pack.template is not None]


def default_pack():
    """The pack to build on when the user names no template, or ``None``.

    The ``templates/default`` marker wins; with no marker and exactly one pack
    installed, that one is the default -- which is the state nearly every
    company is in, and it means the common case needs no configuration at all.
    Two packs and no marker is deliberately **not** guessed: picking one of a
    company's two brands at random is worse than asking.
    """
    installed = packs()
    marker = TEMPLATES / DEFAULT_MARKER
    if marker.is_file():
        try:
            wanted = marker.read_text(encoding="utf-8").strip()
        except OSError:
            wanted = ""
        for pack in installed:
            if pack.name == wanted:
                return pack
    if len(installed) == 1:
        return installed[0]
    return None


def pack_named(name):
    """The installed pack called ``name``, or ``None``."""
    for pack in packs():
        if pack.name == name:
            return pack
    return None


def pack_for(template):
    """The pack a template file belongs to, or ``None``.

    Three cases, in order: the template is inside a pack directory; the
    template has a ``profile.json`` sitting beside it (a template the user
    brought, profiled in place, never installed); or it has neither and the
    caller gets ``None`` -- a real template with no facts known about it, which
    every drawing primitive has to survive.

    Resolved from the file being built on rather than from a global, because
    the colour a deck draws with is a fact about *its* template.  It used to be
    a module-level constant, and a build on a client's template came out in the
    house brand's cyan.
    """
    if template is None:
        return None
    path = Path(template).resolve()
    for pack in packs():
        found = pack.template
        if found is not None and found.resolve() == path:
            return pack
        try:
            path.relative_to(pack.root.resolve())
        except ValueError:
            continue
        return pack
    if (path.parent / PROFILE_NAME).is_file():
        return Pack(path.parent)
    return None


def describe_install():
    """One line for an error message: what is installed, and what to do."""
    installed = packs()
    if not installed:
        return ("No template pack is installed. Run the `presentation-setup` "
                "skill on your company template to make one, or pass a "
                "template path explicitly.")
    return ("Installed packs: %s. Name one with templates/%s, or pass a "
            "template path explicitly."
            % (", ".join(pack.name for pack in installed), DEFAULT_MARKER))


@contextmanager
def active(template):
    """Make ``template``'s pack the answer to "which template is this deck on".

    ``build_deck`` wraps the whole build in this, so every ``D.color()``,
    ``D.series()`` and ``role=`` inside a ``fill`` callback resolves against
    the template the slides are actually being duplicated out of.  Nested
    builds stack, and an exception unwinds the stack rather than leaving a
    pack active for the rest of the process -- which is how a global gets to
    be wrong quietly.

    A template with no pack pushes ``None``: that is a real state (a template
    the user brought, nothing measured about it), and it has to be the *deck's*
    state rather than falling back to whatever else is installed.  Drawing
    without a palette is handled; drawing in another company's palette is not.
    """
    _ACTIVE.append(pack_for(template))
    try:
        yield _ACTIVE[-1]
    finally:
        _ACTIVE.pop()


def active_pack():
    """The pack in force: the build's, else the default, else ``None``."""
    if _ACTIVE:
        return _ACTIVE[-1]
    return default_pack()


def active_profile(template=None):
    """The profile to read facts from.  ``{}`` when nothing is known.

    Explicit ``template`` wins over everything -- a caller that names a file is
    asking about that file.
    """
    if template is not None:
        pack = pack_for(template)
        return pack.profile() if pack is not None else {}
    pack = active_pack()
    return pack.profile() if pack is not None else {}
