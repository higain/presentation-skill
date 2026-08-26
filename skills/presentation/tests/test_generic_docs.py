#!/usr/bin/env python3
"""The documentation may not learn one company's template.

    python3 tests/test_generic_docs.py

A skill that ships no template still has to explain how to build a deck. The
rule that keeps that possible:

    **a generic file may name a role; only a pack may name a design number,
    a retext key or a colour.**

Nothing enforces that by itself, and nothing fails when it is broken -- the
docs still read fine to whoever wrote them, because they have the template in
front of them. It is invisible from inside, which is exactly why it needs a
test: three weeks of edits put 25 design numbers, 15 brand strings and 13
palette colours into files that are supposed to work for everybody, and one of
them arrived the day this rule was written down.

What the generated pack files say is their business. They are the pack.
"""

import re
import sys
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

#: Files that must work for any template.  `templates/README.md` is in here
#: because it ships; the packs beside it do not.
GENERIC_DOCS = sorted(
    [SKILL / "SKILL.md", SKILL / "README.md", SKILL / "templates" / "README.md"]
    + sorted((SKILL / "references").glob("*.md"))
)

#: A specimen of what the pipeline *writes* has to show real output, design
#: numbers and all.  It still may not name a real company.
SPECIMENS = {"catalog.example.md"}

#: Design numbers are facts about one template.  "design 8" in a generic file
#: is a sentence that is wrong for every reader who did not build it.
DESIGN_NUMBER = re.compile(r"\bdesigns?\s+\d+", re.I)

#: Wording, master names and retext keys from the template this skill was
#: written against.  A published skill must not contain any of it -- and the
#: *evidence* those runs produced is worth keeping, so what these forbid is the
#: identity, not the finding: "a 21-design corporate template" says the same
#: thing to a reader who has never seen it.
BRAND = re.compile(
    r"Potentielle Navigatoren|Navigator \(1/2\)|AT_ST|Section_dark"
    r"|Default 36|Gegenstand \| Titel|Kundenlogo"
    r"|\bdelta\b(?!\s*=)",
    re.I,
)

#: Colours a build may name in a generic file: the roles every profile defines.
#: `D.color("cyan")` is one company's palette; `D.color("accent")` is any.
ROLES = {"accent", "alert", "ok", "muted", "neutral"}
COLOR_CALL = re.compile(r'D\.color\(\s*["\']([a-z_]+)["\']\s*\)')


def _hits(path, pattern):
    """``[(line number, the matched text, the whole line)]``."""
    out = []
    for number, line in enumerate(
            path.read_text(encoding="utf-8").splitlines(), start=1):
        for found in pattern.finditer(line):
            out.append((number, found.group(0), line.strip()))
    return out


def _report(name, found):
    return "\n".join(
        f"  {name}:{number}: {match!r}\n      {line[:100]}"
        for number, match, line in found)


class GenericDocs(unittest.TestCase):

    def test_no_design_numbers_outside_a_pack(self):
        for path in GENERIC_DOCS:
            if path.name in SPECIMENS:
                continue
            with self.subTest(doc=path.name):
                found = _hits(path, DESIGN_NUMBER)
                self.assertEqual(
                    found, [],
                    f"\n{path.name} names a design number. Design numbers are "
                    f"facts about one template: state the *role* (the canvas "
                    f"design, a divider design) and let the pack's catalog say "
                    f"which number that is.\n" + _report(path.name, found))

    def test_no_brand_wording_anywhere_generic(self):
        for path in GENERIC_DOCS:
            with self.subTest(doc=path.name):
                found = _hits(path, BRAND)
                self.assertEqual(
                    found, [],
                    f"\n{path.name} names the template this skill was written "
                    f"against. Keep the evidence, drop the identity.\n"
                    + _report(path.name, found))

    def test_colours_are_named_by_role_not_by_palette(self):
        for path in GENERIC_DOCS:
            if path.name in SPECIMENS:
                continue
            with self.subTest(doc=path.name):
                found = [(number, match, line)
                         for number, match, line in _hits(path, COLOR_CALL)
                         if COLOR_CALL.match(match).group(1) not in ROLES]
                self.assertEqual(
                    found, [],
                    f"\n{path.name} names a colour out of one template's "
                    f"palette. Generic docs use roles: {', '.join(sorted(ROLES))}."
                    f"\n" + _report(path.name, found))


class EngineCode(unittest.TestCase):
    """The scripts may cite what a run did; they may not name whose template.

    Design numbers survive in code comments on purpose -- "four runs built
    [8] * N" is the evidence for a check, and rewriting it into the abstract
    would lose the thing that makes it believable. The company's name is not
    part of that evidence.
    """

    def test_no_brand_wording_in_the_engine(self):
        for path in sorted((SKILL / "scripts").rglob("*.py")):
            if "__pycache__" in path.parts:
                continue
            with self.subTest(module=path.name):
                found = _hits(path, BRAND)
                self.assertEqual(
                    found, [],
                    f"\n{path.name} names the template this skill was written "
                    f"against.\n" + _report(path.name, found))


if __name__ == "__main__":
    unittest.main(verbosity=2)
