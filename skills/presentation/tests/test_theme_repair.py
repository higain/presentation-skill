#!/usr/bin/env python3
"""Checks for the theme-colour repair that runs after every ``build_deck``.

    python3 tests/test_theme_repair.py

No LibreOffice and no real template -- the parts a .pptx needs for this are
four small XML strings, so these run in well under a second.

Why this file exists: ``build_deck`` saves through LibreOffice, and that round
trip does not preserve the theme colour scheme.  On a 3-master template whose
themes all set ``dk2`` to the brand navy, the saved deck came back with six
masters, two of them invented by LibreOffice carrying its own default scheme
(``dk2=cccccc``).  A section divider paints its background as
``<a:schemeClr val="dk2"/>``, so it rendered light grey with white text on it
-- about 1.6:1 -- and *every automated check passed*, because QA reads shapes
and this defect lives in the theme behind them.  Only a human looking at a
render could see it.  That is precisely the kind of defect a test has to hold
down.
"""

import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office.deckbuild import (                                   # noqa: E402
    _scheme_colours, repair_theme_colours,
)

SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
         "accent4", "accent5", "accent6", "hlink", "folHlink")


def theme(dk2="022A3A", lt1_as_sysclr=False):
    """A theme part whose ``dk2`` is what the test is about."""
    colours = []
    for slot in SLOTS:
        if slot == "dk2":
            value = dk2
        elif slot == "lt1":
            value = "FFFFFF"
        else:
            value = "123456"
        if slot == "lt1" and lt1_as_sysclr:
            inner = f'<a:sysClr val="window" lastClr="{value}"/>'
        else:
            inner = f'<a:srgbClr val="{value}"/>'
        colours.append(f"<a:{slot}>{inner}</a:{slot}>")
    return (
        '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats'
        '.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">'
        + "".join(colours)
        + "</a:clrScheme><a:fontScheme name=\"x\"/></a:themeElements></a:theme>"
    )


def layout(name):
    return (f'<?xml version="1.0"?><p:sldLayout xmlns:p="p"><p:cSld '
            f'name="{name}"></p:cSld></p:sldLayout>')


def deck(path, masters):
    """Write a minimal .pptx.  ``masters`` is ``[(theme_xml, [layout names])]``."""
    # Layout parts are numbered across the whole file, not per master: the
    # part name has to match ``slideLayout<digits>.xml`` or the builder does
    # not recognise it as a layout at all -- which is how the first version of
    # this fixture silently tested the fallback path in every case.
    part_number = 0
    with zipfile.ZipFile(path, "w") as z:
        for i, (theme_xml, names) in enumerate(masters, start=1):
            z.writestr(f"ppt/slideMasters/slideMaster{i}.xml", "<p:sldMaster/>")
            z.writestr(f"ppt/theme/theme{i}.xml", theme_xml)
            rels = ['<?xml version="1.0"?><Relationships>',
                    f'<Relationship Target="../theme/theme{i}.xml"/>']
            for nm in names:
                part_number += 1
                part = f"slideLayout{part_number}.xml"
                z.writestr("ppt/slideLayouts/" + part, layout(nm))
                rels.append(f'<Relationship Target="../slideLayouts/{part}"/>')
            rels.append("</Relationships>")
            z.writestr(f"ppt/slideMasters/_rels/slideMaster{i}.xml.rels",
                       "".join(rels))
    return path


def dk2_of(path, index=1):
    with zipfile.ZipFile(path) as z:
        return _scheme_colours(z.read(f"ppt/theme/theme{index}.xml").decode())[2]


class ThemeRepair(unittest.TestCase):

    def setUp(self):
        self._tmp = TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_a_scheme_compares_by_colour_not_by_markup(self):
        """``sysClr`` and ``srgbClr`` for the same white are not a difference.

        The template writes ``lt1`` as ``<a:sysClr val="window"
        lastClr="FFFFFF"/>``; LibreOffice writes ``<a:srgbClr val="ffffff"/>``.
        Comparing the XML would call every theme damaged and rewrite all of
        them.
        """
        self.assertEqual(_scheme_colours(theme(lt1_as_sysclr=True)),
                         _scheme_colours(theme()))

    def test_the_default_libreoffice_wrote_is_replaced(self):
        src = deck(self.tmp / "t.pptx", [(theme("022A3A"), ["Divider"])])
        out = deck(self.tmp / "o.pptx", [(theme("cccccc"), ["Divider"])])
        changed = repair_theme_colours(src, out)
        self.assertEqual(len(changed), 1, changed)
        self.assertEqual(dk2_of(out), "022a3a")

    def test_a_deck_that_kept_its_colours_is_left_alone(self):
        """The single-master template that already round-trips cleanly."""
        src = deck(self.tmp / "t.pptx", [(theme("44546A"), ["Title"])])
        out = deck(self.tmp / "o.pptx", [(theme("44546A"), ["Title"])])
        self.assertEqual(repair_theme_colours(src, out), [])

    def test_a_master_libreoffice_invented_falls_back(self):
        """Layout names that match nothing still get the template's colours.

        LibreOffice's extra masters carry layouts the template never had.  As
        long as the template's own masters agree on a scheme, that is the one
        the deck should carry.
        """
        src = deck(self.tmp / "t.pptx",
                   [(theme("022A3A"), ["A"]), (theme("022A3A"), ["B"])])
        out = deck(self.tmp / "o.pptx", [(theme("cccccc"), ["Unrelated"])])
        changed = repair_theme_colours(src, out)
        self.assertEqual(len(changed), 1, changed)
        self.assertEqual(dk2_of(out), "022a3a")

    def test_a_tie_between_equal_schemes_is_not_a_conflict(self):
        """Two template masters sharing layout names and colours name one answer."""
        src = deck(self.tmp / "t.pptx",
                   [(theme("022A3A"), ["Shared"]), (theme("022A3A"), ["Shared"])])
        out = deck(self.tmp / "o.pptx", [(theme("cccccc"), ["Shared"])])
        self.assertEqual(len(repair_theme_colours(src, out)), 1)
        self.assertEqual(dk2_of(out), "022a3a")

    def test_masters_that_disagree_are_left_alone(self):
        """Fixing one brand by painting it in another brand's colours is worse.

        Two template masters, two different schemes, and an output master whose
        layout names match neither: there is no answer, so nothing is written.
        """
        src = deck(self.tmp / "t.pptx",
                   [(theme("022A3A"), ["A"]), (theme("1F497D"), ["B"])])
        out = deck(self.tmp / "o.pptx", [(theme("cccccc"), ["Unrelated"])])
        self.assertEqual(repair_theme_colours(src, out), [])
        self.assertEqual(dk2_of(out), "cccccc")

    def test_the_right_master_of_several_is_matched_by_layout_name(self):
        """With two brands in the file, layout names decide which is which."""
        src = deck(self.tmp / "t.pptx",
                   [(theme("022A3A"), ["Navy divider"]),
                    (theme("1F497D"), ["Blue divider"])])
        out = deck(self.tmp / "o.pptx", [(theme("cccccc"), ["Blue divider"])])
        self.assertEqual(len(repair_theme_colours(src, out)), 1)
        self.assertEqual(dk2_of(out), "1f497d")


if __name__ == "__main__":
    unittest.main(verbosity=2)
