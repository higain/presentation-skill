"""Icons: lookup order, recolouring, the harvest, and one placed on a slide.

The first three need no LibreOffice.  The last builds a slide and reads the
saved ``.pptx``: the icon must arrive as an SVG (PowerPoint's own
``svgBlip``, with a PNG fallback) in the colour asked for.
"""

import importlib.util
import json
import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

from office import icons  # noqa: E402
from office.pack import Pack  # noqa: E402
from office.soffice import isolation_prefix  # noqa: E402

OFFICE_ICON = (
    '<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" '
    'id="Icons_MagnifyingGlass_LTR" overflow="hidden"><style>\n'
    '.MsftOfcThm_Accent1_Fill_v2 {\n fill:#0066CC; \n}\n</style>\n'
    '<path d="M1 1L9 9Z" class="MsftOfcThm_Accent1_Fill_v2" fill="#0066CC"/>'
    '</svg>')
BOOTSTRAP_ICON = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" '
    'fill="currentColor" class="bi bi-cup" viewBox="0 0 16 16">'
    '<path d="M1 1h4v4z"/><rect fill="none" stroke="#123456" '
    'width="2" height="2"/></svg>')
TWO_COLOURS = (
    '<svg viewBox="0 0 10 10" xmlns="http://www.w3.org/2000/svg">'
    '<path fill="#FF0000" d="M0 0h5v5z"/><path fill="#00FF00" d="M5 5h5v5z"/>'
    '</svg>')


def _harvest_module():
    path = SKILL.parent / "presentation-setup" / "scripts" / "harvest_icons.py"
    spec = importlib.util.spec_from_file_location("harvest_icons", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class RecolourTest(unittest.TestCase):
    def test_office_theme_colour_is_replaced(self):
        out = icons.recolour(OFFICE_ICON, 0xCC2B21)
        self.assertNotIn("<style", out)
        self.assertNotIn("class=", out)
        self.assertNotIn("#0066CC", out)
        self.assertIn('fill="#CC2B21"', out)

    def test_current_colour_and_strokes_take_it_none_stays(self):
        out = icons.recolour(BOOTSTRAP_ICON, 0x00AA00)
        self.assertNotIn("currentColor", out)
        self.assertIn('stroke="#00AA00"', out)
        self.assertIn('fill="none"', out)

    def test_root_gets_a_fill_when_paths_set_none(self):
        bare = '<svg viewBox="0 0 4 4"><path d="M0 0h4v4z"/></svg>'
        self.assertIn('<svg viewBox="0 0 4 4" fill="#112233">',
                      icons.recolour(bare, 0x112233))

    def test_aspect_reads_the_view_box(self):
        self.assertEqual(icons.aspect(OFFICE_ICON), 1.0)
        self.assertEqual(icons.aspect('<svg viewBox="0 0 20 10">'), 2.0)
        self.assertEqual(icons.aspect("<svg>"), 1.0)


class LookupTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.pack = Pack(self.tmp / "acme")
        folder = self.tmp / "acme" / "icons"
        folder.mkdir(parents=True)
        (folder / "cup.svg").write_text(OFFICE_ICON)
        (folder / "index.json").write_text(json.dumps({"icons": {
            "cup": {"file": "cup.svg", "keywords": ["coffee", "break"],
                    "mono": True}}}))
        self.extra = self.tmp / "set"
        self.extra.mkdir()
        (self.extra / "cup.svg").write_text(BOOTSTRAP_ICON)
        (self.extra / "shield-lock.svg").write_text(BOOTSTRAP_ICON)
        env = mock.patch.dict(os.environ, {icons.EXTRA_ENV: str(self.extra)})
        env.start()
        self.addCleanup(env.stop)
        system = mock.patch.object(icons, "SYSTEM_SETS", ())
        system.start()
        self.addCleanup(system.stop)
        bundled = mock.patch.object(icons, "BUNDLED_SETS", ())
        bundled.start()
        self.addCleanup(bundled.stop)

    def test_the_pack_wins_over_an_installed_set(self):
        label, path, entry = icons.find("cup", self.pack)
        self.assertEqual(label, "pack")
        self.assertEqual(path.parent.name, "icons")
        self.assertEqual(icons.find("shield-lock", self.pack)[0], "extra")

    def test_search_reads_keywords_and_lists_the_pack_first(self):
        self.assertEqual(icons.search("coffee", self.pack), [("cup", "pack")])
        self.assertEqual([n for n, _ in icons.search("", self.pack)][:1],
                         ["cup"])

    def test_unknown_name_suggests_close_ones(self):
        with self.assertRaises(ValueError) as caught:
            icons.find("shield", self.pack)
        self.assertIn("shield-lock", str(caught.exception))

    def test_no_installed_set_says_how_to_get_one(self):
        with mock.patch.dict(os.environ, {icons.EXTRA_ENV: ""}):
            with self.assertRaises(ValueError) as caught:
                icons.find("nothing-like-it", self.pack)
        self.assertIn("apt install bootstrap-icons", str(caught.exception))


class BundledSetTest(unittest.TestCase):
    """The Element Plus icons ship with the skill and must stay usable."""

    FOLDER = SKILL / "icons" / "ep"

    def test_the_set_and_its_licence_are_there(self):
        self.assertGreaterEqual(len(list(self.FOLDER.glob("*.svg"))), 290)
        self.assertIn("Element Plus", (self.FOLDER / "LICENSE").read_text())
        self.assertIn("Element Plus",
                      (SKILL / "THIRD-PARTY.md").read_text())

    def test_every_icon_is_inert_and_one_colour(self):
        for path in self.FOLDER.glob("*.svg"):
            svg = path.read_text()
            self.assertNotRegex(svg, r"(?i)<script|href|\son\w+=", path.name)
            self.assertLessEqual(len(icons.colours(svg)), 1, path.name)

    def test_a_set_prefix_reaches_an_icon_the_pack_shadows(self):
        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp)
        folder = tmp / "acme" / "icons"
        folder.mkdir(parents=True)
        (folder / "lock.svg").write_text(OFFICE_ICON)
        (folder / "index.json").write_text(
            json.dumps({"icons": {"lock": {"file": "lock.svg"}}}))
        pack = Pack(tmp / "acme")
        self.assertEqual(icons.find("lock", pack)[0], "pack")
        self.assertEqual(icons.find("ep:lock", pack)[0], "ep")
        self.assertIn(("lock", "ep"), icons.search("lock", pack))


class HarvestTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.h = _harvest_module()

    def test_office_id_names_the_icon_without_direction(self):
        self.assertEqual(self.h.name_for(OFFICE_ICON, "", "Grafik 3")[0],
                         "magnifying-glass")

    def test_alt_text_boilerplate_and_generic_names_are_not_names(self):
        bare = '<svg viewBox="0 0 1 1"/>'
        self.assertEqual(
            self.h.name_for(bare, "Lupe mit einfarbiger Füllung", "")[0],
            "lupe")
        self.assertIsNone(self.h.name_for(bare, "", "Inhaltsplatzhalter 76")[0])
        self.assertIsNone(self.h.name_for(bare, "", "Graphic 12")[0])

    def test_office_double_colour_counts_once(self):
        doubled = OFFICE_ICON.replace('fill="#0066CC"', 'fill="#0065CB"')
        self.assertTrue(self.h.is_mono(doubled))
        self.assertFalse(self.h.is_mono(TWO_COLOURS))

    def test_a_recoloured_copy_is_the_same_icon(self):
        self.assertEqual(self.h.fingerprint(OFFICE_ICON),
                         self.h.fingerprint(icons.recolour(OFFICE_ICON, 1)))

    def test_photos_and_wide_logos_are_not_icons(self):
        self.assertTrue(self.h.is_icon(OFFICE_ICON, 400))
        self.assertFalse(self.h.is_icon('<svg viewBox="0 0 300 60"/>', 100))
        self.assertFalse(self.h.is_icon(OFFICE_ICON.replace(
            "<path", "<image href='x.png'/><path"), 400))
        self.assertFalse(self.h.is_icon(OFFICE_ICON, 10 ** 6))


@unittest.skipUnless(shutil.which("soffice") and isolation_prefix(),
                     "needs LibreOffice and a network namespace")
class PlacedIconTest(unittest.TestCase):
    def test_icon_lands_as_svg_in_the_colour_asked(self):
        from office.deckbuild import build_deck
        from office import draw as D

        tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, tmp)
        (tmp / "set").mkdir()
        (tmp / "set" / "cup.svg").write_text(BOOTSTRAP_ICON)

        def fill(doc, page, slot, design):
            D.icon(doc, page, D.Box(2000, 2000, 3000, 3000), "cup",
                   color=0xAB1234)

        out = tmp / "deck.pptx"
        with mock.patch.dict(os.environ, {icons.EXTRA_ENV: str(tmp / "set")}):
            build_deck(str(SKILL / "tests" / "fixtures" /
                           "kite-brand-master.pptx"), out, [1], fill=fill,
                       qa=False)
        with zipfile.ZipFile(out) as z:
            slide = z.read("ppt/slides/slide1.xml").decode()
            svgs = [z.read(n).decode() for n in z.namelist()
                    if n.startswith("ppt/media/") and n.endswith(".svg")]
        self.assertIn("svgBlip", slide)
        self.assertIn('name="Icon cup"', slide)
        self.assertTrue(any("#AB1234" in s and "currentColor" not in s
                            for s in svgs), svgs)


if __name__ == "__main__":
    unittest.main()
