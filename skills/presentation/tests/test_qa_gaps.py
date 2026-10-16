"""The QA gaps a field run exposed: translucent panels, rival headlines,
text across panels.  Pure logic over stand-in shapes -- no LibreOffice."""

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office.qa import colour, layout, text  # noqa: E402


class BlendTest(unittest.TestCase):
    def test_a_translucent_panel_is_mostly_what_is_under_it(self):
        self.assertEqual(colour._blend(0x5CB800, 0.15, 0xFFFFFF) >> 16, 0xE7)
        self.assertEqual(colour._blend(0x0066CC, 1.0, 0xFFFFFF), 0x0066CC)

    def test_white_text_on_a_pale_panel_is_reported(self):
        from office.draw import contrast_ratio
        backdrop = colour._blend(0x5CB800, 0.15, 0xFFFFFF)
        self.assertLess(contrast_ratio(0xFFFFFF, backdrop), 1.5)


def shape(name, x, y, w, h, words="", sizes=()):
    return NS(Name=name, getString=lambda: words,
              getPosition=lambda: NS(X=x, Y=y),
              getSize=lambda: NS(Width=w, Height=h), sizes=sizes)


class HeadlineRivalTest(unittest.TestCase):
    def run_check(self, others):
        title = shape("Title", 1000, 2000, 30000, 2000, "A claim", (22.0,))
        page = NS(Height=19000)
        everything = [("1", title)] + [(str(i + 2), s) for i, s in enumerate(others)]
        with mock.patch.object(text, "placeholder", lambda p, k: title), \
                mock.patch.object(text, "walk_shapes", lambda p: everything), \
                mock.patch.object(text, "_text_sizes",
                                  lambda s: {z: "x" for z in s.sizes}):
            return text.headline_rival_issues(page, 4)

    def test_a_kicker_as_large_as_the_title_is_reported(self):
        found = self.run_check([shape("k", 1000, 500, 20000, 1200,
                                      "Einstieg", (32.0,))])
        self.assertEqual(len(found), 1)
        self.assertIn("above the title", found[0])

    def test_a_small_kicker_is_fine(self):
        self.assertEqual(self.run_check([shape("k", 1000, 500, 20000, 600,
                                               "Einstieg", (12.0,))]), [])

    def test_big_numbers_in_the_body_are_not_rivals(self):
        self.assertEqual(self.run_check([shape("kpi", 1000, 8000, 8000, 3000,
                                               "-35%", (44.0,))]), [])

    def test_a_big_label_in_the_footer_band_is_reported(self):
        found = self.run_check([shape("k", 1000, 17000, 9000, 1200,
                                      "Live Hack 1", (32.0,))])
        self.assertIn("footer band", found[0])


class StraddleTest(unittest.TestCase):
    def test_text_across_two_drawn_panels_is_reported(self):
        def panel(x):
            s = shape("D:rectangle", x, 6000, 7000, 9000)
            s.getShapeType = lambda: "com.sun.star.drawing.RectangleShape"
            s.FillStyle = "SOLID"
            return s
        line = shape("D:text", 1000, 9000, 20000, 600, "Quellen: Studie 2025")
        line.getShapeType = lambda: "com.sun.star.drawing.TextShape"
        items = [("1", panel(1000)), ("2", panel(9000)), ("3", line)]
        page = NS(Width=33000, Height=19000)
        with mock.patch.object(layout, "walk_shapes", lambda p: items), \
                mock.patch.object(layout, "_is_solid", lambda f: True), \
                mock.patch.object(layout, "ink_box",
                                  lambda d, p, s: (1000, 9000, 15000, 500,
                                                   s.getString())):
            found = layout.straddle_issues(None, page, 9)
        self.assertEqual(len(found), 1)
        self.assertIn("lies across 2 panels", found[0][0])
        self.assertEqual(found[0][1], ("straddle", 9, "3"))


if __name__ == "__main__":
    unittest.main()
