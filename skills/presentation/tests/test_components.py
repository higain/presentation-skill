"""Pure parts of the consulting components -- no LibreOffice needed."""

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office import draw as D  # noqa: E402


class TintTest(unittest.TestCase):
    def test_a_tint_lightens_toward_white(self):
        self.assertEqual(D.tint(0x0066CC, 1.0), 0x0066CC)
        self.assertEqual(D.tint(0x0066CC, 0.0), 0xFFFFFF)
        r, g, b = (D.tint(0x0066CC, 0.15) >> s & 255 for s in (16, 8, 0))
        self.assertTrue(r > 200 and g > 200 and b > 230)


class HeaderWidthTest(unittest.TestCase):
    def test_a_too_narrow_column_takes_room_from_a_wide_one(self):
        with mock.patch.object(D, "min_readable", lambda *a: 12.0):
            out = D._fit_header_widths([1000, 9000, 1000],
                                       ["#", "Maßnahme", "Verantwortlich"],
                                       12.0, 200)
        self.assertAlmostEqual(sum(out), 11000, delta=1)
        self.assertGreater(out[2], 1000)
        self.assertLess(out[1], 9000)

    def test_wide_enough_columns_are_left_alone(self):
        with mock.patch.object(D, "min_readable", lambda *a: 12.0):
            self.assertEqual(D._fit_header_widths([5000, 5000], ["A", "B"],
                                                  12.0, 200), [5000, 5000])


class CalibrationTest(unittest.TestCase):
    def test_the_factor_scales_every_estimate(self):
        with mock.patch.object(D, "_font", lambda f=None: "Test Sans"), \
                mock.patch.dict(D._FONT_FACTOR, {"Test Sans": 1.0}, clear=False):
            base = D.text_width_est("Verantwortlich", 12)
            D._FONT_FACTOR["Test Sans"] = 1.2
            self.assertAlmostEqual(D.text_width_est("Verantwortlich", 12),
                                   base * 1.2, places=3)
            D._FONT_FACTOR[("Test Sans", "bold")] = 1.3
            self.assertGreater(D.text_width_est("Verantwortlich", 12, D.BOLD),
                               D.text_width_est("Verantwortlich", 12) * 1.05)


class PairColumnsTest(unittest.TestCase):
    def test_two_equal_columns_with_a_gutter(self):
        left, right = D._pair_cols(D.Box(0, 0, 10000, 2000))
        self.assertEqual(left.w, right.w)
        self.assertEqual(right.x2, 10000)
        self.assertGreater(right.x, left.x2)


if __name__ == "__main__":
    unittest.main()
