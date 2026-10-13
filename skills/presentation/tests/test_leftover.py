"""Leftover wording is judged against the slide's own design.

A word a build drew itself -- a phase bar's "Testen" on a canvas slide -- is
content, even when some other design of the template carries the same word.
One run reported 112 such findings and allowlisted every one.
"""

import sys
import tempfile
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

from office.qa.baseline import _Baseline, _baseline_load, _baseline_save  # noqa: E402


class LeftoverStringsTest(unittest.TestCase):
    def setUp(self):
        self.base = _Baseline(
            strings={"Bereitstellen", "Thema (1 von 2)"},
            by_design={14: {"Thema (1 von 2)"}, 18: {"Bereitstellen"}})

    def test_another_designs_word_is_not_a_leftover(self):
        self.assertNotIn("Bereitstellen", self.base.leftover_strings(14))

    def test_the_slides_own_design_label_still_is(self):
        self.assertIn("Thema (1 von 2)", self.base.leftover_strings(14))
        self.assertIn("Bereitstellen", self.base.leftover_strings(18))

    def test_no_recorded_design_falls_back_to_the_whole_template(self):
        self.assertEqual(self.base.leftover_strings(None), self.base.strings)
        self.assertEqual(self.base.leftover_strings(99), self.base.strings)

    def test_the_per_design_sets_survive_the_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "baseline.json"
            _baseline_save(cache, self.base)
            back = _baseline_load(cache)
        self.assertEqual(back.by_design, self.base.by_design)

    def test_a_cache_without_them_is_a_miss(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "baseline.json"
            cache.write_text('{"strings": [], "bleeds": [], "contrast": [], '
                             '"occlusion": [], "collisions": []}')
            self.assertIsNone(_baseline_load(cache))


if __name__ == "__main__":
    unittest.main()
