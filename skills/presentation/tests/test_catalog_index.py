"""The catalog index and the build helpers behind it.

Pure logic over ``structure.json`` shapes -- no LibreOffice needed.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import catalog_import  # noqa: E402
from catalog_build import _canvas_leftovers, _headline  # noqa: E402
from office import draw as D  # noqa: E402
from office.template import _fold  # noqa: E402


def _slot(index, text="", placeholder=None, y=0, size=14.0, empty=False):
    return {"index": index, "sample_text": text, "placeholder": placeholder,
            "y": y, "fonts": [{"size": size}], "empty": empty}


class HeadlineTest(unittest.TestCase):
    def test_empty_title_placeholder_needs_set_title(self):
        s = {"slots": [_slot(0, placeholder="title", empty=True),
                       _slot(1, "Topic label", y=100)]}
        self.assertIn("set_title", _headline(s))

    def test_filled_title_placeholder_is_its_key(self):
        s = {"slots": [_slot(0, "Click to add title", placeholder="title")],
             "retext_keys": [{"address": 0, "key": "Click to add title"}]}
        self.assertIn('"Click to add title"', _headline(s))

    def test_largest_text_skips_a_section_number(self):
        s = {"slide_height": 9000,
             "slots": [_slot(1, "01", y=100, size=99),
                       _slot(2, "Getting started", y=4000, size=72),
                       _slot(3, "small print", y=8000, size=10)],
             "retext_keys": [{"address": 1, "key": "01"},
                             {"address": 2, "key": "Getting started"},
                             {"address": 3, "key": "small print"}]}
        self.assertIn('"Getting started"', _headline(s))


class CanvasLeftoverTest(unittest.TestCase):
    def test_kept_label_with_sample_text_is_reported(self):
        s = {"canvas": {"keep": [0, 1, 2]},
             "slots": [_slot(0, placeholder="title", empty=True),
                       _slot(1, "Double Diamond"), _slot(2, ""),
                       _slot(3, "removed anyway")]}
        self.assertEqual(_canvas_leftovers(s), [(1, "Double Diamond")])


class FoldTest(unittest.TestCase):
    def test_umlaut_less_key_matches(self):
        self.assertIn(_fold("Schlagworter"), _fold("Kennzahlen, Schlagwörter"))
        self.assertEqual(_fold("Verfügbarer"), _fold("verfugbarer"))


class CapacityParseTest(unittest.TestCase):
    def test_old_and_new_capacity_lines(self):
        old = catalog_import._capacity(
            "headline ✓ · body slots 3 · max bullets 4 · images 0 · chart ✗")
        new = catalog_import._capacity(
            "headline ✗ · body slots 0 · holds 4 item(s) · images 0 · chart ✗")
        self.assertEqual(old["max_bullets_per_body"], 4)
        self.assertEqual(new["max_items"], 4)
        self.assertIsNone(new["max_bullets_per_body"])


class VarietyTest(unittest.TestCase):
    def test_one_design_on_most_slides_is_a_finding(self):
        from office.qa.layout import variety_findings
        (text, key), = variety_findings([1, 37, 37, 37, 37, 31, 37], 62)
        self.assertEqual(key, ("variety", 2, "37"))
        self.assertIn("5 of 7", text)

    def test_varied_or_short_decks_are_fine(self):
        from office.qa.layout import variety_findings
        self.assertEqual(variety_findings([1, 4, 7, 4, 9, 2], 9), [])
        self.assertEqual(variety_findings([2, 2, 2], 9), [])


class CardTypeTest(unittest.TestCase):
    ITEMS = [{"title": "Vendor keys", "body": "Escalate by week 42."}]

    def test_short_cards_in_a_tall_panel_grow(self):
        small = D._grow_card_type(self.ITEMS, 9000, 2000, None, None,
                                  D.size_for("caption"), 400, False)
        big = D._grow_card_type(self.ITEMS, 9000, 12000, None, None,
                                D.size_for("caption"), 400, False)
        self.assertGreater(big[0], small[0])

    def test_a_size_the_caller_set_is_kept(self):
        size, _ = D._grow_card_type(self.ITEMS, 9000, 12000, 11.0, None,
                                    D.size_for("caption"), 400, False)
        self.assertEqual(size, 11.0)


if __name__ == "__main__":
    unittest.main()
