"""Hand-placed slides: a reason to look, and near-misses that fail.

Pure logic over stand-in shapes -- no LibreOffice needed.  That the names
reach the saved file is checked by building a deck, not here.
"""

import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office.qa import layout  # noqa: E402
from office.qa.allowlist import load_allowlist  # noqa: E402

MM = 100


def page(*names):
    return [(f"{i}", SimpleNamespace(Name=name)) for i, name in enumerate(names)]


def note(*names):
    with mock.patch.object(layout, "walk_shapes", lambda p: page(*names)):
        return layout.hand_placed_note(None)


def box(address, x, y, w=60 * MM, h=12 * MM, kind="rectangle"):
    return (address, kind, x, y, w, h)


class HandPlacedNoteTest(unittest.TestCase):
    def test_many_primitives_and_no_helper_is_a_slide_to_look_at(self):
        self.assertIn("14 shapes placed one by one",
                      note(*(["D:rectangle", "D:text"] * 7)))

    def test_one_helper_on_the_slide_clears_it(self):
        self.assertIsNone(note(*(["D:text"] * 20 + ["D:cards"])))

    def test_at_the_limit_is_fine(self):
        self.assertIsNone(note(*(["D:text"] * layout.HAND_PLACED)))

    def test_template_shapes_and_icons_are_never_counted(self):
        self.assertIsNone(note(*(["PlaceHolder 1", "Grafik 3", "",
                                  "Icon shield"] * 10)))


class MisalignmentTest(unittest.TestCase):
    def test_a_column_laid_out_exactly_passes(self):
        rows = [box(f"{i}", 20 * MM, 40 * MM + i * 14 * MM) for i in range(6)]
        self.assertEqual(layout.misalignments(rows), [])

    def test_a_row_shifted_by_a_millimetre_and_a_half_fails(self):
        rows = [box("0", 20 * MM, 40 * MM), box("1", 20 * MM, 54 * MM),
                box("2", 21.5 * MM, 68 * MM)]
        found = layout.misalignments(rows)
        self.assertEqual(len(found), 1, found)
        self.assertIn("3 shapes miss a peer's edge by 1.5 mm", found[0])

    def test_an_offset_over_three_millimetres_is_a_choice(self):
        rows = [box("0", 20 * MM, 40 * MM), box("1", 26 * MM, 54 * MM)]
        self.assertEqual(layout.misalignments(rows), [])

    def test_a_label_inset_in_its_box_is_not_a_miss(self):
        cell = box("0", 20 * MM, 40 * MM)
        label = box("1", 22 * MM, 40 * MM, w=56 * MM, kind="text")
        same_kind_inset = box("2", 22 * MM, 41 * MM, w=56 * MM, h=10 * MM)
        self.assertEqual(layout.misalignments([cell, label, same_kind_inset]), [])

    def test_shapes_of_other_sizes_are_not_peers(self):
        big = box("0", 20 * MM, 40 * MM)
        small = box("1", 21.5 * MM, 60 * MM, w=20 * MM)
        self.assertEqual(layout.misalignments([big, small]), [])

    def test_a_row_with_almost_even_gaps_fails(self):
        cards = [box(f"{i}", x * MM, 40 * MM, w=50 * MM)
                 for i, x in enumerate((20, 74, 129.5))]  # gaps 4.0 and 5.5
        found = layout.misalignments(cards)
        self.assertTrue(any("spaced 4.0, 5.5 mm" in f for f in found), found)

    def test_a_row_with_even_gaps_passes(self):
        cards = [box(f"{i}", x * MM, 40 * MM, w=50 * MM)
                 for i, x in enumerate((20, 74, 128))]
        self.assertEqual(layout.misalignments(cards), [])

    def test_the_finding_names_the_slide_and_can_be_acknowledged(self):
        rows = [box("0", 20 * MM, 40 * MM), box("1", 21.5 * MM, 54 * MM)]
        with mock.patch.object(layout, "_drawn", lambda p: rows):
            (text, key), = layout.alignment_issues(None, 7)
        self.assertTrue(text.startswith("slide 7: drawn shapes almost line up"))
        self.assertEqual(key, ("alignment", 7, "near-miss"))
        with tempfile.TemporaryDirectory() as tmp:
            deck = Path(tmp) / "d.pptx"
            deck.write_bytes(b"")
            (Path(tmp) / "d-allowlist.json").write_text(json.dumps(
                {"acknowledgements": [
                    {"slide": 7, "kind": "alignment", "content": "near-miss",
                     "reason": "a deliberately staggered row"}]}))
            self.assertIn("alignment", load_allowlist(deck))


if __name__ == "__main__":
    unittest.main()
