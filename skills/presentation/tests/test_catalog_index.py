"""The catalog index and the build helpers behind it.

Pure logic over ``structure.json`` shapes -- no LibreOffice needed.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import catalog_import  # noqa: E402
from catalog_build import (_canvas_design, _canvas_leftovers,  # noqa: E402
                           _headline, _roles_block, _unnamed_designs)
from catalog_extract import _canvas_hint, _covered_by, _title_style  # noqa: E402
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
                       _slot(1, "Sample topic"), _slot(2, ""),
                       _slot(3, "removed anyway")]}
        self.assertEqual(_canvas_leftovers(s), [(1, "Sample topic")])


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


class HiddenTitleTest(unittest.TestCase):
    """A divider whose title placeholder lies under a full-slide picture."""

    def test_title_under_a_picture_is_never_set_title(self):
        title = _slot(1, placeholder="title", empty=True)
        title["covered_by"] = 4
        s = {"slide_height": 19050,
             "slots": [title, _slot(5, "Demo", y=14000, size=24)],
             "retext_keys": [{"address": 5, "key": "Demo"}]}
        line = _headline(s)
        self.assertIn('retext `"Demo"`', line)
        self.assertIn("hidden under picture `[4]`", line)
        self.assertNotIn("set_title(page", line)

    def test_a_filled_title_shape_beats_a_big_subtitle(self):
        # Probe reports ``placeholder`` only for an empty one; the cover's
        # title is filled, and its 32pt subtitle used to win on type size.
        title = dict(_slot(1, "Präsentationstitel", y=3200, size=24),
                     type="TitleTextShape")
        s = {"slots": [_slot(0, placeholder="subtitle", empty=True, size=32),
                       title],
             "title_budgets": [{"address": 0, "pt": 32.0, "width": 10851,
                                "chars_per_line": 21}],
             "retext_keys": [{"address": 1, "key": "Präsentationstitel"}]}
        self.assertEqual(_headline(s), 'retext `"Präsentationstitel"`')

    def test_extractor_marks_a_slot_under_a_later_picture(self):
        entries = [
            {"index": 1, "type": "TitleTextShape", "x": 1400, "y": 1300,
             "width": 31000, "height": 2200},
            {"index": 4, "type": "GraphicObjectShape", "x": 0, "y": 0,
             "width": 33866, "height": 19050},
            {"index": 5, "type": "TextShape", "x": 1400, "y": 14000,
             "width": 31000, "height": 1500},
        ]
        self.assertEqual(_covered_by(entries[0], entries), 4)
        self.assertIsNone(_covered_by(entries[2], entries))   # drawn above it
        inherited = dict(entries[0], index=None)
        self.assertEqual(_covered_by(inherited, entries), 4)


class FreeAreaTest(unittest.TestCase):
    W, H = 33866, 19050

    def test_a_full_bleed_photo_is_pulled_back_to_the_title_margins(self):
        photo = (0, 5616, self.W, self.H)
        title = (1464, 3203, 32432, 4661)
        number = (32367, 17970, 33865, 18767)
        x0, y0, x1, y1 = D.free_area([photo], [title, number], self.W, self.H)
        self.assertEqual((x0, x1), (1464, 32432))
        self.assertEqual(y0, 5616)
        self.assertLess(y1, 17970)

    def test_an_edge_to_edge_footer_does_not_set_the_margin(self):
        photo = (0, 5616, self.W, self.H)
        title = (1464, 3203, 32432, 4661)
        footer = (0, 18000, 10356, 18700)
        x0, _, _, _ = D.free_area([photo], [title, footer], self.W, self.H)
        self.assertEqual(x0, 1464)

    def test_a_body_placeholder_inside_the_margins_is_unchanged(self):
        body = (1433, 4026, 32432, 15124)
        title = (1433, 1323, 32432, 3522)
        self.assertEqual(D.free_area([body], [title], self.W, self.H), body)

    def test_the_extractor_counts_what_a_group_takes_with_it(self):
        entries = [
            {"index": 0, "type": "TitleTextShape", "x": 1400, "y": 1300,
             "width": 31000, "height": 2200},
            {"index": 3, "type": "GroupShape", "x": 5000, "y": 4000,
             "width": 26000, "height": 12000},
        ] + [{"index": f"3.{i}", "type": "TextShape", "x": 5000, "y": 4000,
              "width": 2000, "height": 1000} for i in range(1, 14)]
        hint = _canvas_hint(entries, self.W, self.H)
        self.assertEqual(hint["drops"], [3])
        self.assertEqual(hint["dropped_shapes"], 14)


def _design(master, keep, drops, box, dropped_shapes=None, slots=()):
    canvas = {"keep": keep, "drops": drops, "free_box": dict(
        zip("xywh", box))}
    if dropped_shapes is not None:
        canvas["dropped_shapes"] = dropped_shapes
    return {"master": master, "canvas": canvas, "slots": list(slots)}


class CanvasChoiceTest(unittest.TestCase):
    AREA = 33866 * 19050

    def test_a_cover_on_another_master_loses_to_a_content_slide(self):
        structures = {
            2: _design("Titelfolie", [0, 1, 2, 3, 4], [5],
                       (1464, 5616, 30968, 12164), 1),
            13: _design("Content", [0, 1, 2], [3], (1433, 4026, 30999, 11098), 1),
            14: _design("Content", [0, 1, 2], [3, 4], (1433, 4026, 30999, 9000), 2),
        }
        self.assertEqual(_canvas_design(structures, self.AREA), (13, True))

    def test_a_grouped_diagram_is_not_a_single_body(self):
        structures = {
            26: _design("Content", [0, 1, 2], [3], (5208, 3522, 27224, 13035), 14),
            14: _design("Content", [0, 1, 2], [3, 4, 5], (1433, 4026, 30999, 11098), 3),
            15: _design("Content", [0, 1, 2], [3], (1433, 4026, 30999, 11098), 30),
        }
        # Nothing qualifies; the plainest roomy design is the least-bad one.
        self.assertEqual(_canvas_design(structures, self.AREA), (14, False))

    def test_a_canvas_needs_room_and_a_visible_title(self):
        hidden = dict(_slot(1, placeholder="title", empty=True), covered_by=4)
        structures = {
            12: _design("Content", [0, 1, 2], [3, 4], (1433, 4026, 30999, 12000), 2,
                        slots=[hidden]),
            28: _design("Content", [0, 1, 2], [3, 4], (14658, 6620, 17641, 8747), 2),
            13: _design("Content", [0, 1, 2], [3, 4, 5, 6], (1433, 4026, 30999, 11098), 4),
        }
        self.assertEqual(_canvas_design(structures, self.AREA), (13, False))


    def test_a_divider_is_never_the_canvas(self):
        structures = {
            4: _design("Content", [0, 1, 2], [3], (1433, 4026, 30999, 11098), 1),
            6: _design("Content", [0, 1, 2], [3, 4], (1433, 4026, 30999, 9000), 2),
        }
        self.assertEqual(_canvas_design(structures, self.AREA, exclude={4}),
                         (6, True))
        # Excluding every candidate falls back rather than leaving no canvas.
        self.assertEqual(
            _canvas_design(structures, self.AREA, exclude={4, 6})[0], 4)


class RolesTest(unittest.TestCase):
    def test_an_opening_slide_row_names_the_cover(self):
        rules = [{"when": "Opening slide of the presentation", "use_slide": 1},
                 {"when": "Section start or divider", "use_slide": 12}]
        block = "\n".join(_roles_block(14, rules, {}))
        self.assertIn("| **cover** | `1` |", block)
        self.assertNotIn("There is no cover design", block)

    def test_a_question_row_is_not_a_cover(self):
        rules = [{"when": "Opening question to the audience", "use_slide": 29}]
        self.assertIn("There is no cover design",
                      "\n".join(_roles_block(14, rules, {})))

    def test_designs_no_row_names_are_listed(self):
        rules = [{"use_slide": 1}, {"use_slide": "`3`"}]
        entries = {1: {"name": "Cover"}, 3: {"name": "Agenda"},
                   14: {"name": "Canvas"}, 24: {"name": "Two columns"}, 25: {}}
        self.assertEqual(_unnamed_designs(rules, entries, 14),
                         [(24, "Two columns")])


class TitleStyleTest(unittest.TestCase):
    """The style filled titles share, for set_title to give an empty one."""

    @staticmethod
    def _title(text, color, size):
        return {"type": "TitleTextShape", "text": text,
                "paragraphs": [{"text": text, "color": color, "size": size}]}

    def test_the_commonest_filled_title_wins(self):
        slides = [{"shapes": [self._title("A", "#0066CC", 22.0)]},
                  {"shapes": [self._title("B", "#0066CC", 22.0)]},
                  {"shapes": [self._title("C", "#5A6367", 18.0)]},
                  {"shapes": [self._title("", "#5A6367", 18.0)]}]
        self.assertEqual(_title_style(slides),
                         {"color": "#0066CC", "size": 22.0, "seen": 2})

    def test_no_filled_title_means_no_style(self):
        self.assertIsNone(_title_style([{"shapes": [
            {"type": "OutlinerShape", "text": "Body", "paragraphs": []}]}]))

    def test_set_title_reads_it_from_the_pack(self):
        import json
        import tempfile
        from office.template import title_style
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            (root / "template.pptx").write_bytes(b"")
            (root / "profile.json").write_text("{}")
            (root / "catalog").mkdir()
            (root / "catalog" / "manifest.json").write_text(json.dumps(
                {"title_style": {"color": "#0066CC", "size": 22.0}}))
            self.assertEqual(title_style(root / "template.pptx"),
                             (0x0066CC, 22.0))
            (root / "catalog" / "manifest.json").write_text("{}")
            self.assertIsNone(title_style(root / "template.pptx"))


if __name__ == "__main__":
    unittest.main()
