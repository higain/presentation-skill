#!/usr/bin/env python3
"""Fast checks for the arithmetic that decides whether a slide is right.

    python3 tests/test_geometry.py

No LibreOffice, no document, no network -- these are the pure functions under
``office.draw`` and the layout invariants a build script depends on.  They run
in well under a second, so there is no excuse for not running them.

Why these and not others: every case here is a defect somebody actually
shipped.  The ``cut_*`` ordering, the table widths that read like points, a
Venn whose labels leave the box -- each one rendered a slide wrong while every
automated check passed, because none of them is a *text* defect.
"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office import draw as D                                    # noqa: E402
from office.impress import mm, parse_slide_spec                 # noqa: E402


class FakePage:
    """Enough of a slide for the geometry: the delta template's size."""
    Width, Height = 32001, 18000


class BoxSlicing(unittest.TestCase):
    def setUp(self):
        self.box = D.Box(1000, 2000, 30000, 15000)

    def test_cut_returns_pair_in_reading_order(self):
        head, body = self.box.cut_top(mm(20))
        self.assertEqual(head.y, self.box.y)
        self.assertEqual(head.h, mm(20))
        self.assertEqual(body.y, self.box.y + mm(20))
        # The band is SECOND going the other way -- this is the pair everybody
        # unpacks backwards, and doing so silently fills the whole panel.
        body2, foot = self.box.cut_bottom(mm(20))
        self.assertEqual(foot.h, mm(20))
        self.assertEqual(foot.y2, self.box.y2)
        self.assertEqual(body2.y, self.box.y)

    def test_cut_pieces_cover_the_parent(self):
        for cut in ("cut_top", "cut_bottom"):
            a, b = getattr(self.box, cut)(mm(30))
            self.assertEqual(a.h + b.h, self.box.h)
        for cut in ("cut_left", "cut_right"):
            a, b = getattr(self.box, cut)(mm(30))
            self.assertEqual(a.w + b.w, self.box.w)

    def test_gap_is_taken_out_of_the_remainder(self):
        head, body = self.box.cut_top(mm(20), gap=mm(5))
        self.assertEqual(head.h, mm(20))
        self.assertEqual(body.y, self.box.y + mm(20) + mm(5))
        self.assertEqual(head.h + mm(5) + body.h, self.box.h)

    def test_named_halves_have_no_order_to_get_wrong(self):
        self.assertEqual(self.box.top(mm(20)).y, self.box.y)
        self.assertEqual(self.box.bottom(mm(20)).y2, self.box.y2)
        self.assertEqual(self.box.left(mm(60)).x, self.box.x)
        self.assertEqual(self.box.right(mm(60)).x2, self.box.x2)

    def test_cols_and_rows_fill_the_parent(self):
        cols = self.box.cols(4, gap=mm(8))
        self.assertEqual(len(cols), 4)
        self.assertEqual(cols[0].x, self.box.x)
        self.assertAlmostEqual(cols[-1].x2, self.box.x2, delta=1)
        widths = sum(c.w for c in cols) + 3 * mm(8)
        self.assertAlmostEqual(widths, self.box.w, delta=2)

        rows = self.box.rows(3, gap=mm(6))
        self.assertAlmostEqual(rows[-1].y2, self.box.y2, delta=1)

    def test_grid_is_rows_of_cols(self):
        grid = self.box.grid(2, 3, gap=mm(5))
        self.assertEqual(len(grid), 2)
        self.assertEqual([len(r) for r in grid], [3, 3])
        self.assertAlmostEqual(grid[-1][-1].x2, self.box.x2, delta=1)
        self.assertAlmostEqual(grid[-1][-1].y2, self.box.y2, delta=1)

    def test_inset_shrinks_on_every_side(self):
        inner = self.box.inset(mm(10))
        self.assertEqual(inner.x, self.box.x + mm(10))
        self.assertEqual(inner.w, self.box.w - 2 * mm(10))
        self.assertEqual(inner.centre, self.box.centre)

    def test_square_stays_inside(self):
        sq = self.box.square()
        self.assertEqual(sq.w, sq.h)
        self.assertLessEqual(sq.w, min(self.box.w, self.box.h))
        self.assertEqual(sq.centre, self.box.centre)


class BoxMisuse(unittest.TestCase):
    """The error messages are the feature: each one names the fix."""

    def test_a_box_in_a_coordinate_slot_is_refused(self):
        with self.assertRaises(TypeError) as e:
            D._num(D.Box(0, 0, 10, 10), "w", "rect")
        self.assertIn("Pass the box on its own", str(e.exception))

    def test_a_splatted_cut_pair_is_refused(self):
        pair = D.Box(0, 0, 100, 100).cut_top(10)
        with self.assertRaises(TypeError) as e:
            D._num(pair, "x", "rect")
        self.assertIn("TWO boxes", str(e.exception))

    def test_a_list_of_boxes_is_refused(self):
        cols = D.Box(0, 0, 100, 100).cols(3)
        with self.assertRaises(TypeError) as e:
            D._num(cols, "x", "rect")
        self.assertIn("one box per shape", str(e.exception))

    def test_a_misnamed_keyword_names_the_right_one(self):
        # Gemma wrote opacity=0.25 on an ellipse and font_size= on text.
        # Python's error says which word is wrong, never which is right.
        for fn, kw, wanted in (
                (D.ellipse, {"fill": 1, "opacity": 0.25}, "transparence"),
                (D.rect, {"fill": 1, "border": 2}, "line="),
                (D.text, {"font_size": 18}, "size="),
                (D.text, {"colour": 1}, "color="),
        ):
            with self.assertRaises(TypeError) as e:
                fn(None, FakePage(), D.Box(0, 0, 1000, 1000), "x", **kw)
            self.assertIn(wanted, str(e.exception))

    def test_a_box_plus_a_coordinate_keyword_is_refused(self):
        # Gemma wrote D.text(doc, page, body, "...", size=18, y=mm(60)).
        # Python's own error names an argument the caller never wrote
        # positionally; this one says what happened.
        with self.assertRaises(TypeError) as e:
            D.text(None, FakePage(), D.Box(0, 0, 1000, 1000), "x", y=mm(60))
        message = str(e.exception)
        self.assertIn("already supplies", message)
        self.assertIn("area.bottom", message)


class LiveRunRegressions(unittest.TestCase):
    """Misuses a model actually wrote while building a deck with this skill."""

    def test_a_box_is_not_subscriptable_but_says_what_to_write(self):
        # D.text(doc, page, D.canvas(page)[1], ...) -- expecting a cut_* pair.
        with self.assertRaises(TypeError) as e:
            D.Box(0, 0, 100, 100)[1]
        message = str(e.exception)
        self.assertIn("box.x", message)
        self.assertIn("ONE Box", message)

    def test_a_box_still_unpacks(self):
        self.assertEqual(tuple(D.Box(1, 2, 3, 4)), (1, 2, 3, 4))

    def test_canvas_with_nothing_to_keep_is_refused(self):
        # canvas(page) removes nothing and returns a guessed area, so the
        # template's own text stays on the slide under whatever you draw.
        with self.assertRaises(TypeError) as e:
            D.canvas(FakePage())
        self.assertIn("keep=(0, 1, 2, 3)", str(e.exception))

    def test_mm_refuses_a_value_already_in_hundredths(self):
        # D.Box(mm(500), mm(5000), mm(26000), mm(3000)) -- 26 metres wide.
        with self.assertRaises(ValueError) as e:
            mm(26000)
        message = str(e.exception)
        self.assertIn("metres", message)
        self.assertIn("ALREADY in 1/100 mm", message)

    def test_mm_still_converts_every_real_slide_dimension(self):
        self.assertEqual(mm(320), 32000)      # a slide's width
        self.assertEqual(mm(0.5), 50)         # a hairline
        self.assertEqual(mm(-20), -2000)      # an offset off the left edge

    def test_a_missing_box_method_names_the_real_one(self):
        # Gemma wrote box.move(...) and box.split_vertical(2), two builds
        # apiece.  Python's "'Box' object has no attribute 'move'" does not
        # hint that rows() and inset() are what it wanted.
        for name, wanted in (("move", "inset"),
                             ("split_vertical", "rows("),
                             ("split_horizontal", "cols("),
                             ("width", "no 'width'")):
            with self.assertRaises(AttributeError) as e:
                getattr(D.Box(0, 0, 100, 100), name)
            self.assertIn(wanted, str(e.exception))

    def test_an_unknown_box_method_still_lists_the_real_ones(self):
        with self.assertRaises(AttributeError) as e:
            D.Box(0, 0, 100, 100).nonsense
        message = str(e.exception)
        self.assertIn("box.cols(", message)
        self.assertIn("cut_top", message)

    def test_a_box_survives_copy_and_the_dunder_probes_it_makes(self):
        # __getattr__ must not answer __deepcopy__ and friends with layout
        # advice: copy, pickle and inspect all probe for absent dunders.
        import copy
        self.assertEqual(tuple(copy.deepcopy(D.Box(1, 2, 3, 4))), (1, 2, 3, 4))

    def test_a_template_function_asked_of_draw_names_its_module(self):
        # Gemma wrote D.find_shapes and D.shrink_to_fit -- both real, both in
        # office.template -- and spent four builds guessing at spellings.
        for name in ("find_shapes", "shrink_to_fit", "retext"):
            with self.assertRaises(AttributeError) as e:
                getattr(D, name)
            self.assertIn(f"from office.template import {name}", str(e.exception))

    def test_draw_points_at_impress_for_what_impress_defines(self):
        # office.template imports add_slide from office.impress, so a
        # first-hit-wins search sends the caller to the wrong module.
        with self.assertRaises(AttributeError) as e:
            D.add_slide
        self.assertIn("from office.impress import add_slide", str(e.exception))

    def test_a_name_in_neither_module_says_what_draw_is_for(self):
        with self.assertRaises(AttributeError) as e:
            D.flumox
        self.assertIn("office.template", str(e.exception))

    def test_a_nudge_keyword_is_told_to_move_the_box(self):
        # D.text(..., y_offset=mm(8)): there is no offset keyword anywhere,
        # because the geometry comes from the Box.
        with self.assertRaises(TypeError) as e:
            D.text(None, FakePage(), D.Box(0, 0, 1000, 1000), "x", y_offset=800)
        self.assertIn("inset", str(e.exception))


class TableWidths(unittest.TestCase):
    """The call that produced 54 QA findings in one run."""

    page = FakePage()

    def test_absolute_widths_that_look_like_points_are_refused(self):
        with self.assertRaises(ValueError) as e:
            D._column_widths([80, 60, 40, 120], None, 2031, self.page)
        message = str(e.exception)
        self.assertIn("cannot hold text", message)
        self.assertIn("proportions", message.lower())

    def test_proportions_scale_to_the_box_exactly(self):
        box = D.Box(2031, 3545, 27972, 12157)
        widths = D._column_widths([4, 3, 3, 8], box, None, self.page)
        self.assertAlmostEqual(sum(widths), box.w, delta=1)
        self.assertAlmostEqual(widths[0] / widths[1], 4 / 3, places=6)

    def test_fractions_summing_to_one_also_work(self):
        box = D.Box(0, 0, 10000, 5000)
        widths = D._column_widths([0.5, 0.25, 0.25], box, None, self.page)
        self.assertEqual([round(w) for w in widths], [5000, 2500, 2500])

    def test_absolute_mm_widths_are_accepted(self):
        widths = D._column_widths([mm(60), mm(40), mm(40)], None, mm(20),
                                  self.page)
        self.assertEqual(widths, [6000, 4000, 4000])

    def test_columns_past_the_slide_edge_are_refused(self):
        with self.assertRaises(ValueError) as e:
            D._column_widths([mm(150), mm(150), mm(80)], None, mm(20),
                             self.page)
        self.assertIn("past the right edge", str(e.exception))


class TextFitting(unittest.TestCase):
    def test_fit_size_shrinks_as_the_string_grows(self):
        short = D.fit_size("Phase I", mm(60))
        long = D.fit_size("Phase I: discovery and alignment", mm(60))
        self.assertGreater(short, long)
        self.assertGreaterEqual(long, 7.0)

    def test_fit_size_never_exceeds_its_cap(self):
        self.assertLessEqual(D.fit_size("x", mm(200), max_size=24.0), 24.0)

    def test_wrap_words_never_splits_a_word(self):
        text = "Untrusted content reaches the model"
        for line in D.wrap_words(text, mm(40), 14):
            for word in line.split():
                self.assertIn(word, text)
        self.assertEqual(" ".join(" ".join(D.wrap_words(text, mm(40), 14)).split()),
                         text)

    def test_fit_lines_returns_a_size_and_its_lines(self):
        size, lines = D.fit_lines("DATA EXFILTRATION", mm(30), mm(30))
        self.assertGreater(size, 0)
        self.assertTrue(lines)

    def test_est_lines_grows_with_length(self):
        self.assertLessEqual(D.est_lines("short", mm(80), 14),
                             D.est_lines("a great deal longer than that one",
                                         mm(80), 14))


class Contrast(unittest.TestCase):
    def test_known_ratios(self):
        self.assertAlmostEqual(D.contrast_ratio(0x000000, 0xFFFFFF), 21.0,
                               places=1)
        self.assertAlmostEqual(D.contrast_ratio(0xFFFFFF, 0xFFFFFF), 1.0,
                               places=1)

    def test_symmetric(self):
        self.assertAlmostEqual(D.contrast_ratio(0x009EE0, 0xFFFFFF),
                               D.contrast_ratio(0xFFFFFF, 0x009EE0), places=6)

    def test_readable_on_picks_the_visible_one(self):
        self.assertEqual(D.readable_on(0x1E2761), D.WHITE)   # navy
        self.assertEqual(D.readable_on(0xF2F2F2), D.BLACK)   # near-white
        # And what it picks always beats the threshold qa.py fails a build on.
        for bg in (0x1E2761, 0xF2F2F2, 0x009EE0, 0xE8730C):
            self.assertGreater(D.contrast_ratio(D.readable_on(bg), bg), 3.0)


class VennGeometry(unittest.TestCase):
    def test_three_circles_overlap(self):
        centres = D.venn_geometry(10000, 10000, 3000, 3)
        self.assertEqual(len(centres), 3)
        for i, a in enumerate(centres):
            for b in centres[i + 1:]:
                gap = ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5
                # Tangent or apart is the hand-rolled Venn everybody draws.
                self.assertLess(gap, 2 * 3000)
                self.assertGreater(gap, 0)

    def test_two_circles_overlap(self):
        (ax, ay), (bx, by) = D.venn_geometry(0, 0, 1000, 2)
        self.assertLess(abs(bx - ax), 2000)

    def test_venn_fit_keeps_the_labelled_figure_inside_its_box(self):
        box = D.Box(2031, 3545, 27972, 12157)
        labels = ["Private data", "Untrusted content", "External comms"]
        cx, cy, r = D.venn_fit(box, 3, labels)
        bounds = D.venn_bounds(cx, cy, r, 3, labels)
        self.assertLessEqual(bounds.w, box.w + 1)
        self.assertLessEqual(bounds.h, box.h + 1)
        self.assertGreater(r, 0)

    def test_venn_fit_uses_most_of_the_box(self):
        box = D.Box(0, 0, 20000, 12000)
        cx, cy, r = D.venn_fit(box, 3, None)
        bounds = D.venn_bounds(cx, cy, r, 3, None)
        self.assertGreater(bounds.h, box.h * 0.8)

    def test_pair_spots_sit_between_two_circles(self):
        spots = D.pair_spots(0, 0, 1000)
        self.assertEqual(len(spots), 3)
        for x, y in spots:
            self.assertLess((x * x + y * y) ** 0.5, 2000)


class SlideSpec(unittest.TestCase):
    def test_ranges_and_singles(self):
        self.assertEqual(parse_slide_spec("3,7-9"), [3, 7, 8, 9])
        self.assertEqual(parse_slide_spec(" 2 , 1 "), [1, 2])

    def test_rejects_nonsense(self):
        for bad in ("9-3", "0", "x", ""):
            with self.assertRaises(ValueError):
                parse_slide_spec(bad)


class PageBox(unittest.TestCase):
    def test_page_box_is_the_whole_slide(self):
        box = D.page_box(FakePage())
        self.assertEqual((box.x, box.y, box.w, box.h), (0, 0, 32001, 18000))


class VisionReviewer(unittest.TestCase):
    """The parts of review_renders.py that run without a model."""

    def setUp(self):
        import review_renders
        self.R = review_renders

    def test_contact_sheets_are_not_slides(self):
        # find_renders globs a render/ directory that also holds
        # contact-sheet.jpg / contact-sheet-1.jpg; only slide-NN.png is a slide.
        self.assertIsNone(self.R.slide_number(Path("contact-sheet.jpg")))
        self.assertEqual(self.R.slide_number(Path("slide-07.png")), 7)
        self.assertEqual(self.R.slide_number(Path("slide-12.png")), 12)

    def test_a_provider_qualified_id_implies_openrouter(self):
        # "google/gemma-4-31b-it" is an OpenRouter name; ollama calls the same
        # model "gemma3:27b" and would 404 on the first.
        url, _ = self.R.resolve_endpoint(model="google/gemma-4-31b-it")
        self.assertEqual(url, self.R.OPENROUTER_URL)
        url, _ = self.R.resolve_endpoint(model="gemma3:27b")
        self.assertIn("localhost", url)

    def test_an_explicit_base_url_wins(self):
        url, _ = self.R.resolve_endpoint(model="google/gemma-4-31b-it",
                                         base_url="http://127.0.0.1:8177/v1")
        self.assertEqual(url, "http://127.0.0.1:8177/v1")

    def test_json_survives_a_fenced_reply(self):
        got = self.R._extract_json('```json\n{"verdict": "ok"}\n```')
        self.assertEqual(got["verdict"], "ok")

    def test_prose_instead_of_json_is_an_error_not_a_finding(self):
        with self.assertRaises(ValueError):
            self.R._extract_json("Looks fine to me!")

    def test_findings_are_the_substance_and_the_verdict_is_derived(self):
        # A model that says "ok" and then lists a defect has contradicted
        # itself; the defect is what a reader can act on.
        rows = [(1, "ok", [{"what": "circles do not overlap", "where": "",
                            "severity": "high"}], None)]
        self.assertIn("circles do not overlap", self.R.as_markdown(rows, "m"))

    def test_a_clean_review_does_not_claim_a_guarantee(self):
        out = self.R.as_markdown([(1, "ok", [], None)], "gemma")
        self.assertIn("saw nothing wrong", out)
        self.assertIn("not a guarantee", out)

    def test_failed_slides_are_named_not_silently_dropped(self):
        out = self.R.as_markdown([(1, "ok", [], None),
                                  (2, None, [], "HTTP 500")], "gemma")
        self.assertIn("could not be reviewed", out)
        self.assertIn("slide 2", out)

    def test_one_defect_on_many_slides_is_one_line(self):
        # Verbatim from a gemma-4-31b-it run over test-gemma-v11: the
        # template's German master footer, reported four times in four
        # different wordings because every slide is a separate call.
        footer = [
            (2, "Placeholder text in footer ('Gegenstand | Titel | Ort | Datum')"),
            (9, "placeholder text in footer ('Gegenstand | Titel | Ort | Datum')"),
            (10, "Placeholder text in German/English ('Gegenstand | Titel | Ort | Datum')"),
            (6, "Footer text in German ('Gegenstand | Titel | Ort | Datum') "
                "while the rest of the slide is in English"),
        ]
        groups = self.R.cluster(
            [(n, {"what": w, "where": "footer", "severity": "high"})
             for n, w in footer])
        self.assertEqual(len(groups), 1, groups)
        self.assertEqual(groups[0][0], [2, 6, 9, 10])

    def test_where_joins_wordings_that_share_almost_nothing(self):
        # Verbatim from the second live run over the same deck.  The same
        # master footer, but this time worded two ways that share one content
        # word out of four -- "placeholder" -- and so printed as two lines
        # until the location joined the clustering key.
        footer = [
            (1, "placeholder text in a different language", "bottom left footer"),
            (10, "placeholder text in a different language", "bottom left footer"),
            (7, "placeholder text in footer", "bottom left"),
        ]
        groups = self.R.cluster(
            [(n, {"what": w, "where": p, "severity": "medium"})
             for n, w, p in footer])
        self.assertEqual(len(groups), 1, groups)
        self.assertEqual(groups[0][0], [1, 7, 10])

    def test_where_does_not_merge_different_defects_in_one_region(self):
        # Both live runs put several distinct findings at "center of the
        # slide"; sharing a location must not be enough on its own.
        rows = [
            (1, "large empty region that reads as unfinished",
             "center of the slide"),
            (3, "The Venn diagram circles are not aligned to a common center "
                "or balanced layout, leaving a large empty region on the left "
                "and right sides of the slide.", "center of the slide"),
            (6, "The table is not centered on the slide, leaving a large empty "
                "region on the right side.", "the table, center of slide"),
        ]
        groups = self.R.cluster(
            [(n, {"what": w, "where": p, "severity": "medium"})
             for n, w, p in rows])
        self.assertEqual(len(groups), 3, groups)

    def test_unrelated_findings_stay_apart(self):
        rows = [(1, "large empty region that reads as unfinished"),
                (5, "missing arrow between the second and third cards"),
                (3, "the Venn circles do not overlap")]
        groups = self.R.cluster(
            [(n, {"what": w, "where": "", "severity": "medium"})
             for n, w in rows])
        self.assertEqual(len(groups), 3)

    def test_a_cluster_leads_with_its_worst_severity(self):
        rows = [(1, {"what": "footer is in German", "where": "", "severity": "low"}),
                (2, {"what": "footer is in German still", "where": "",
                     "severity": "high"})]
        (numbers, _, worst), = self.R.cluster(rows)
        self.assertEqual(numbers, [1, 2])
        self.assertEqual(worst, self.R.RANK["high"])

    def test_worst_findings_come_first(self):
        rows = [(9, "look", [{"what": "low one", "where": "", "severity": "low"}], None),
                (2, "look", [{"what": "high one", "where": "", "severity": "high"}], None)]
        out = self.R.as_markdown(rows, "m")
        self.assertLess(out.index("high one"), out.index("low one"))


class EmptyBand(unittest.TestCase):
    """The sweep behind "this slide may be half empty".

    Both agent decks reviewed on 2026-08-05 shipped slides whose content
    stopped at the middle of the page, through a QA pass that reported clean.
    The arithmetic is here; the threshold it feeds is in ``qa.SPARSE_BAND``.
    """

    def setUp(self):
        from qa import _widest_empty_band
        self.band = _widest_empty_band

    def test_a_gap_between_two_blocks_is_found(self):
        widest, at = self.band([(0, 100), (600, 700)], 0, 1000)
        self.assertEqual((widest, at), (500, 100))

    def test_the_run_below_the_last_block_counts(self):
        widest, at = self.band([(0, 100), (100, 300)], 0, 1000)
        self.assertEqual((widest, at), (700, 300))

    def test_overlapping_blocks_do_not_invent_a_gap(self):
        # Bands arrive in z-order, overlapping freely -- a caption on a panel.
        widest, _ = self.band([(0, 500), (200, 400), (100, 900)], 0, 1000)
        self.assertEqual(widest, 100)

    def test_a_full_slide_has_no_band(self):
        self.assertEqual(self.band([(0, 1000)], 0, 1000)[0], 0)

    def test_blocks_outside_the_range_are_clipped_not_counted(self):
        widest, at = self.band([(-500, 0), (900, 1500)], 0, 1000)
        self.assertEqual((widest, at), (900, 0))


if __name__ == "__main__":
    unittest.main(verbosity=2)
