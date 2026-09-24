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
    """Enough of a slide for the geometry: a 16:9 page in 1/100 mm."""
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


class BalancedLines(unittest.TestCase):
    """The wrap behind ``D.statement``.

    A greedy wrap fills each line to the brim and drops the remainder on the
    last one, which is how a statement slide shipped ending on lines reading
    ``it --``, ``an``, ``email.``
    """

    SENTENCE = ("We are deploying software that reliably follows instructions "
                "from anyone who can get text in front of it \u2014 including "
                "anyone who can send us an email.")

    def test_no_line_is_a_stub(self):
        lines = D.balance_lines(self.SENTENCE, 26000, 32.0, D.BOLD)
        widths = [D.text_width_est(line, 32.0, D.BOLD) for line in lines]
        # The point of balancing: the last line is not a fragment of the rest.
        self.assertGreater(min(widths), max(widths) * 0.5)

    def test_it_never_costs_a_line(self):
        for width in (12000, 20000, 26000, 30000):
            greedy = D.wrap_words(self.SENTENCE, width, 24.0, D.BOLD)
            balanced = D.balance_lines(self.SENTENCE, width, 24.0, D.BOLD)
            self.assertEqual(len(balanced), len(greedy), width)

    def test_no_line_is_wider_than_the_box(self):
        lines = D.balance_lines(self.SENTENCE, 26000, 32.0, D.BOLD)
        for line in lines:
            self.assertLessEqual(D.text_width_est(line, 32.0, D.BOLD), 26000)

    def test_a_short_string_is_left_alone(self):
        self.assertEqual(D.balance_lines("Two words", 26000, 32.0), ["Two words"])

    def test_a_statement_refuses_a_hand_broken_string(self):
        with self.assertRaises(ValueError) as caught:
            D.statement(None, None, D.Box(0, 0, 20000, 8000), "one\ntwo")
        self.assertIn("newline", str(caught.exception))

    def test_an_em_dash_is_a_full_em(self):
        # Counted as an ordinary letter it was the few percent that let a line
        # LibreOffice would re-break look like one that fits.
        self.assertGreater(D.advance_em("\u2014"), D.advance_em("n"))


class DesignVariety(unittest.TestCase):
    """The line that says a deck was built on one design.

    Four separate runs against the house template built ``[8] * N``, and
    nothing in the report said so.
    """

    def setUp(self):
        from qa import design_note
        self.note = design_note

    def test_nothing_recorded_says_nothing(self):
        self.assertIsNone(self.note([]))

    def test_one_design_is_called_out(self):
        note = self.note([8] * 15, catalog_size=21)
        self.assertIn("100%", note)
        self.assertIn("21", note)

    def test_a_varied_deck_is_reported_without_the_lecture(self):
        note = self.note([7, 8, 12, 13, 8, 1], catalog_size=21)
        self.assertIn("6 slide(s) on 5 design(s)", note)
        self.assertNotIn("%", note)

    def test_the_spread_leads_with_the_commonest(self):
        self.assertIn("3x design 8", self.note([8, 1, 8, 8, 1]))


class Wireframe(unittest.TestCase):
    """The ink map a reviewer with no eyes reads instead of the PNG."""

    def setUp(self):
        import wireframe
        self.wf = wireframe

    def grid(self, rows, cols, painted):
        return [[1.0 if (y, x) in painted else 0.0 for x in range(cols)]
                for y in range(rows)]

    def test_an_empty_slide_measures_empty(self):
        self.assertEqual(self.wf.measure(self.grid(10, 10, set()))["ink"], 0.0)

    def test_the_centroid_finds_a_block_off_to_one_side(self):
        painted = {(y, x) for y in range(4, 6) for x in range(0, 3)}
        stats = self.wf.measure(self.grid(10, 10, painted))
        self.assertLess(stats["cx"], 0.3)
        self.assertAlmostEqual(stats["cy"], 0.5, places=1)

    def test_the_blank_band_is_the_tallest_run(self):
        stats = self.wf.measure(self.grid(10, 10, {(0, 0), (9, 0)}))
        self.assertAlmostEqual(stats["band"], 0.8, places=2)
        self.assertAlmostEqual(stats["band_at"], 0.1, places=2)

    def test_margins_are_the_empty_gutters(self):
        margins = self.wf.measure(self.grid(10, 10, {(5, 5)}))["margins"]
        self.assertEqual(margins, (0.5, 0.4, 0.5, 0.4))

    def test_a_blank_cell_draws_as_a_space(self):
        drawn = self.wf.draw(self.grid(2, 2, {(0, 0)}))
        self.assertEqual(drawn[1], "  ")
        self.assertEqual(drawn[0][0], "@")


#: A pack belonging to no company. See its profile.json for why.
FIXTURE_PACK = Path(__file__).resolve().parent / "fixtures" / "pack"
#: The path a pack is resolved *from*. It need not exist: `pack_for` looks for
#: a profile beside the template, and a 2MB binary would prove nothing here.
FIXTURE_TEMPLATE = FIXTURE_PACK / "template.pptx"


class PackStyle(unittest.TestCase):
    """The palette and the type scale, read off whatever pack is in force.

    Four runs against a corporate template wrote their own hex -- 0xE8730C for
    orange, 0x6B3FA0 for purple -- and neither colour was anywhere in the deck
    they were matching. That is what the palette exists to prevent.

    Every assertion here is about *behaviour*, against an invented fixture:
    this class used to assert one real company's cyan, so it failed on any
    installation but the one it was written on -- including the published
    skill, which ships no packs at all.
    """

    def setUp(self):
        from office.pack import active
        self.ctx = active(FIXTURE_TEMPLATE)
        self.pack = self.ctx.__enter__()
        self.addCleanup(self.ctx.__exit__, None, None, None)

    def test_the_palette_is_the_pack_s_own(self):
        colours = D.palette()
        self.assertEqual(colours["signal"], 0x3355FF)
        self.assertNotIn(0xE8730C, colours.values())

    def test_roles_resolve_to_colours(self):
        self.assertEqual(D.color("accent"), D.color("signal"))
        self.assertEqual(D.color("alert"), D.palette()["warn"])

    def test_series_cycles_rather_than_running_out(self):
        seven = D.series(7)
        self.assertEqual(len(seven), 7)
        self.assertEqual(seven[0], seven[len(D.series())])

    def test_an_unknown_colour_names_the_real_ones(self):
        with self.assertRaises(ValueError) as caught:
            D.color("puce")
        self.assertIn("signal", str(caught.exception))

    def test_the_scale_descends_and_bottoms_out_at_the_floor(self):
        scale = D.type_scale()
        self.assertGreater(scale["title"], scale["footnote"])
        self.assertEqual(min(scale.values()), D.min_readable())

    def test_the_readable_floor_follows_the_template_not_a_constant(self):
        """A master whose smallest role is 10pt has a 10pt floor, not 9."""
        self.assertEqual(D.min_readable(), 10.0)
        self.assertNotEqual(D.min_readable(), D.MIN_READABLE_PT)

    def test_an_unknown_role_names_the_real_ones(self):
        with self.assertRaises(ValueError) as caught:
            D.size_for("tiny")
        self.assertIn("footnote", str(caught.exception))

    def test_contrast_ratio_reads_the_pack_s_colours(self):
        # Why a pack records the ratio: most brand colours are chosen against
        # paper, not against the 3:1 a *word* has to clear. `sun` is a fill and
        # can never be type; `ink` is type.
        self.assertLess(D.contrast_ratio(D.color("sun"), 0xFFFFFF), 3.0)
        self.assertGreater(D.contrast_ratio(D.color("ink"), 0xFFFFFF), 4.5)


class NoPackInstalled(unittest.TestCase):
    """A published skill ships no template. Everything here is that state.

    Not an edge case: it is what every company sees before they run the setup
    skill, and it is the state the suite never exercised -- which is how six
    tests came to fail on a fresh clone.
    """

    def setUp(self):
        from office.pack import active
        # A path with no profile anywhere beside it: a template the user
        # brought and nothing has measured.
        self.ctx = active(Path(__file__).resolve().parent / "nothing.pptx")
        self.pack = self.ctx.__enter__()
        self.addCleanup(self.ctx.__exit__, None, None, None)

    def test_there_is_no_pack(self):
        self.assertIsNone(self.pack)

    def test_the_palette_is_empty_rather_than_somebody_else_s(self):
        self.assertEqual(D.palette(), {})

    def test_asking_for_a_colour_raises_and_says_what_to_do(self):
        with self.assertRaises(ValueError) as caught:
            D.color("accent")
        message = str(caught.exception)
        self.assertIn("no pack", message)
        self.assertIn("presentation-setup", message)

    def test_the_type_scale_falls_back_so_a_build_still_runs(self):
        """Drawing without a palette is handled; drawing in the wrong one is not."""
        self.assertEqual(D.type_scale(), {})
        self.assertGreater(D.size_for("title"), D.size_for("footnote"))
        self.assertEqual(D.min_readable(), D.MIN_READABLE_PT)


class VennLabels(unittest.TestCase):
    """Labels belong to the lobe they name.

    They used to be pushed out along the ray from the figure's centre until
    they cleared the circle, which for a wide label on a diagonal is a long way
    -- so the two lower labels of a three-lobe Venn ended up nearer the slide's
    corners than the circles they name.
    """

    LABELS = ["PRIVATE DATA", "UNTRUSTED CONTENT", "EXTERNAL COMMUNICATION"]

    def labels(self, r=3000):
        return D._venn_labels(0, 0, r, 3, self.LABELS)

    def test_each_label_touches_its_own_lobe(self):
        r = 3000
        centres = D.venn_geometry(0, 0, r, 3)
        for (lx, ly, lw, lh, size, lines), (ex, ey) in zip(self.labels(r),
                                                           centres):
            ink_w = max(D.text_width_est(ln, size, D.BOLD) for ln in lines)
            ink_h = len(lines) * size * 35.28 * 1.22
            cx, cy = lx + lw / 2, ly + lh / 2
            # Nearest edge of the label to the lobe's centre: just outside it,
            # never a multiple of the radius away.
            dx = max(0.0, abs(cx - ex) - ink_w / 2)
            dy = max(0.0, abs(cy - ey) - ink_h / 2)
            gap = (dx * dx + dy * dy) ** 0.5
            self.assertGreaterEqual(gap, r * 0.9)
            self.assertLessEqual(gap, r * 1.35)

    def test_the_top_lobe_is_labelled_above_it_on_its_own_axis(self):
        (lx, ly, lw, lh, _, _) = self.labels()[0]
        ex, ey = D.venn_geometry(0, 0, 3000, 3)[0]
        self.assertAlmostEqual(lx + lw / 2, ex, delta=1)
        self.assertLess(ly + lh / 2, ey)

    def test_the_lower_lobes_are_labelled_beside_them(self):
        r = 3000
        centres = D.venn_geometry(0, 0, r, 3)
        left, right = self.labels(r)[1], self.labels(r)[2]
        self.assertLess(left[0] + left[2] / 2, centres[1][0])     # to its left
        self.assertGreater(right[0] + right[2] / 2, centres[2][0])  # to its right
        # On the lobe's own centre line, not dragged down the diagonal.
        self.assertAlmostEqual(left[1] + left[3] / 2, centres[1][1], delta=1)

    def test_a_labelled_figure_still_fits_the_box_it_was_given(self):
        box = D.Box(2031, 3545, 27972, 12157)
        cx, cy, r = D.venn_fit(box, 3, self.LABELS)
        bounds = D.venn_bounds(cx, cy, r, 3, self.LABELS)
        self.assertLessEqual(bounds.w, box.w + 1)
        self.assertLessEqual(bounds.h, box.h + 1)

    def test_a_centre_given_as_a_list_keeps_its_own_lines(self):
        size, lines = D._centre_fit(["DATA", "LEAK"], True, 2000)
        self.assertEqual(lines, ["DATA", "LEAK"])
        self.assertGreater(size, 0)


class ActionTitles(unittest.TestCase):
    """Whether a headline makes a point or names a topic.

    Measured on two real decks built from the same research: the one judged
    good read as an argument end to end, the one judged weak read as a table
    of contents -- and nothing in the report told them apart.
    """

    def setUp(self):
        from qa import _title_findings
        self.find = _title_findings

    def _reasons(self, text):
        return " ".join(self.find(3, "title", text))

    def test_a_so_what_sentence_is_clean(self):
        self.assertEqual(
            self.find(3, "title",
                      "Q3 revenue beat forecast by 15%, driven by APAC."), [])

    def test_german_counts_as_a_sentence_too(self):
        self.assertEqual(
            self.find(3, "title",
                      "Auf 42 geprüfte Bibliotheken kommen 1.447 "
                      "installierte."), [])

    def test_a_label_is_a_topic_not_a_point(self):
        self.assertIn("names a topic", self._reasons("Marktgröße"))
        self.assertIn("names a topic", self._reasons("Next Steps"))

    def test_a_question_is_not_a_title(self):
        self.assertIn("asks a question",
                      self._reasons("Warum ist das ein Problem?"))

    def test_the_missing_period_is_what_gives_a_label_away(self):
        self.assertIn("does not end in a period",
                      self._reasons("Growth trends over the last quarters"))

    def test_length_is_not_reported_per_slide(self):
        """A house style runs long everywhere; ``prose_findings`` says so once."""
        self.assertEqual(self.find(3, "title", " ".join(["word"] * 20) + "."),
                         [])


class BannedWords(unittest.TestCase):
    """The tells, matched literally so the list is the specification."""

    def setUp(self):
        from qa import BANNED, CLOSING_WASTE
        self.banned, self.closing = BANNED, CLOSING_WASTE

    def test_the_tells_are_caught_in_both_languages(self):
        for text in ("We leverage the pipeline", "best-in-class tooling",
                     "Eine ganzheitliche Lösung", "Synergien heben",
                     "die Mitarbeiter abholen", "seamlessly integrated"):
            self.assertIsNotNone(self.banned.search(text), text)

    def test_a_word_boundary_leaves_the_compound_alone(self):
        """``Mehrwertsteuer`` is VAT, not a buzzword."""
        self.assertIsNone(self.banned.search("Die Mehrwertsteuer beträgt 19%"))
        self.assertIsNone(self.banned.search("Empowerment Ltd is a customer"))

    def test_plain_writing_passes(self):
        self.assertIsNone(self.banned.search("Wir nutzen die Pipeline"))
        self.assertIsNone(
            self.banned.search("42 direkte, 1.447 installierte Pakete"))

    def test_a_wasted_last_slide_is_named(self):
        for text in ("Vielen Dank!", "Questions?", "Q & A", "Danke", "Fragen?"):
            self.assertIsNotNone(self.closing.match(text), text)

    def test_an_ask_is_not_a_wasted_last_slide(self):
        self.assertIsNone(
            self.closing.match("Drei Maßnahmen starten im Oktober."))


class PlanTitleTransfer(unittest.TestCase):
    """Whether the titles the user approved reached the deck that was built.

    Every string here is from the run that produced the defect: `plan.md`
    block 1 was approved as a sentence, and the build wrote the branch name
    "Current Blindspots" into the title slot on that slide and ten others.
    """

    #: Stage 3 writes exactly this shape -- three fields, one block per slide.
    PLAN = """## 0
title: A tiered control model balances security and agility.
proof: [TBD: ratio of direct vs transitive dependencies]
visual: statement

## 1
title: Our current dependency visibility is a major blindspot.
proof: [TBD: % of dependencies without known owners]
visual: area-compare
"""

    def setUp(self):
        import tempfile
        from qa import lost_plan_titles, plan_titles
        self.lost_plan_titles, self.plan_titles = lost_plan_titles, plan_titles
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        (self.dir / "plan.md").write_text(self.PLAN)
        self.deck = self.dir / "out.pptx"          # never opened, only located

    def _run(self, *slide_texts):
        return self.lost_plan_titles(
            self.deck, [(n, t) for n, t in enumerate(slide_texts, start=1)])

    def test_the_plan_is_read_in_order(self):
        self.assertEqual(
            self.plan_titles(self.deck),
            [("0", "A tiered control model balances security and agility."),
             ("1", "Our current dependency visibility is a major blindspot.")])

    def test_a_title_carried_over_verbatim_is_not_reported(self):
        lost, reworded = self._run(
            "A tiered control model balances security and agility.",
            "Our current dependency visibility is a major blindspot.")
        self.assertEqual((lost, reworded), ([], []))

    def test_the_branch_name_in_the_title_slot_loses_the_title(self):
        """The actual defect: the approved sentence is on no slide."""
        lost, _ = self._run(
            "A tiered control model balances security and agility.",
            "Current Blindspots", "Visibility")
        self.assertEqual(len(lost), 1)
        self.assertIn("block 1", lost[0])
        self.assertIn("on no slide", lost[0])

    def test_a_dropped_full_stop_is_not_a_lost_title(self):
        """Punctuation is a rule about writing, not about transfer."""
        lost, reworded = self._run(
            "A tiered control model balances security and agility",
            "Our current dependency visibility is a major blindspot")
        self.assertEqual((lost, reworded), ([], []))

    def test_a_tightened_headline_is_prose_and_never_lost(self):
        lost, reworded = self._run(
            "A tiered control model balances security and agility.",
            "Our dependency visibility is a blindspot.")
        self.assertEqual(lost, [])
        self.assertEqual(len(reworded), 1)
        self.assertIn("reworded", reworded[0])

    def test_a_title_that_landed_in_the_body_still_counts_as_arrived(self):
        """Placement is not measured -- see ``lost_plan_titles``."""
        lost, reworded = self._run(
            "Third-Party Dependency Risk Management A tiered control model "
            "balances security and agility. Working Session 2026",
            "Our current dependency visibility is a major blindspot.")
        self.assertEqual((lost, reworded), ([], []))

    def test_no_plan_beside_the_deck_means_no_findings(self):
        (self.dir / "plan.md").unlink()
        self.assertEqual(self._run("Anything at all"), ([], []))


class RepeatedTitles(unittest.TestCase):
    """The ghost deck reading as four labels instead of sixteen claims."""

    def setUp(self):
        from qa import repeated_titles
        self.repeated = repeated_titles

    def test_a_branch_name_used_as_four_titles_is_named_once(self):
        found = self.repeated([(5, "Tiered Control Logic"),
                               (6, "Tiered Control Logic"),
                               (7, "Tiered Control Logic"),
                               (8, "Something else entirely.")])
        self.assertEqual(len(found), 1)
        self.assertIn("slides 5, 6, 7", found[0])

    def test_distinct_titles_are_clean(self):
        self.assertEqual(
            self.repeated([(2, "Revenue beat forecast by 15%."),
                           (3, "APAC carried the quarter.")]), [])

    def test_dividers_without_a_readable_title_are_not_repeats(self):
        """Designs 1, 5 and 7 report ``None``; that is not four same titles."""
        self.assertEqual(self.repeated([(1, None), (4, None), (9, None)]), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
