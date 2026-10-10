"""The vision-agent's briefs and the reading of its answers.

The briefs were tuned against gemma-4-31b on planted defects; these tests pin
the parts that are code -- the JSON extraction, the verdict cells, and the
round trip from answer files into the report's verdict table.
"""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from make_deck import open_fix_verdicts  # noqa: E402
from office.vision import (CONTENT_TYPES, DEFAULT_LANGUAGE,  # noqa: E402
                           design_brief, last_json, slide_brief, verdict_line)
from vision_brief import collect  # noqa: E402


class LastJsonTest(unittest.TestCase):
    def test_outer_object_not_its_findings(self):
        text = ('STEP 1 ...\n{"verdict": "fix", "findings": '
                '[{"q": "Q5", "what": "banner covers cards"}]}')
        self.assertEqual(last_json(text)["verdict"], "fix")

    def test_last_of_several(self):
        self.assertEqual(last_json('{"a": 1} then {"b": 2}'), {"b": 2})

    def test_cannot_see_is_no_answer(self):
        # Even when the model then improvises a verdict from the brief's text.
        self.assertIsNone(last_json('CANNOT SEE IMAGE\n{"verdict": "ok"}'))

    def test_fenced_and_none(self):
        self.assertEqual(last_json('```json\n{"verdict": "ok"}\n```'),
                         {"verdict": "ok"})
        self.assertIsNone(last_json("no json here {broken"))


class VerdictLineTest(unittest.TestCase):
    def test_ok(self):
        self.assertTrue(verdict_line({"verdict": "ok", "findings": []})
                        .startswith("ok"))

    def test_fix_names_question_and_escapes_pipes(self):
        cell = verdict_line({"verdict": "fix", "findings": [
            {"q": "Q5", "what": "a | b", "fix": "move it"}]})
        self.assertIn("overlap: a / b → move it", cell)
        self.assertNotIn("|", cell)

    def test_unparsed_is_blank(self):
        self.assertEqual(verdict_line(None), "")


class BriefTest(unittest.TestCase):
    def test_language_defaults_to_the_slide(self):
        self.assertIn(DEFAULT_LANGUAGE, slide_brief("/x.png"))
        self.assertIn("Deck language: German", slide_brief("/x.png",
                                                           language="German"))

    def test_design_brief_carries_the_shared_vocabulary(self):
        brief = design_brief("/x.png", {"number": 3, "master": "M"})
        for kind in CONTENT_TYPES:
            self.assertIn(f'"{kind}"', brief)


class CollectTest(unittest.TestCase):
    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.dir = Path(d.name)
        self.deck = self.dir / "deck.pptx"
        self.render = self.dir / "render"
        self.render.mkdir()
        self.report = self.dir / "deck-report.md"
        # Exactly as make_deck writes the rows: one space in a blank cell.
        self.report.write_text(
            "# deck.pptx\n\n| Slide | Measured | Verdict |\n|---|---|---|\n"
            "| 1 | paints 9% of the slide | |\n| 2 |  | |\n"
            "| 3 |  | ok — by hand |\n| 4 |  | |\n")

    def test_round_trip(self):
        (self.render / "slide-1.answer.md").write_text(
            'Q1 HEADLINE: no\n{"verdict": "ok", "findings": []}')
        (self.render / "slide-2.answer.md").write_text(
            '{"verdict": "fix", "findings": [{"q": "Q7", "what": "hollow"}]}')
        (self.render / "slide-3.answer.md").write_text('{"verdict": "fix"}')
        (self.render / "slide-4.answer.md").write_text("I could not see it")
        filled, unanswered, unparsed = collect(self.deck, self.render)
        self.assertEqual((filled, unanswered, unparsed), (2, [], [4]))
        text = self.report.read_text()
        self.assertIn("| 1 | paints 9% of the slide | ok (vision-agent) |", text)
        self.assertIn("| 3 |  | ok — by hand |", text)   # a person wins
        self.assertEqual(open_fix_verdicts(self.report), [2])


if __name__ == "__main__":
    unittest.main()
