"""--require-verdicts: the visual half of QA turned into an exit code.

The gate passes only when every slide has an answer that gives back the review
code printed under its image -- proof that something with eyes saw that exact
render -- and no medium or high finding is open, or the user approved the
overrule.  The verdict table is a view of the answers and decides nothing.
"""

import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from PIL import Image  # noqa: E402

from make_deck import DECK_HASH_MARK, _gate_on_verdicts, deck_hash  # noqa: E402
from office import review  # noqa: E402

CODE = "ACDE-FHJK"


class GateOnVerdictsTest(unittest.TestCase):
    """The gate reads the report and never rewrites it.

    It used to rebuild first, which rewrote the report with blank verdicts and
    then found them blank: a gate nothing could pass.
    """

    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.deck = Path(d.name) / "deck.pptx"
        self.deck.write_bytes(b"not really a deck")
        self.report = Path(d.name) / "deck-report.md"
        self.render = Path(d.name) / "render"
        self.render.mkdir()
        for n, colour in ((1, "red"), (2, "green")):
            Image.new("RGB", (160, 90), colour).save(
                self.render / f"slide-{n}.png")
        with mock.patch.object(review, "new_code", return_value=CODE):
            review.prepare(self.render)
        self._write(deck_hash(self.deck))

    def _write(self, digest, verdicts=("", "")):
        rows = [f"| {n} |  | {v} |" for n, v in enumerate(verdicts, 1)]
        self.report.write_text("\n".join(
            ["# deck.pptx", f"{DECK_HASH_MARK}{digest} -->", "",
             "| Slide | Measured | Verdict |", "|---|---|---|", *rows]) + "\n")

    def _answer(self, number, severity=None, code=CODE):
        findings = ([{"q": "Q7", "what": "hollow", "severity": severity}]
                    if severity else [])
        (self.render / f"slide-{number}.answer.md").write_text(json.dumps(
            {"code": code, "verdict": "fix" if findings else "ok",
             "findings": findings}))

    def _allow(self, *entries):
        (self.deck.parent / "deck-allowlist.json").write_text(json.dumps(
            {"acknowledgements": list(entries)}))

    def test_seen_slides_pass_and_the_report_survives(self):
        self._answer(1)
        self._answer(2)
        before = self.report.read_text()
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)
        self.assertEqual(self.report.read_text(), before)

    def test_a_missing_answer_fails(self):
        self._answer(1)
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def test_an_answer_without_the_code_fails_whatever_the_cell_says(self):
        # What a build model with no eyes writes: plausible, and codeless.
        self._write(deck_hash(self.deck), ("ok", "ok"))
        self._answer(1)
        (self.render / "slide-2.answer.md").write_text(
            '{"verdict": "ok", "findings": []}')
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)
        self._answer(2, code="ACDE-FHJM")
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def test_report_for_another_build_fails(self):
        self._answer(1)
        self._answer(2)
        self._write("0" * 64)
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def test_a_medium_finding_holds_and_a_low_one_does_not(self):
        self._answer(1)
        self._answer(2, "medium")
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)
        self._answer(2, "low")
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)

    def test_an_overrule_counts_only_once_the_user_approved_it(self):
        self._answer(1)
        self._answer(2, "medium")
        entry = {"slide": 2, "kind": "vision", "content": "Q7",
                 "reason": "probe.py: text fills 70%"}
        self._allow(entry)
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)
        self._allow(dict(entry, approved_by="Maxi"))
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)

    def test_an_approved_person_review_stands_in_for_the_vision_agent(self):
        self._answer(1)
        self._allow({"slide": 2, "kind": "vision", "content": "seen",
                     "reason": "no vision-agent; looked at it myself",
                     "approved_by": "Maxi"})
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)

    def test_build_script_is_refused(self):
        self.assertEqual(_gate_on_verdicts(self.deck, Path("build.py")), 2)


if __name__ == "__main__":
    unittest.main()