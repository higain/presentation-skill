"""--require-verdicts: the visual half of QA turned into an exit code.

The report's verdict table is the record that something with eyes looked at
every slide.  This checks that a blank verdict cell makes the run fail -- a
slide nobody has opened cannot be finished, and the deck must not ship with it.
"""

import json
import tempfile
import unittest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from make_deck import (DECK_HASH_MARK, _gate_on_verdicts, deck_hash,
                       missing_verdicts)


class MissingVerdictsTest(unittest.TestCase):
    def _report(self, lines):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        p = Path(d.name) / "deck-report.md"
        p.write_text("\n".join(lines) + "\n")
        return p

    def test_all_blank(self):
        p = self._report([
            "# deck",
            "| Slide | Measured | Verdict |",
            "|---|---|---|",
            "| 1 |  | |",
            "| 2 |  | |",
        ])
        self.assertEqual(missing_verdicts(p, 2), [1, 2])

    def test_mixed(self):
        p = self._report([
            "| Slide | Measured | Verdict |",
            "|---|---|---|",
            "| 1 |  | ok |",
            "| 2 |  | |",
            "| 3 |  | ok — cover |",
        ])
        self.assertEqual(missing_verdicts(p, 3), [2])

    def test_missing_report_counts_all_missing(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(missing_verdicts(Path(d) / "nope.md", 5),
                             [1, 2, 3, 4, 5])

    def test_no_table_rows(self):
        p = self._report(["# just a heading", ""])
        self.assertEqual(missing_verdicts(p, 2), [1, 2])


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

    def _write(self, digest, verdicts):
        rows = [f"| {n} |  | {v} |" for n, v in enumerate(verdicts, 1)]
        self.report.write_text("\n".join(
            ["# deck.pptx", f"{DECK_HASH_MARK}{digest} -->", "",
             "| Slide | Measured | Verdict |", "|---|---|---|", *rows]) + "\n")

    def test_filled_verdicts_pass_and_survive(self):
        self._write(deck_hash(self.deck), ["ok", "ok — cover"])
        before = self.report.read_text()
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)
        self.assertEqual(self.report.read_text(), before)

    def test_blank_verdict_fails(self):
        self._write(deck_hash(self.deck), ["ok", ""])
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def test_report_for_another_build_fails(self):
        self._write("0" * 64, ["ok", "ok"])
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def _answer(self, number, severity="medium"):
        render = self.deck.parent / "render"
        render.mkdir(exist_ok=True)
        (render / f"slide-{number}.answer.md").write_text(
            '{"verdict": "fix", "findings": [{"q": "Q7", "what": "hollow", '
            f'"severity": "{severity}"}}]}}')

    def test_editing_the_cell_does_not_clear_a_saved_finding(self):
        self._write(deck_hash(self.deck), ["ok", "ok — fills the area"])
        self._answer(2)
        self.assertEqual(_gate_on_verdicts(self.deck, None), 1)

    def test_low_severity_does_not_block(self):
        self._write(deck_hash(self.deck), ["ok", "ok"])
        self._answer(2, severity="low")
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)

    def test_an_acknowledged_finding_clears(self):
        self._write(deck_hash(self.deck), ["ok", "ok"])
        self._answer(2)
        (self.deck.parent / "deck-allowlist.json").write_text(json.dumps(
            {"acknowledgements": [{"slide": 2, "kind": "vision",
                                   "content": "Q7",
                                   "reason": "probe.py: text fills 70%"}]}))
        self.assertEqual(_gate_on_verdicts(self.deck, None), 0)

    def test_build_script_is_refused(self):
        self.assertEqual(_gate_on_verdicts(self.deck, Path("build.py")), 2)


if __name__ == "__main__":
    unittest.main()