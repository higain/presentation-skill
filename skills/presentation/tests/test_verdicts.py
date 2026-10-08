"""--require-verdicts: the visual half of QA turned into an exit code.

The report's verdict table is the record that something with eyes looked at
every slide.  This checks that a blank verdict cell makes the run fail -- a
slide nobody has opened cannot be finished, and the deck must not ship with it.
"""

import tempfile
import unittest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from make_deck import missing_verdicts


class MissingVerdictsTest(unittest.TestCase):
    def _report(self, lines):
        d = tempfile.TemporaryDirectory()
        self._tmp = d
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


if __name__ == "__main__":
    unittest.main()