"""The allowlist: how a deck gets a clean report without hacking content.

Every case here is a failure a deck shipped.  The allowlist exists so a
correct-but-matching-template finding (an SDLC phase name, a divider's
background graphic) can be recorded as intentional instead of either hacked or
shipped with a permanently red report.  The rules it encodes:

- an acknowledgement counts only when it matches a real finding by
  kind + slide + content, and carries a non-empty reason;
- a stale acknowledgement -- one that matches nothing -- is itself reported;
- a malformed allowlist is an error, not a silent no-op;
- there is no wildcard, so an entry cannot accidentally cover a finding the
  author had not seen.
"""

import json
import tempfile
import unittest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import qa


def _deck(tmp, name="deck.pptx"):
    p = tmp / name
    p.write_bytes(b"")
    return p


class LoadAllowlistTest(unittest.TestCase):
    def _load(self, tmp, payload):
        p = _deck(tmp)
        qa._allowlist_path(p).write_text(json.dumps(payload))
        return qa.load_allowlist(p)

    def test_missing_file_is_no_allowlist(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(qa.load_allowlist(_deck(Path(d))), {})

    def test_empty_list_is_no_allowlist(self):
        with tempfile.TemporaryDirectory() as d:
            p = _deck(Path(d))
            qa._allowlist_path(p).write_text("[]")
            self.assertEqual(qa.load_allowlist(p), {})

    def test_accepts_acknowledgements_object(self):
        with tempfile.TemporaryDirectory() as d:
            out = self._load(Path(d), {"acknowledgements": [
                {"slide": 16, "kind": "leftover-wording", "content": "Testen",
                 "reason": "SDLC phase"}]})
            self.assertEqual(out["leftover-wording"][16]["Testen"], "SDLC phase")

    def test_accepts_plain_list(self):
        with tempfile.TemporaryDirectory() as d:
            out = self._load(Path(d), [
                {"slide": 5, "kind": "stray-image", "content": "image5.png",
                 "reason": "divider background"}])
            self.assertIn("divider background",
                          out["stray-image"][5]["image5.png"])

    def test_rejects_unknown_kind(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(qa.AllowlistError):
                self._load(Path(d), [{"slide": 1, "kind": "wat",
                                      "content": "x", "reason": "why"}])

    def test_rejects_bad_slide(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(qa.AllowlistError):
                self._load(Path(d), [{"slide": "a", "kind": "footer",
                                      "content": "x", "reason": "why"}])

    def test_rejects_missing_content(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(qa.AllowlistError):
                self._load(Path(d), [{"slide": 1, "kind": "footer",
                                      "reason": "why"}])

    def test_rejects_empty_reason(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(qa.AllowlistError):
                self._load(Path(d), [{"slide": 1, "kind": "footer",
                                      "content": "x", "reason": "  "}])

    def test_rejects_unreadable_json(self):
        with tempfile.TemporaryDirectory() as d:
            p = _deck(Path(d))
            qa._allowlist_path(p).write_text("{not json")
            with self.assertRaises(qa.AllowlistError):
                qa.load_allowlist(p)

    def test_rejects_non_list_object(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(qa.AllowlistError):
                self._load(Path(d), {"not": "acknowledgements"})


class AcknowledgeTest(unittest.TestCase):
    def test_exact_match_returns_reason(self):
        allow = {"leftover-wording": {16: {"Testen": "SDLC phase"}}}
        self.assertEqual(qa._acknowledged(allow, "leftover-wording", 16,
                                          "Testen"), "SDLC phase")

    def test_wrong_slide_does_not_match(self):
        allow = {"leftover-wording": {16: {"Testen": "SDLC phase"}}}
        self.assertIsNone(qa._acknowledged(allow, "leftover-wording", 17,
                                           "Testen"))

    def test_wrong_kind_does_not_match(self):
        allow = {"leftover-wording": {16: {"Testen": "SDLC phase"}}}
        self.assertIsNone(qa._acknowledged(allow, "footer", 16, "Testen"))

    def test_wrong_content_does_not_match(self):
        allow = {"leftover-wording": {16: {"Testen": "SDLC phase"}}}
        self.assertIsNone(qa._acknowledged(allow, "leftover-wording", 16,
                                           "Planen"))

    def test_no_wildcard(self):
        """An entry for one phrase must not cover a sibling on the same slide."""
        allow = {"leftover-wording": {16: {"Testen": "SDLC phase"}}}
        self.assertIsNone(qa._acknowledged(allow, "leftover-wording", 16,
                                           "Entwerfen"))


class StaleAcknowledgementsTest(unittest.TestCase):
    def test_stale_is_reported(self):
        allow = {"leftover-wording": {16: {"Testen": "x"}}}
        out = qa.stale_acknowledgements(allow, [("footer", 5, "y")])
        self.assertEqual(len(out), 1)
        self.assertIn("16", out[0])
        self.assertIn("Testen", out[0])

    def test_matched_is_not_stale(self):
        allow = {"leftover-wording": {16: {"Testen": "x"}}}
        out = qa.stale_acknowledgements(
            allow, [("leftover-wording", 16, "Testen")])
        self.assertEqual(out, [])

    def test_mixed(self):
        allow = {"footer": {5: {"a": "x"}}, "stray-image": {5: {"b.png": "y"}}}
        findings = [("footer", 5, "a")]
        out = qa.stale_acknowledgements(allow, findings)
        self.assertEqual(len(out), 1)
        self.assertIn("b.png", out[0])


class AllowlistPathTest(unittest.TestCase):
    def test_path_suffix(self):
        self.assertEqual(
            str(qa._allowlist_path(Path("deck.pptx"))),
            "deck-allowlist.json")

    def test_non_presentation_returns_none(self):
        self.assertIsNone(qa._allowlist_path(Path("notes.txt")))


class TopHeavyTest(unittest.TestCase):
    """The thresholds that turn a top-heavy content slide into a hard finding.

    Measured from the deck that shipped: the phase-canvas slides painted
    24-34 % with a 33-38 % empty band from 54-58 % down, and every one had
    been marked "sparse by design".  A genuine one-sentence statement paints
    under a quarter and must not trip this -- that stays advisory.
    """
    H = 10000

    def test_top_heavy_content_slide_fails(self):
        # A phase canvas: 30% painted, 38% empty band from 54% down.
        self.assertTrue(qa._is_top_heavy(0.30, 0.38 * self.H,
                                         0.54 * self.H, self.H))

    def test_low_area_statement_is_not_top_heavy(self):
        # A one-liner statement: 14% painted, band below it.
        self.assertFalse(qa._is_top_heavy(0.14, 0.71 * self.H,
                                          0.21 * self.H, self.H))

    def test_barely_below_area_threshold_is_not_top_heavy(self):
        self.assertFalse(qa._is_top_heavy(0.17, 0.40 * self.H,
                                          0.50 * self.H, self.H))

    def test_band_too_small_is_not_top_heavy(self):
        # Content fills most of the slide; only a footer gap below.
        self.assertFalse(qa._is_top_heavy(0.60, 0.25 * self.H,
                                          0.70 * self.H, self.H))

    def test_band_starting_in_top_third_is_not_top_heavy(self):
        # Closing slide: content top + branding bottom, band in between.
        self.assertFalse(qa._is_top_heavy(0.30, 0.50 * self.H,
                                          0.17 * self.H, self.H))

    def test_empty_slide_is_not_top_heavy(self):
        # Nothing painted is a different defect (sparse_slides catches it).
        self.assertFalse(qa._is_top_heavy(0.0, self.H, 0, self.H))


if __name__ == "__main__":
    unittest.main()