"""Source fidelity: nothing dropped, changed or invented.  Pure text."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office.qa.fidelity import (fidelity_findings, invented,  # noqa: E402
                                missing_units, source_units)

SOURCE = """\
# Auftrag

Texte wörtlich übernehmen, „…“-Anführungszeichen inklusive.

## 4 · Ein Filter, der 95 % der Angriffe erkennt, ist eine Durchfallquote

- **Struktur:** Gegenüberstellung zweier Ansätze („Erkennung allein“ gegen Architektur).
- **Links – Erkennung** (Icon: Filter-Trichter):
  - Guardrail-Klassifikatoren und Prompt-Filter
  - **Kennzahl:** „> 90 %“, dazu: „der 12 publizierten Abwehrmaßnahmen wurden
    adaptiv umgangen“
- **Zitat:** „Hat ein Agent untrusted Input aufgenommen, muss er so
  eingeschränkt sein, dass dieser Input keine folgenreiche Aktion auslösen
  kann.“

| Schritt | Icon | Überschrift |
|---|---|---|
| 1 | Liste | Tool-Inventar erstellen |

**Sprechernotizen:**

> Guardrails sind nicht nutzlos – als zusätzliche Schicht über Least Privilege.
"""

FAITHFUL = [(1, ["Ein Filter, der 95 % der Angriffe erkennt, ist eine Durchfallquote",
                 "Erkennung allein", "Guardrail-Klassifikatoren und Prompt-Filter",
                 "> 90 %", "der 12 publizierten Abwehrmaßnahmen wurden adaptiv umgangen",
                 "„Hat ein Agent untrusted Input aufgenommen, muss er so "
                 "eingeschränkt sein, dass dieser Input keine folgenreiche "
                 "Aktion auslösen kann.“", "1", "Tool-Inventar erstellen"])]
NOTES = ["Guardrails sind nicht nutzlos – als zusätzliche Schicht über Least Privilege."]


class SourceUnitsTest(unittest.TestCase):
    def test_instructions_and_icon_columns_are_not_slide_text(self):
        texts = [u.text for u in source_units(SOURCE)]
        self.assertIn("> 90 %", texts)
        self.assertIn("Tool-Inventar erstellen", texts)
        self.assertNotIn("Liste", texts)
        self.assertFalse(any(t.startswith("Struktur") for t in texts))
        self.assertFalse(any("Texte wörtlich" in t for t in texts))

    def test_quotes_inside_an_instruction_still_count(self):
        self.assertIn("Erkennung allein", [u.text for u in source_units(SOURCE)])


class FidelityTest(unittest.TestCase):
    def test_a_faithful_deck_is_clean(self):
        self.assertEqual(fidelity_findings(SOURCE, FAITHFUL, NOTES), [])

    def test_a_dropped_figure_and_a_shortened_quote_are_reported(self):
        deck = [(1, [t for t in FAITHFUL[0][1] if t not in ("> 90 %",)
                     and not t.startswith("„Hat")]
                 + ["Untrusted Input darf keine folgenreiche Aktion auslösen."])]
        missing = [u.text for u in missing_units(source_units(SOURCE),
                                                 deck[0][1], NOTES)]
        self.assertIn("> 90 %", missing)
        self.assertTrue(any(m.startswith("Hat ein Agent") for m in missing))

    def test_dropped_notes_are_reported(self):
        found = fidelity_findings(SOURCE, FAITHFUL, [""])
        self.assertTrue(any("speaker notes" in text for text, _ in found))

    def test_invented_text_from_the_sources_own_words_is_reported(self):
        deck = [(1, FAITHFUL[0][1] + [
            "Konkrete Filter für die Architektur der Angriffe erkennen"])]
        self.assertEqual(len(invented(deck, SOURCE)), 1)

    def test_template_wording_is_not_invented(self):
        deck = [(1, FAITHFUL[0][1] + ["Hier kann das Fazit der Folie stehen"])]
        self.assertEqual(invented(deck, SOURCE,
                                  {"Hier kann das Fazit der Folie stehen"}), [])


if __name__ == "__main__":
    unittest.main()
