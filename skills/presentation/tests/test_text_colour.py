"""Several text colours in one .pptx shape survive the save.

The docs used to say a .pptx keeps one colour per shape.  It was the theme
colour: a run from a .pptx carries a reference (``schemeClr``) that the export
writes instead of a plain ``CharColor``.  ``set_char_color`` clears it.  This
pins both halves, so neither the claim nor the fix can drift unnoticed.
"""

import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

from office.soffice import isolation_prefix  # noqa: E402

FIXTURE = SKILL / "tests" / "fixtures" / "kite-brand-master.pptx"
WORDS = ("RED_PARA", "BLUE_PARA", "GREEN_PARA")
COLOURS = (0xCC0000, 0x0000CC, 0x00AA00)


def _run_colours(deck):
    """``{word: colour reference}`` for each word's run in the saved XML."""
    out = {}
    with zipfile.ZipFile(deck) as z:
        for name in z.namelist():
            if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                continue
            xml = z.read(name).decode("utf-8")
            for word in WORDS:
                at = xml.find(word)
                if at < 0:
                    continue
                run = xml[xml.rfind("<a:r>", 0, at):at]
                found = re.findall(r'(srgbClr val="\w+"|schemeClr val="\w+")',
                                   run)
                out[word] = found[-1] if found else None
    return out


@unittest.skipUnless(shutil.which("soffice") and isolation_prefix(),
                     "needs LibreOffice and a network namespace")
class TextColourTest(unittest.TestCase):
    def _build(self, recolour):
        from office.deckbuild import build_deck
        from office.impress import shapes
        from office.template import set_text

        def fill(doc, page, slot, design):
            body = next(s for s in shapes(page)
                        if hasattr(s, "getString") and s.getString().strip())
            set_text(body, *WORDS)
            # What a run from a corporate .pptx carries: a theme reference.
            enum = body.getText().createEnumeration()
            while enum.hasMoreElements():
                para = enum.nextElement()
                para.CharColorTheme = 1
            recolour(body)

        out = Path(self.tmp.name) / "deck.pptx"
        build_deck(str(FIXTURE), out, [1], fill=fill, qa=False)
        return _run_colours(out)

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def test_set_char_color_keeps_one_colour_per_paragraph(self):
        from office.template import set_char_color

        def recolour(body):
            for i, colour in enumerate(COLOURS):
                set_char_color(body, colour, paragraph=i)

        self.assertEqual(self._build(recolour), {
            "RED_PARA": 'srgbClr val="cc0000"',
            "BLUE_PARA": 'srgbClr val="0000cc"',
            "GREEN_PARA": 'srgbClr val="00aa00"'})

    def test_a_bare_char_color_loses_to_the_theme(self):
        def recolour(body):
            enum = body.getText().createEnumeration()
            for colour in COLOURS:
                enum.nextElement().CharColor = colour

        # The export writes the theme colour -- as a reference on one
        # template, resolved to its RGB on another -- and never the colours
        # set: the three paragraphs come back as one.
        found = self._build(recolour)
        self.assertEqual(len(set(found.values())), 1, found)
        self.assertFalse({f'srgbClr val="{c:06x}"' for c in COLOURS}
                         & set(found.values()), found)


class PinnedTextColourTest(TextColourTest):
    """The same results with every text proxy pinned (the crash workaround)."""

    def setUp(self):
        super().setUp()
        from office.textenum import PIN_ENV
        patcher = mock.patch.dict(os.environ, {PIN_ENV: "1"})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _build(self, recolour):
        from office import impress, textenum
        found = super()._build(recolour)
        self.assertTrue(textenum._PINS, "nothing was pinned")
        impress.shutdown()                    # soffice gone: pins released
        self.assertEqual(textenum._PINS, [])
        return found


if __name__ == "__main__":
    unittest.main()
