"""The per-slide QA cache: fingerprints that survive a re-save, and packing.

Pure file logic over hand-built .pptx zips -- no LibreOffice needed.
"""

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from office.qa import cache  # noqa: E402

PRES = ('<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst>'
        '<p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/>'
        '</p:sldIdLst></p:presentation>')
PRES_RELS = ('<Relationships>'
             '<Relationship Id="rId1" Type="x/slide" Target="slides/slide1.xml"/>'
             '<Relationship Id="rId2" Type="x/slide" Target="slides/slide2.xml"/>'
             '</Relationships>')


def slide(text, shape_id=2):
    return (f'<p:sld><p:cNvPr id="{shape_id}" name="T"/>'
            f'<a:t>{text}</a:t></p:sld>')


def rels(notes="notesSlide1.xml"):
    return ('<Relationships>'
            '<Relationship Id="rId1" Type="x/slideLayout" '
            'Target="../slideLayouts/slideLayout1.xml"/>'
            f'<Relationship Id="rId2" Type="x/notesSlide" '
            f'Target="../notesSlides/{notes}"/></Relationships>')


def write_deck(path, one, two, layout="<layout/>", notes="n"):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/presentation.xml", PRES)
        z.writestr("ppt/_rels/presentation.xml.rels", PRES_RELS)
        z.writestr("ppt/slides/slide1.xml", one)
        z.writestr("ppt/slides/slide2.xml", two)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", rels())
        z.writestr("ppt/slides/_rels/slide2.xml.rels", rels("notesSlide2.xml"))
        z.writestr("ppt/slideLayouts/slideLayout1.xml", layout)
        z.writestr("ppt/notesSlides/notesSlide1.xml", f"<notes>{notes}</notes>")
        z.writestr("ppt/notesSlides/notesSlide2.xml", "<notes/>")


class FingerprintTest(unittest.TestCase):
    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.dir = Path(d.name)

    def prints(self, *args, **kw):
        path = self.dir / f"deck{len(list(self.dir.iterdir()))}.pptx"
        write_deck(path, *args, **kw)
        return cache.slide_fingerprints(path)

    def test_renumbered_shape_ids_do_not_count(self):
        # A save renumbers shape ids deck-wide; an edit on slide 1 must not
        # make slide 2 look changed.
        a = self.prints(slide("A", 2), slide("B", 3))
        b = self.prints(slide("A", 2), slide("B", 7))
        self.assertEqual(a, b)

    def test_a_text_change_changes_that_slide_only(self):
        a = self.prints(slide("A"), slide("B"))
        b = self.prints(slide("A"), slide("C"))
        self.assertEqual(a[0], b[0])
        self.assertNotEqual(a[1], b[1])

    def test_the_layout_counts_and_the_notes_do_not(self):
        a = self.prints(slide("A"), slide("B"))
        self.assertNotEqual(a, self.prints(slide("A"), slide("B"),
                                           layout="<layout2/>"))
        self.assertEqual(a, self.prints(slide("A"), slide("B"),
                                        notes="other notes"))

    def test_not_a_pptx_is_none(self):
        path = self.dir / "deck.odp"
        path.write_bytes(b"not a zip")
        self.assertIsNone(cache.slide_fingerprints(path))


class PackTest(unittest.TestCase):
    def test_a_finding_moves_with_its_slide(self):
        packed = cache.pack_finding(
            3, "slide 3: no title found", ("no-title", 3, "title"))
        self.assertEqual(packed, [": no title found", ["no-title", "#", "title"]])
        moved = cache.pack_finding(2, "slide 2 shape [4]: overruns", None)
        self.assertEqual(cache.unpack_finding(7, moved),
                         ("slide 7 shape [4]: overruns", None))
        self.assertEqual(cache.unpack_finding(5, packed),
                         ("slide 5: no title found", ("no-title", 5, "title")))

    def test_a_finding_not_about_the_slide_is_not_packed(self):
        self.assertIsNone(cache.pack_finding(3, "deck: too many designs", None))

    def test_the_cache_is_dropped_when_the_engine_changes(self):
        with tempfile.TemporaryDirectory() as d:
            deck = Path(d) / "deck.pptx"
            cache.save(deck, "engine-1", {"abc": {"a": [], "b": []}})
            self.assertEqual(cache.load(deck, "engine-1"),
                             {"abc": {"a": [], "b": []}})
            self.assertEqual(cache.load(deck, "engine-2"), {})


if __name__ == "__main__":
    unittest.main()
