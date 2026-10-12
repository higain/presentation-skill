"""The review ledger: codes only eyes can read, answers keyed by render.

Pure file logic over small PNGs -- no LibreOffice needed.
"""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from PIL import Image  # noqa: E402

from office import review  # noqa: E402

CODE = "ACDE-FHJK"


def png(path, colour):
    Image.new("RGB", (160, 90), colour).save(path)


def answer(code=CODE, verdict="ok"):
    return ('STEP 1 ...\n{"code": "%s", "verdict": "%s", "findings": []}'
            % (code, verdict))


class ReviewTest(unittest.TestCase):
    def setUp(self):
        d = tempfile.TemporaryDirectory()
        self.addCleanup(d.cleanup)
        self.out = Path(d.name)

    def prepare(self):
        with mock.patch.object(review, "new_code", return_value=CODE):
            return review.prepare(self.out)

    def save(self, number, text):
        review.answer_path(self.out, number).write_text(text)

    def test_the_review_image_carries_a_strip_below_the_slide(self):
        png(self.out / "slide-1.png", "red")
        (image, fresh), = self.prepare().values()
        self.assertTrue(fresh)
        with Image.open(image) as im:
            self.assertEqual(im.size[0], 160)
            self.assertGreater(im.size[1], 90)
            self.assertEqual(im.getpixel((5, 5)), (255, 0, 0))   # slide intact
        self.assertNotIn(CODE.replace("-", ""),
                         (review.review_dir(self.out) / "ledger.json").read_text())

    def test_only_the_right_code_counts(self):
        png(self.out / "slide-1.png", "red")
        self.prepare()
        self.assertEqual(review.verify(self.out, 1)[0], "no-answer")
        self.save(1, '{"verdict": "ok", "findings": []}')
        self.assertEqual(review.verify(self.out, 1)[0], "no-code")
        self.save(1, answer("ACDE-FHJM"))
        self.assertEqual(review.verify(self.out, 1)[0], "wrong-code")
        self.save(1, answer("acde fhjk"))            # case and spacing are free
        self.assertEqual(review.verify(self.out, 1)[0], "ok")
        self.save(1, "CANNOT SEE IMAGE")
        self.assertEqual(review.verify(self.out, 1)[0], "cannot-see")

    def test_a_changed_render_is_unbriefed_until_prepared_again(self):
        png(self.out / "slide-1.png", "red")
        self.prepare()
        self.save(1, answer())
        png(self.out / "slide-1.png", "blue")
        self.assertEqual(review.verify(self.out, 1)[0], "unbriefed")

    def test_sync_keeps_unchanged_follows_moved_resets_changed(self):
        for n, colour in ((1, "red"), (2, "green"), (3, "blue")):
            png(self.out / f"slide-{n}.png", colour)
        self.prepare()
        for n in (1, 2, 3):
            self.save(n, answer())
        # A slide inserted at 2: green and blue move down, the new one is new.
        for n, colour in ((1, "red"), (2, "yellow"), (3, "green"),
                          (4, "blue")):
            png(self.out / f"slide-{n}.png", colour)
        moved = review.sync(self.out)
        self.assertEqual(moved["kept"], [1])
        self.assertEqual(moved["moved"], {3: 2, 4: 3})
        self.assertEqual(moved["reset"], [2])
        status = {n: review.verify(self.out, n)[0] for n in (1, 2, 3, 4)}
        self.assertEqual(status, {1: "ok", 2: "unbriefed", 3: "ok", 4: "ok"})
        self.assertFalse(review.answer_path(self.out, 2).exists())

    def test_prepare_keeps_an_answer_whose_render_did_not_change(self):
        png(self.out / "slide-1.png", "red")
        png(self.out / "slide-2.png", "green")
        self.prepare()
        self.save(1, answer())
        self.save(2, answer())
        png(self.out / "slide-2.png", "blue")
        fresh = {n: f for n, (_, f) in self.prepare().items()}
        self.assertEqual(fresh, {1: False, 2: True})
        self.assertTrue(review.answer_path(self.out, 1).exists())
        self.assertFalse(review.answer_path(self.out, 2).exists())

    def test_codes_avoid_letters_a_reader_confuses(self):
        for _ in range(50):
            code = review.new_code()
            self.assertRegex(code, r"^[%s]{4}-[%s]{4}$"
                             % (review.ALPHABET, review.ALPHABET))
        self.assertFalse(set("0O1IL2Z5S6G8B") & set(review.ALPHABET))


if __name__ == "__main__":
    unittest.main()
