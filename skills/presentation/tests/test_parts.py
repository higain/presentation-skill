"""Big decks: one file per slide, a split source, stale locks, the API page.

A 58-slide spec rebuild by a 200k-context agent read the 288 KB spec whole,
lost it to pruning, parsed it with a regex that silently emptied a field on
every slide, and died on a LibreOffice lock its own crash had left.  These pin
the pieces that replace that: ``split_source.py``, ``build_parts`` with its
``--parts`` preview and its all-failures report, ``check_lock``, ``D.para``
and the generated ``references/api.md``.
"""

import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
import zipfile
from contextlib import redirect_stderr
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from office.soffice import isolation_prefix  # noqa: E402

FIXTURE = SKILL / "tests" / "fixtures" / "kite-brand-master.pptx"
HAVE_SOFFICE = bool(shutil.which("soffice") and isolation_prefix())


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(textwrap.dedent(text), encoding="utf-8")


PART = """\
    import deck_common as C
    DESIGN = 1
    def fill(doc, page):
        C.say(doc, page, "Slide {n} says something")
    """

COMMON = """\
    from office.impress import mm, add_textbox
    def say(doc, page, text):
        add_textbox(doc, page, mm(20), mm(20), mm(200), mm(20), text)
    """


class TempDir(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.dir = Path(self._tmp.name)


class LoadPartsTest(TempDir):
    def _load(self):
        from office.deckbuild import load_parts
        with redirect_stderr(io.StringIO()):
            return load_parts(self.dir / "slides")

    def test_parts_load_in_number_order_with_the_common_module(self):
        _write(self.dir / "slides" / "_common.py", COMMON)
        for n in (2, 10, 1):
            _write(self.dir / "slides" / f"s{n:02d}.py", PART.format(n=n))
        for n in range(3, 10):
            _write(self.dir / "slides" / f"{n}-x.py", PART.format(n=n))
        parts = self._load()
        self.assertEqual([n for n, _, _ in parts], list(range(1, 11)))
        self.assertTrue(all(m.DESIGN == 1 for _, _, m in parts))

    def test_every_problem_is_reported_at_once(self):
        from office.deckbuild import PartsError
        _write(self.dir / "slides" / "_common.py", COMMON)
        _write(self.dir / "slides" / "s01.py", PART.format(n=1))
        _write(self.dir / "slides" / "s1-again.py", PART.format(n=1))
        _write(self.dir / "slides" / "s03.py", "def fill(doc, page): pass\n")
        _write(self.dir / "slides" / "s04.py", "DESIGN = 1\ndef fill(:\n")
        _write(self.dir / "slides" / "notes.py", "DESIGN = 1\n")
        with self.assertRaises(PartsError) as caught:
            self._load()
        msg = str(caught.exception)
        self.assertIn("both slide 1", msg)
        self.assertIn("no file for slide(s) 2", msg)
        self.assertIn("s03.py: DESIGN", msg)
        self.assertIn("s04.py: SyntaxError", msg)
        self.assertIn("notes.py: no slide number", msg)

    def test_a_preview_loads_only_its_own_parts(self):
        from office.deckbuild import load_parts
        _write(self.dir / "slides" / "_common.py", COMMON)
        _write(self.dir / "slides" / "s01.py", PART.format(n=1))
        _write(self.dir / "slides" / "s03.py", PART.format(n=3))
        _write(self.dir / "slides" / "s05.py", "half written (\n")
        parts = load_parts(self.dir / "slides", wanted=[3])
        self.assertEqual([n for n, _, _ in parts], [3])

    def test_preview_name_never_is_the_deck(self):
        from office.deckbuild import parts_preview
        self.assertEqual(parts_preview("out/deck.pptx", "23,25-27").name,
                         "deck-parts-23_25_27.pptx")
        self.assertNotEqual(parts_preview("deck.pptx", "1").name, "deck.pptx")


class LockTest(TempDir):
    def _lock(self, profile):
        deck = self.dir / "deck.pptx"
        deck.write_bytes(b"")
        from office.impress import lock_file
        lock = lock_file(deck)
        lock.write_text(f",maxi,debian,13.10.2026 11:47,file://{profile};")
        return deck, lock

    def test_a_dead_own_instance_lock_is_removed(self):
        from office.impress import check_lock
        deck, lock = self._lock("/tmp/lo-uno-profile-gone12345")
        check_lock(deck)
        self.assertFalse(lock.exists())

    def test_a_live_own_instance_lock_stays(self):
        from office.impress import SofficeError, check_lock
        # This test process carries the profile path on its command line.
        profile = "/tmp/lo-uno-profile-" + "alive"
        deck, lock = self._lock(profile)
        proc = subprocess.Popen([sys.executable, "-c",
                                 "import time; time.sleep(30)", profile])
        self.addCleanup(proc.kill)
        with self.assertRaises(SofficeError):
            check_lock(deck)
        self.assertTrue(lock.exists())

    def test_someone_elses_lock_is_refused_with_the_way_out(self):
        from office.impress import SofficeError, check_lock
        deck, lock = self._lock("/home/maxi/.config/libreoffice/4")
        with self.assertRaises(SofficeError) as caught:
            check_lock(deck)
        self.assertIn("open in another LibreOffice (maxi since 13.10.2026",
                      str(caught.exception))
        self.assertIn(str(lock), str(caught.exception))
        self.assertTrue(lock.exists())

    def test_no_lock_is_nothing(self):
        from office.impress import check_lock
        deck = self.dir / "deck.pptx"
        deck.write_bytes(b"")
        check_lock(deck)


class SplitSourceTest(TempDir):
    SOURCE = """\
        # Spezifikation

        ## 1. Designsystem
        ### 1.1 Raster
        Farben hier.
        ### 1.2 Farben
        Blau.

        ## Folien im Detail

        ### Folie 1: Titel
        - Text eins
        ```
        ### Folie 9: im Code, keine Folie
        ```
        ### Folie 2: Agenda
        - Text zwei

        ## Anhang
        Quellen.
        """

    def _split(self, *extra):
        _write(self.dir / "src.md", self.SOURCE)
        return subprocess.run(
            [sys.executable, str(SCRIPTS / "split_source.py"),
             str(self.dir / "src.md"), "--out", str(self.dir / "spec"), *extra],
            capture_output=True, text=True)

    def test_slides_shared_part_and_index(self):
        run = self._split()
        self.assertEqual(run.returncode, 0, run.stderr)
        spec = self.dir / "spec"
        one = (spec / "slide-01.md").read_text()
        self.assertIn("### Folie 1: Titel", one)
        self.assertIn("### Folie 9: im Code", one)   # fenced: not a heading
        self.assertNotIn("Agenda", one)
        two = (spec / "slide-02.md").read_text()
        self.assertNotIn("Anhang", two)               # a higher heading ends it
        shared = (spec / "_shared.md").read_text()
        for text in ("Farben hier.", "## Folien im Detail", "Quellen."):
            self.assertIn(text, shared)
        self.assertNotIn("Text eins", shared)
        self.assertIn("| 2 | slide-02.md |", (spec / "index.md").read_text())

    def test_an_explicit_heading(self):
        run = self._split("--heading", r"^### 1\.(\d+)")
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertIn("Blau.", (self.dir / "spec" / "slide-02.md").read_text())


class ApiReferenceTest(unittest.TestCase):
    def test_api_md_is_current(self):
        sys.path.insert(0, str(SCRIPTS))
        import api_ref
        current = (SKILL / "references" / "api.md").read_text(encoding="utf-8")
        self.assertEqual(current, api_ref.render(),
                         "references/api.md is stale: run "
                         "scripts/api_ref.py --write")

    def test_build_parts_and_para_are_listed(self):
        text = (SKILL / "references" / "api.md").read_text(encoding="utf-8")
        for name in ("build_parts(", "para(text", "check_lock(", "text(doc"):
            self.assertIn(name, text)


@unittest.skipUnless(HAVE_SOFFICE, "needs LibreOffice and a network namespace")
class BuildPartsTest(TempDir):
    def _setup(self, broken=()):
        _write(self.dir / "slides" / "_common.py", COMMON)
        for n in (1, 2, 3):
            body = textwrap.dedent(PART.format(n=n)).rstrip() + "\n"
            if n == 2:
                body += "    page.Visible = False\n"
            if n in broken:
                body += f"    raise KeyError('broken {n}')\n"
            (self.dir / "slides" / f"s{n:02d}.py").write_text(body)
        _write(self.dir / "build.py", f"""\
            from office.deckbuild import build_parts
            build_parts({str(FIXTURE)!r}, "deck.pptx", "slides", qa=False)
            """)

    def _make(self, *extra):
        return subprocess.run(
            [sys.executable, str(SCRIPTS / "make_deck.py"), "build.py",
             "deck.pptx", "--build-only", *extra],
            cwd=self.dir, capture_output=True, text=True, timeout=600)

    def _texts(self, deck):
        with zipfile.ZipFile(deck) as z:
            pres = z.read("ppt/presentation.xml").decode()
            rels = z.read("ppt/_rels/presentation.xml.rels").decode()
            out = []
            for rid in re.findall(r'<p:sldId [^>]*r:id="(rId\d+)"', pres):
                target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels)
                if target is None:
                    target = re.search(rf'Target="([^"]+)"[^>]*Id="{rid}"',
                                       rels)
                xml = z.read("ppt/" + target.group(1)).decode()
                said = re.search(r"Slide (\d) says", xml)
                out.append((said.group(1), 'show="0"' in xml))
            return out

    def test_the_deck_and_a_preview(self):
        self._setup()
        run = self._make()
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertEqual(self._texts(self.dir / "deck.pptx"),
                         [("1", False), ("2", True), ("3", False)])
        before = (self.dir / "deck.pptx").read_bytes()

        run = self._make("--parts", "3,1")
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)
        self.assertEqual(self._texts(self.dir / "deck-parts-3_1.pptx"),
                         [("1", False), ("3", False)])
        self.assertEqual((self.dir / "deck.pptx").read_bytes(), before)

    def test_every_failing_slide_is_reported(self):
        self._setup(broken=(1, 3))
        run = self._make()
        self.assertNotEqual(run.returncode, 0)
        err = run.stderr
        self.assertIn("fill failed on 2 slide(s)", err)
        self.assertIn("slide 1: KeyError", err)
        self.assertIn("slide 3: KeyError", err)
        self.assertIn("s03.py", err)
        self.assertFalse((self.dir / "deck.pptx").exists())

    def test_para_styles_one_paragraph(self):
        from office import draw as D
        from office.deckbuild import build_deck
        from office.impress import mm

        def fill(doc, page, slot, design):
            D.text(doc, page, mm(20), mm(20), mm(200), mm(40),
                   D.para("HEAD_PARA", size=20, weight=D.BOLD,
                          color=0xCC2B21),
                   "BODY_PARA", size=12, color=0x5A6367, valign="top")

        out = self.dir / "para.pptx"
        build_deck(str(FIXTURE), out, [1], fill=fill, qa=False)
        with zipfile.ZipFile(out) as z:
            xml = next(z.read(n).decode() for n in z.namelist()
                       if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)
                       and b"HEAD_PARA" in z.read(n))

        def run_props(word):
            at = xml.find(word)
            return xml[xml.rfind("<a:rPr", 0, at):at]

        head, body = run_props("HEAD_PARA"), run_props("BODY_PARA")
        self.assertIn('sz="2000"', head)
        self.assertIn('b="1"', head)
        self.assertIn('srgbClr val="cc2b21"', head)
        self.assertIn('sz="1200"', body)
        self.assertNotIn('b="1"', body)
        self.assertIn('srgbClr val="5a6367"', body)


if __name__ == "__main__":
    unittest.main()
