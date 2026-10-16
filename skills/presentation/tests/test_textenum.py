"""``enumerate_text`` pins text proxies only when asked to.

The workaround for LibreOffice's double free on released paragraph proxies
(``PRESENTATION_PIN_TEXT_PROXIES``) must leave the default path untouched and,
when on, keep the enumeration *and* every element alive until
``release_pins()``.  No LibreOffice needed: fake UNO objects stand in.
"""

import os
import sys
import unittest
from pathlib import Path
from unittest import mock

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

from office import textenum  # noqa: E402
from office.textenum import PIN_ENV, enumerate_text, release_pins  # noqa: E402


class _Enum:
    def __init__(self, items):
        self.items = list(items)

    def hasMoreElements(self):
        return bool(self.items)

    def nextElement(self):
        return self.items.pop(0)


class _Text:
    def __init__(self, *items):
        self.items = items

    def createEnumeration(self):
        return _Enum(self.items)


class EnumerateTextTest(unittest.TestCase):
    def setUp(self):
        release_pins()
        self.addCleanup(release_pins)

    def test_off_yields_everything_and_pins_nothing(self):
        with mock.patch.dict(os.environ, {PIN_ENV: ""}):
            self.assertEqual(list(enumerate_text(_Text("a", "b", "c"))),
                             ["a", "b", "c"])
        self.assertEqual(textenum._PINS, [])

    def test_on_pins_the_enumeration_and_every_element(self):
        with mock.patch.dict(os.environ, {PIN_ENV: "1"}):
            self.assertEqual(list(enumerate_text(_Text("a", "b"))), ["a", "b"])
        self.assertEqual(len(textenum._PINS), 3)
        self.assertIsInstance(textenum._PINS[0], _Enum)
        self.assertEqual(textenum._PINS[1:], ["a", "b"])
        self.assertEqual(release_pins(), 3)
        self.assertEqual(textenum._PINS, [])

    def test_the_variable_is_read_on_every_call(self):
        for value, pinned in (("true", True), ("on", True), ("0", False),
                              ("no", False)):
            with mock.patch.dict(os.environ, {PIN_ENV: value}):
                self.assertEqual(textenum.pinning(), pinned, value)

    def test_a_textless_object_raises_on_first_next(self):
        it = enumerate_text(object())       # the call itself does not raise
        with self.assertRaises(AttributeError):
            next(it)


if __name__ == "__main__":
    unittest.main()
