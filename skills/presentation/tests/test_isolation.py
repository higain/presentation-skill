"""LibreOffice must not reach the network from an untrusted deck.

The live test builds a deck whose first slide links an image by URL, points
the URL at a listener on this machine, then renders and opens the deck the
way the skill does.  Before the isolation, one render sent six requests.
"""

import http.server
import io
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))

from office.external import external_resources  # noqa: E402
from office.helpers import check_archive_size  # noqa: E402
from office.soffice import isolation_prefix  # noqa: E402

FIXTURE = SKILL / "tests" / "fixtures" / "kite-brand-master.pptx"
IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
LINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"


def beacon_deck(dest: Path, url: str) -> Path:
    """The fixture, with one picture on slide 1 linked from ``url``."""
    with zipfile.ZipFile(FIXTURE) as zin, \
            zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "ppt/slides/_rels/slide1.xml.rels":
                data = data.replace(b"</Relationships>", (
                    f'<Relationship Id="rIdBeacon" Type="{IMAGE}" Target="{url}" '
                    f'TargetMode="External"/></Relationships>').encode())
            if info.filename == "ppt/slides/slide1.xml":
                data = data.replace(b"</p:spTree>", (
                    b'<p:pic><p:nvPicPr><p:cNvPr id="999" name="Beacon"/>'
                    b'<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>'
                    b'<a:blip r:link="rIdBeacon"/><a:stretch><a:fillRect/>'
                    b'</a:stretch></p:blipFill><p:spPr><a:xfrm>'
                    b'<a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>'
                    b'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                    b'</p:spPr></p:pic></p:spTree>'), 1)
            zout.writestr(info, data)
    return dest


class ExternalResourcesTest(unittest.TestCase):
    def _zip(self, parts):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for name, body in parts.items():
                z.writestr(name, body)
        path = Path(self._dir.name) / "x.pptx"
        path.write_bytes(buf.getvalue())
        return path

    def setUp(self):
        self._dir = tempfile.TemporaryDirectory()
        self.addCleanup(self._dir.cleanup)

    def test_linked_image_is_found_hyperlink_is_not(self):
        rels = (f'<Relationships><Relationship Id="a" Type="{IMAGE}" '
                f'Target="https://evil.example/x.png" TargetMode="External"/>'
                f'<Relationship Id="b" Type="{LINK}" Target="https://ok.example" '
                f'TargetMode="External"/></Relationships>')
        found = external_resources(self._zip({"ppt/slides/_rels/s.xml.rels": rels}))
        self.assertEqual([u for _, u in found], ["https://evil.example/x.png"])

    def test_odf_image_by_url(self):
        xml = ('<office:document-content><draw:image '
               'xlink:href="http://intranet/secret.png"/></office:document-content>')
        found = external_resources(self._zip({"content.xml": xml}))
        self.assertEqual([u for _, u in found], ["http://intranet/secret.png"])

    def test_the_shipped_fixture_links_nothing(self):
        self.assertEqual(external_resources(FIXTURE), [])


class ArchiveSizeTest(unittest.TestCase):
    def test_zip_bomb_is_refused(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("ppt/slides/slide1.xml", b"\0" * (50 * 1024 * 1024))
        with zipfile.ZipFile(buf) as z, self.assertRaises(ValueError):
            check_archive_size(z)

    def test_a_real_deck_passes(self):
        with zipfile.ZipFile(FIXTURE) as z:
            check_archive_size(z)


@unittest.skipUnless(shutil.which("soffice") and isolation_prefix(),
                     "needs LibreOffice and a network namespace")
class NoNetworkTest(unittest.TestCase):
    def setUp(self):
        hits = self.hits = []

        class Handler(http.server.BaseHTTPRequestHandler):
            def _hit(self):
                hits.append(self.path)
                self.send_response(404)
                self.end_headers()
            do_GET = do_HEAD = do_OPTIONS = _hit

            def log_message(self, *args):
                pass

        self.server = http.server.HTTPServer(("127.0.0.1", 0), Handler)
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.addCleanup(self.server.server_close)
        self.addCleanup(self.server.shutdown)
        self._dir = tempfile.TemporaryDirectory()
        self.addCleanup(self._dir.cleanup)
        port = self.server.server_address[1]
        self.deck = beacon_deck(Path(self._dir.name) / "beacon.pptx",
                                f"http://127.0.0.1:{port}/beacon.png")

    def test_render_and_open_fetch_nothing(self):
        for script in ("render.py", "dump.py"):
            subprocess.run([sys.executable, str(SKILL / "scripts" / script),
                            str(self.deck)], cwd=self._dir.name,
                           capture_output=True, timeout=300)
        self.assertEqual(self.hits, [])


if __name__ == "__main__":
    unittest.main()
