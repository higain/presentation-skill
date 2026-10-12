"""The review ledger: which render the vision-agent saw, and proof that it saw it.

A build model with no eyes delegates every look at a slide to the
vision-agent and saves the reply as ``render/slide-N.answer.md``.  Nothing
used to tell a saved reply from one the build model typed itself, and one run
did exactly that: ``make_deck.py`` had deleted every answer on a rebuild, the
model "restored" twenty-nine of them from memory with a one-line script, the
gate passed, and three dividers shipped with invisible titles that the real
reviewer had flagged.

Two things here close that, and neither needs the build model's cooperation:

* **A code only eyes can read.**  ``vision_brief.py`` hands the vision-agent
  ``render/review/slide-N.png``: the render with a strip along the bottom
  that carries a random code.  The brief asks for it back in the answer's
  JSON.  The ledger keeps only the code's SHA-256, so the build model can read
  every file on disk and still not produce it.  An answer without the right
  code is not an answer.
* **Answers keyed by the render, not deleted with it.**  Each ledger entry
  records the SHA-256 of the PNG it was issued for.  After a re-render,
  :func:`sync` keeps the answer of every slide whose pixels did not change --
  following it if the slide moved -- and resets only the slides that did.
  Deleting all of them on every rebuild is what made re-typing them tempting:
  a one-slide fix used to cost twenty-nine new vision calls.

LibreOffice renders an unchanged slide to the same bytes, so the file hash is
the pixel hash.
"""

from __future__ import annotations

import hashlib
import json
import re
import secrets
from pathlib import Path

from .vision import CANNOT_SEE, last_json

#: No 0/O, 1/I/L, 2/Z, 5/S, 6/G, 8/B: the letters a vision model reads as
#: each other.  22 symbols, eight of them: 5.5e10 codes, past guessing and
#: past brute-forcing the hash from a shell loop.
ALPHABET = "ACDEFHJKMNPRTUVWXY3479"
CODE_LENGTH = 8

#: Beside the renders, out of every ``slide-*.png`` glob.
REVIEW_DIR = "review"
LEDGER_NAME = "ledger.json"

#: Fonts tried for the strip, in order; Pillow's own is the fallback.
STRIP_FONTS = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf",
)


def render_hash(png: Path) -> str:
    return hashlib.sha256(Path(png).read_bytes()).hexdigest()


def normalize(code) -> str:
    """``"ab3c-d7ef"`` -> ``"AB3CD7EF"``: case, dashes and spaces do not count."""
    return re.sub(r"[^A-Z0-9]", "", str(code or "").upper())


def code_hash(code) -> str:
    return hashlib.sha256(normalize(code).encode()).hexdigest()


def new_code() -> str:
    raw = "".join(secrets.choice(ALPHABET) for _ in range(CODE_LENGTH))
    return f"{raw[:4]}-{raw[4:]}"


def review_dir(outdir: Path) -> Path:
    return Path(outdir) / REVIEW_DIR


def review_image(outdir: Path, number: int) -> Path:
    return review_dir(outdir) / f"slide-{number}.png"


def answer_path(outdir: Path, number: int) -> Path:
    return Path(outdir) / f"slide-{number}.answer.md"


def brief_path(outdir: Path, number: int) -> Path:
    return Path(outdir) / f"slide-{number}.brief.md"


def renders(outdir: Path) -> dict:
    """``{number: png}`` for the slide renders in ``outdir``."""
    return {int(m.group(1)): png for png in Path(outdir).glob("slide-*.png")
            if (m := re.fullmatch(r"slide-(\d+)\.png", png.name))}


def load(outdir: Path) -> dict:
    """``{number: {"render": sha, "code": sha}}``; ``{}`` when there is none."""
    path = review_dir(outdir) / LEDGER_NAME
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    out = {}
    for key, entry in (raw.get("slides") or {}).items():
        if (str(key).isdigit() and isinstance(entry, dict)
                and entry.get("render") and entry.get("code")):
            out[int(key)] = {"render": entry["render"], "code": entry["code"]}
    return out


def save(outdir: Path, ledger: dict) -> None:
    folder = review_dir(outdir)
    folder.mkdir(parents=True, exist_ok=True)
    (folder / LEDGER_NAME).write_text(json.dumps(
        {"slides": {str(n): ledger[n] for n in sorted(ledger)}}, indent=1),
        encoding="utf-8")


def _font(size):
    from PIL import ImageFont
    for path in STRIP_FONTS:
        if Path(path).is_file():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default(size=size)


def stamp(png: Path, code: str, dest: Path) -> Path:
    """Write ``png`` with a strip *below* it carrying ``code``.

    Below, not over: the strip must not cover any of the slide, or the
    reviewer would be judging a different picture from the one that ships.
    """
    from PIL import Image, ImageDraw
    slide = Image.open(png).convert("RGB")
    w, h = slide.size
    strip = max(48, h // 9)
    out = Image.new("RGB", (w, h + strip), (255, 255, 255))
    out.paste(slide, (0, 0))
    draw = ImageDraw.Draw(out)
    draw.rectangle((0, h, w, h + 3), fill=(0, 0, 0))
    font = _font(int(strip * 0.55))
    text = f"REVIEW CODE  {code}"
    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
    draw.text(((w - (right - left)) // 2 - left,
               h + (strip - (bottom - top)) // 2 - top),
              text, fill=(0, 0, 0), font=font)
    dest.parent.mkdir(parents=True, exist_ok=True)
    out.save(dest)
    return dest


def _forget(outdir: Path, number: int) -> None:
    for path in (answer_path(outdir, number), brief_path(outdir, number),
                 review_image(outdir, number)):
        path.unlink(missing_ok=True)


def sync(outdir: Path) -> dict:
    """Carry the ledger over to a fresh set of renders.

    Returns ``{"kept": [...], "moved": {new: old}, "reset": [...]}``.  A slide
    whose render is byte-identical to one the ledger knows keeps that answer
    and review image -- under its new number if the slide moved.  Every other
    slide loses its answer, brief and review image, and so does every number
    that no longer has a render.
    """
    outdir = Path(outdir)
    old = load(outdir)
    current = {n: render_hash(png) for n, png in renders(outdir).items()}
    by_render = {}
    for number, entry in old.items():
        by_render.setdefault(entry["render"], number)

    # Read everything worth carrying before anything is deleted: a deck with a
    # slide inserted shifts every later one, and moving files one by one would
    # overwrite the next slide's answer before it was read.
    carried, kept, moved, reset = {}, [], {}, []
    for number, digest in sorted(current.items()):
        source = (number if old.get(number, {}).get("render") == digest
                  else by_render.get(digest))
        if source is None:
            reset.append(number)
            continue
        answer, image = answer_path(outdir, source), review_image(outdir, source)
        carried[number] = (
            old[source],
            answer.read_text(encoding="utf-8") if answer.is_file() else None,
            image.read_bytes() if image.is_file() else None)
        if source == number:
            kept.append(number)
        else:
            moved[number] = source

    # Everything is in memory now: clear the slate, then write back what
    # carries over, under the number it carries over to.
    for number in set(old) | set(current):
        _forget(outdir, number)

    ledger = {}
    for number, (entry, answer, image) in carried.items():
        if image is None:                  # nothing to show again: start over
            reset.append(number)
            continue
        ledger[number] = entry
        review_image(outdir, number).parent.mkdir(parents=True, exist_ok=True)
        review_image(outdir, number).write_bytes(image)
        if answer is not None:
            answer_path(outdir, number).write_text(answer, encoding="utf-8")
    save(outdir, ledger)
    return {"kept": sorted(n for n in kept if n in ledger),
            "moved": {n: o for n, o in moved.items() if n in ledger},
            "reset": sorted(reset)}


def _issue(png: Path, image: Path, answer: Path, entry) -> tuple:
    """``(entry, fresh)``: keep ``entry`` if it still describes ``png``.

    Otherwise stamp a new code onto ``image`` and delete ``answer``, which
    belonged to a picture that no longer exists.
    """
    digest = render_hash(png)
    if entry and entry["render"] == digest and image.is_file():
        return entry, False
    code = new_code()
    stamp(png, code, image)
    answer.unlink(missing_ok=True)
    return {"render": digest, "code": code_hash(code)}, True


def prepare(outdir: Path) -> dict:
    """``{number: (review_png, fresh)}`` for every render, issuing codes.

    A slide the ledger already holds for this exact render keeps its code and
    image, and so its answer -- ``fresh`` is False and it needs no new call
    unless that answer is missing.  Any other slide gets a new code, a new
    image, and loses whatever answer sat beside it.
    """
    outdir = Path(outdir)
    ledger = load(outdir)
    out = {}
    for number, png in sorted(renders(outdir).items()):
        image = review_image(outdir, number)
        ledger[number], fresh = _issue(png, image, answer_path(outdir, number),
                                       ledger.get(number))
        out[number] = (image, fresh)
    for number in set(ledger) - set(out):
        del ledger[number]
    save(outdir, ledger)
    return out


#: Why a saved answer does or does not count, by :func:`verify`.
STATUS = {
    "ok": "seen by the vision-agent",
    "unbriefed": "no brief was issued for this render -- run vision_brief.py "
                 "slides",
    "no-answer": "no answer saved",
    "cannot-see": "the vision-agent could not open the image",
    "no-code": "the answer carries no review code",
    "wrong-code": "the answer's review code is not the one on the image",
}


def verify(outdir: Path, number: int, ledger=None) -> tuple[str, dict]:
    """``(status, parsed_answer)`` for slide ``number``; see :data:`STATUS`.

    ``ok`` only when the ledger holds this exact render and the answer gives
    back the code printed on its review image.
    """
    outdir = Path(outdir)
    ledger = load(outdir) if ledger is None else ledger
    # Through ``renders``: a deck of ten or more slides renders as
    # ``slide-01.png``, and a path built from the number found no image and
    # called every answer of a 24-slide deck unbriefed.
    png = renders(outdir).get(number) or outdir / f"slide-{number}.png"
    return _check(png, answer_path(outdir, number), ledger.get(number))


def _check(png: Path, answer: Path, entry) -> tuple[str, dict]:
    """:func:`verify` on explicit paths, shared with the design catalog."""
    if (entry is None or not png.is_file()
            or entry["render"] != render_hash(png)):
        return "unbriefed", {}
    if not answer.is_file():
        return "no-answer", {}
    text = answer.read_text(encoding="utf-8")
    if CANNOT_SEE in text:
        return "cannot-see", {}
    parsed = last_json(text) or {}
    if not normalize(parsed.get("code")):
        return "no-code", parsed
    if code_hash(parsed["code"]) != entry["code"]:
        return "wrong-code", parsed
    return "ok", parsed


# ------------------------------------------------------- the design catalog
#
# ``presentation-setup`` has the vision-agent judge what each template design
# is for, one ``slide-NN/`` folder per design, and the same temptation: the
# build model saves each reply as ``answer.md`` and nothing told a saved reply
# from a typed one.  Same codes, same ledger rules, the catalog's own layout.

#: Beside ``manifest.json``.
DESIGN_LEDGER = "review-ledger.json"


def design_folder(catalog: Path, number: int) -> Path:
    return Path(catalog) / f"slide-{number:02d}"


def _load_designs(catalog: Path) -> dict:
    try:
        raw = json.loads((Path(catalog) / DESIGN_LEDGER).read_text(
            encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return {int(k): v for k, v in (raw.get("designs") or {}).items()
            if str(k).isdigit() and isinstance(v, dict)
            and v.get("render") and v.get("code")}


def prepare_designs(catalog: Path, numbers) -> dict:
    """``{number: (review_png, fresh)}`` for each design's ``render.png``."""
    catalog = Path(catalog)
    ledger, out = _load_designs(catalog), {}
    for number in numbers:
        folder = design_folder(catalog, number)
        if not (folder / "render.png").is_file():
            continue
        image = folder / "review.png"
        ledger[number], fresh = _issue(folder / "render.png", image,
                                       folder / "answer.md", ledger.get(number))
        out[number] = (image, fresh)
    (catalog / DESIGN_LEDGER).write_text(json.dumps(
        {"designs": {str(n): ledger[n] for n in sorted(ledger)}}, indent=1),
        encoding="utf-8")
    return out


def verify_design(catalog: Path, number: int, ledger=None) -> tuple[str, dict]:
    """:func:`verify` for one design of a template catalog."""
    ledger = _load_designs(catalog) if ledger is None else ledger
    folder = design_folder(catalog, number)
    return _check(folder / "render.png", folder / "answer.md",
                  ledger.get(number))
