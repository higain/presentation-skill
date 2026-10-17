"""QA a finished deck without looking at it.

``check(deck, template)`` returns a :class:`QAResult`.  The checks live in the
modules beside this one -- ``text``, ``colour``, ``layout``, ``prose`` -- and
``baseline`` measures what the template itself already does, so a template's
own habits are never reported as this build's defects.  ``allowlist`` lets a
deck record a finding as intentional, with a reason.

``scripts/qa.py`` is the command line; ``make_deck.py`` and ``build_deck`` call
:func:`check` and print :func:`summary`.
"""

from __future__ import annotations

import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from ..deckbuild import recorded_designs
from ..impress import PPTX_EXTS, load_document, notes_text, slides
from ..pack import active
from ..template import (_bleed_key, colliding, footer_text, measure_cache,
                        offslide, overflowing)
from . import cache as slide_cache
from .allowlist import _acknowledged, load_allowlist, stale_acknowledgements
from .baseline import _collision_key, _template_baseline
from .colour import (buried_text_issues, contrast_issues, occlusion_issues,
                     strike_issues)
from .layout import (_catalog_size, alignment_issues, balance_findings,
                     design_note, look_notes, sparse_slides, straddle_issues,
                     stray_template_images, variety_findings)
from .prose import lost_plan_titles, prose_findings
from .text import (PLACEHOLDER, _page_texts, emoji_issues,
                   headline_rival_issues, markdown_issues,
                   missing_title_issues, tiny_type_issues)


@dataclass
class QAResult:
    """What :func:`check` found.

    ``issues`` fail a build.  ``sparse``, ``notes`` and the prose findings are
    advisory and deliberately separate: a statement slide is sparse on purpose
    and a banned word can have an honest use, so a caller that fails a build on
    findings must not fail it on these.
    """
    issues: list = field(default_factory=list)
    checked: int = 0
    sparse: list = field(default_factory=list)      # [(slide, reason)]
    notes: list = field(default_factory=list)
    prose_issues: list = field(default_factory=list)
    titles: list = field(default_factory=list)      # the ghost deck
    acknowledged: list = field(default_factory=list)


def check(path, template=None, wanted=None, source=None) -> QAResult:
    """Every automated check on ``path``, gated by the deck's allowlist.

    An acknowledged finding (``<deck>-allowlist.json``) is moved from
    ``issues`` to ``acknowledged``, and an acknowledgement that matches
    nothing is itself reported there, so the list cannot quietly grow.
    """
    # Every check that reads a template fact -- the readable floor, the
    # palette a contrast check compares against -- has to read *this* deck's,
    # not whichever pack happens to be installed.
    with active(template), measure_cache():
        findings, checked, sparse, notes, prose = _check(path, template, wanted,
                                                         source)

    result = QAResult(checked=checked, sparse=sparse, notes=notes,
                      prose_issues=list(prose[0]), titles=list(prose[1]))
    allow = load_allowlist(path)
    for text, key in findings:
        if allow and key is not None and _acknowledged(allow, *key):
            result.acknowledged.append(text)
        else:
            result.issues.append(text)
    if allow:
        result.acknowledged += stale_acknowledgements(
            allow, [key for _, key in findings if key is not None])
    return result




def _fresh_page_checks(page, number, add, deck_texts, leftovers, footers,
                       base, design=None):
    """The cheap per-slide reads every run makes, cached slide or not.

    They feed the deck-wide checks -- template wording left on several
    slides, the footer, the ghost deck -- which compare slides with each
    other and so can never come out of a per-slide cache.
    """
    own = base.leftover_strings(design)
    for address, text in _page_texts(page):
        deck_texts.append((number, text))
        hit = PLACEHOLDER.search(text)
        if hit:
            add(
                f"slide {number} shape [{address}]: placeholder text "
                f"{hit.group(0)!r} in {text[:60]!r}"
            )
        elif text in own:
            leftovers.setdefault(text, []).append((number, address))

    # The footer is a page property, so it is not in _page_texts and
    # nothing above this line can see it.  See template.footer_text.
    foot = footer_text(page).strip()
    if foot and foot in base.footers:
        footers.setdefault(foot, []).append(number)


def _check(path, template, wanted, source=None):
    findings = []            # (text, allowlist key or None), in report order
    leftovers, footers, deck_texts = {}, {}, []

    def add(text, key=None):
        findings.append((text, key))

    # Whatever the template already does -- its wording, the decorations it
    # bleeds past the trim on purpose, its brand colour pairings, the shapes it
    # stacks -- is duplicated into the deck unchanged, and is not a defect this
    # build introduced.  One document load for all four.
    base = _template_baseline(template)
    tpl_strings, tpl_bleeds = base.strings, base.bleeds

    # Slides whose fingerprint matches the last run reuse its per-slide
    # findings; see ``cache.py``.
    prints = slide_cache.slide_fingerprints(path)
    engine = slide_cache.engine_key(template) if prints else None
    stored = slide_cache.load(path, engine) if prints else {}
    fresh = {}

    doc = load_document(path, read_only=True)
    designs = []
    hand = {}                # slide -> reasons to look at it first
    try:
        pages = slides(doc)
        recorded = recorded_designs(doc)

        def _design_of(number):
            return recorded[number - 1] if number <= len(recorded) else None

        if prints and len(prints) != len(pages):
            prints, stored = None, {}
        checked = 0
        for number, page in enumerate(pages, start=1):
            if wanted and number not in wanted:
                continue
            checked += 1
            print_ = prints[number - 1] if prints else None
            entry = stored.get(print_) if print_ else None
            if entry is not None:
                hand[number] = entry.get("hand")
                for packed in entry["a"]:
                    add(*slide_cache.unpack_finding(number, packed))
                _fresh_page_checks(page, number, add, deck_texts, leftovers,
                                   footers, base, _design_of(number))
                for packed in entry["b"]:
                    add(*slide_cache.unpack_finding(number, packed))
                continue
            start = len(findings)

            for address, reason in overflowing(doc, page):
                add(f"slide {number} shape [{address}]: {reason}")

            # One block of text laid over a row of table cells collides with
            # each cell separately, and the pair of texts is the same every
            # time -- so it is one defect and one line, with a count.  Same
            # principle as the off-slide grouping below.
            seen = {}
            for address, reason in colliding(doc, page):
                key = _collision_key(reason)
                if key in base.collisions:
                    continue
                if key in seen:
                    seen[key][2] += 1
                    continue
                seen[key] = [address, reason, 1]
            for address, reason, count in seen.values():
                more = (f" (and {count - 1} more block(s) it covers the same "
                        f"way -- one thing drawn over another)"
                        if count > 1 else "")
                add(f"slide {number} shape [{address}]: {reason}{more}")

            # One table drawn too wide is 11 shapes over the edge by the
            # same amount.  Report the defect once, with a count -- eleven
            # lines for one mistake reads as eleven mistakes.
            bleeds = {}
            for address, reason in offslide(page):
                amount, direction = _bleed_key(reason)
                # Round: the same object's cells differ by a unit of rounding.
                key = (round(amount / 50) * 50 if amount else amount, direction)
                if (direction, amount) in tpl_bleeds:
                    continue
                bleeds.setdefault(key, []).append((address, reason))
            for (amount, direction), group in sorted(
                    bleeds.items(), key=lambda kv: -kv[0][0]):
                address, reason = group[0]
                more = (f" (and {len(group) - 1} more shape(s) by the same amount"
                        f" -- one object drawn too wide)" if len(group) > 1 else "")
                add(f"slide {number} shape [{address}]: {reason}{more}")

            part_a = findings[start:]
            _fresh_page_checks(page, number, add, deck_texts, leftovers,
                               footers, base, _design_of(number))
            start = len(findings)

            for text in missing_title_issues(page, number):
                add(text, ("no-title", number, "title"))
            for text in (markdown_issues(page, number)
                         + emoji_issues(page, number, tpl_strings)
                         + tiny_type_issues(page, number, tpl_strings)
                         + contrast_issues(page, number, base.contrast)
                         + occlusion_issues(page, number, base.occlusion)
                         + buried_text_issues(page, number)
                         + headline_rival_issues(page, number)
                         + strike_issues(doc, page, number)):
                add(text)
            for text, key in straddle_issues(doc, page, number):
                add(text, key)
            for text, key in alignment_issues(page, number):
                add(text, key)
            hand[number] = look_notes(page, _page_texts(page))
            if print_:
                packs = [[slide_cache.pack_finding(number, text, key)
                          for text, key in part]
                         for part in (part_a, findings[start:])]
                if all(pack is not None for part in packs for pack in part):
                    fresh[print_] = {"a": packs[0], "b": packs[1], "n": number,
                                     "hand": hand.get(number)}

        # The paint metrics are per slide too: measured for the slides that
        # changed, read back for the rest.
        numbers = [n for n in range(1, len(pages) + 1)
                   if not wanted or n in wanted]
        cached = {n: stored[prints[n - 1]] for n in numbers
                  if prints and prints[n - 1] in stored}
        todo = set(numbers) - set(cached)
        new_sparse = dict(sparse_slides(doc, todo)) if todo else {}
        new_balance = dict(balance_findings(doc, todo)) if todo else {}
        sparse = []
        for number in numbers:
            if number in cached:
                reason, top = cached[number].get("sparse"), cached[number].get("balance")
            else:
                reason, top = new_sparse.get(number), new_balance.get(number)
                if prints and prints[number - 1] in fresh:
                    fresh[prints[number - 1]].update(sparse=reason, balance=top)
            # Advisory like a sparse slide: a reason to look first, which
            # the report's verdict table hands the vision-agent.
            reason = "; ".join(r for r in (reason, hand.get(number)) if r)
            if reason:
                sparse.append((number, reason))
            if top:
                add(f"slide {number}: {top}", ("balance", number, "top-heavy"))
        # A deck-wide note too: on a few slides "100% one design" is no news.
        notes = [note for note in (design_note(recorded_designs(doc),
                                               _catalog_size(template)),)
                 if note and not wanted]
        if not wanted:
            for text, key in variety_findings(recorded_designs(doc),
                                              _catalog_size(template)):
                add(text, key)
        designs = recorded_designs(doc) if not wanted else []
        prose = prose_findings(doc, wanted, tpl_strings)
        # Whole decks only.  A title the user approved is absent from a
        # three-slide check exactly the way it is absent from a deck that threw
        # it away, and the check cannot tell those apart on a partial view.
        if not wanted:
            lost, reworded = lost_plan_titles(path, deck_texts)
            for text in lost:
                add(text)
            prose = (prose[0] + reworded, prose[1])
            # The source the deck was built from: nothing dropped, nothing
            # changed, nothing invented.  Whole decks only, like the titles.
            if source:
                from .fidelity import fidelity_findings
                texts, spoken = [], []
                for number, page in enumerate(pages, start=1):
                    texts.append((number, [t for _, t in _page_texts(page)]))
                    try:
                        spoken.append(notes_text(page))
                    except Exception:
                        spoken.append("")
                markdown = Path(source).read_text(encoding="utf-8")
                for text, key in fidelity_findings(markdown, texts, spoken,
                                                   tpl_strings):
                    add(text, key)
    finally:
        doc.close(False)

    if prints:
        keep = dict(stored) if wanted else {
            p: stored[p] for p in prints if p in stored}
        for print_, entry in fresh.items():
            entry.pop("n", None)
            keep[print_] = entry
        slide_cache.save(path, engine, keep)

    # One label nobody replaced, on thirteen slides, is one thing to fix, and
    # the fix is one retext key.  It is still reported **per slide**, one line
    # each, because a grouped "slides 16, 27, 33 shape [6.5]" made the runs
    # that shipped these decks burn a transcript working out which slide each
    # finding was actually on -- three of them shared a design, one did not,
    # and the shape index meant something different on each.  A per-slide
    # line names exactly the slide, the shape and the text, and the design
    # each one sits on is in the same sentence.
    for text, places in sorted(leftovers.items(), key=lambda kv: -len(kv[1])):
        for number, address in places:
            design = designs[number - 1] if number <= len(designs) else None
            where = f"slide {number}"
            if design is not None:
                where += f" (design {design})"
            add(f"{where} shape [{address}]: still the template's own "
                f"wording -- {text[:60]!r}. Replace it with `retext`, or draw "
                f"over the design that carries it",
                ("leftover-wording", number, text))

    # Reported after the wording it looks like, and worded to say where it is
    # not: the shape hunt is what a run does next, and there is no shape.
    # Also per slide -- same reason as above: the footer is a page property
    # and its "shape" is the page, so a grouped line names no address at all.
    for text, numbers in sorted(footers.items(), key=lambda kv: -len(kv[1])):
        for number in sorted(numbers):
            add(
                f"slide {number} footer: still the template's own footer -- "
                f"{text[:60]!r}. It is a page property, not a shape: `retext`, "
                f"`find_shapes` and every other check here walk shapes and "
                f"cannot see it. Fix it with "
                f"`office.template.set_footer(page, ...)` in `fill`, or "
                f"`build_deck(..., footer='...')` for the whole deck; "
                f"`footer=None` hides it",
                ("footer", number, text))

    # Read from the zip, so it runs after the document is closed.
    if template and Path(path).suffix.lower() in PPTX_EXTS \
            and Path(template).suffix.lower() in PPTX_EXTS:
        for text, key in stray_template_images(path, template, wanted):
            add(text, key)

    return findings, checked, sparse, notes, prose


def validate_pptx(path, template=None):
    """Run ``office/validate.py``, baselining against the template when given.

    ``--original`` suppresses schema findings the template already had, so what
    is left is yours.  It does not cover broken relationships or undeclared
    content types -- those are repaired at the source, in ``impress.save_as``.
    """
    script = Path(__file__).absolute().parent.parent / "validate.py"
    cmd = [sys.executable, str(script), str(path)]
    if template and Path(template).suffix.lower() in PPTX_EXTS:
        cmd += ["--original", str(template)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    return proc.returncode, (proc.stdout + proc.stderr).strip()


def summary(result: QAResult, prefix: str = "", ghost: bool = True) -> list:
    """The findings as printable lines -- one formatter for every caller.

    ``prefix`` marks each line (``build_deck`` prints ``[build_deck] ``);
    ``ghost=False`` leaves the title list out, for a build that prints as it
    goes and would bury its findings under sixteen titles.
    """
    out = [f"{result.checked} slide(s) checked.", ""]
    for note in result.notes:
        out += [note, ""]
    if result.acknowledged:
        out += [f"## {len(result.acknowledged)} acknowledged (allowlist)", ""]
        out += [f"- {ack}" for ack in result.acknowledged] + [""]
    if result.issues:
        out += [f"## {len(result.issues)} issue(s) — fix every one", ""]
        out += [f"- {issue}" for issue in result.issues]
    else:
        out.append("No overflow, collisions, leftovers, markdown, contrast or "
                   "occlusion defects.")
    if result.sparse:
        out += ["", f"## {len(result.sparse)} slide(s) to look at", ""]
        out += [f"- slide {n}: {reason}" for n, reason in result.sparse]
        out += ["", "Not findings -- a statement slide is sparse on purpose, and a "
                "diagram no helper draws is placed by hand. "
                "Have the vision-agent look at these first."]
    if ghost and result.titles:
        out += ["", "## The ghost deck", ""]
        out += [f"{n:>3}. {title or '(no title found)'}"
                for n, title in result.titles]
        out += ["", "Read these in order. They should tell the whole argument "
                "on their own (references/storyline.md)."]
    if result.prose_issues:
        out += ["", f"## {len(result.prose_issues)} line(s) to look at", ""]
        out += [f"- {finding}" for finding in result.prose_issues]
        out += ["", "Not findings -- a banned word can have an honest use and a "
                "short title is right on a divider (references/language.md)."]
    return [prefix + line if line else line for line in out]
