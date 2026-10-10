"""How the deck reads: titles, banned words, and titles lost from an approved plan."""

from __future__ import annotations

import re
from pathlib import Path

from ..impress import placeholder, slides
from ..template import walk_shapes
from .colour import _geometry
from .text import _norm, _page_texts, _text_sizes

#: The tells -- the words that make slide copy read as generated.  Matched
#: literally, case-insensitively, on word boundaries, so this list is the
#: specification and `references/language.md` is its prose copy.
#:
#: Word boundaries do most of the disambiguating work: `\bMehrwert\b` leaves
#: "Mehrwertsteuer" alone, `\bleverag(e|es|ed|ing)\b` catches the verb in every
#: form without touching "leverages" inside a quoted product name.  What they
#: cannot do is tell an honest use from a lazy one -- "robust standard errors"
#: and a literal "deep dive" session are both correct -- which is why these are
#: prose to look at and never fail a build.  See `prose_findings`.
BANNED = re.compile(
    r"\b("
    # English
    r"delve|leverag(?:e|es|ed|ing)|robust|seamless(?:ly)?|holistic(?:ally)?"
    r"|synerg(?:y|ies)|cutting[- ]edge|state[- ]of[- ]the[- ]art|empower(?:s|ing)?"
    r"|unlock(?:s|ing)?|unleash(?:es|ing)?|supercharg(?:e|es|ed|ing)"
    r"|transformative|game[- ]changing|best[- ]in[- ]class|world[- ]class"
    r"|next[- ]generation|paradigm shift|low[- ]hanging fruit|win[- ]win"
    r"|actionable insights?|key learnings|drive value|a wide range of"
    r"|various stakeholders|in today's fast[- ]paced|it's important to note"
    r"|at the end of the day"
    # German
    r"|ganzheitlich(?:e|er|es|en)?|nahtlos(?:e|er|es|en)?"
    r"|zukunftsweisend(?:e|er|es|en)?|wegweisend(?:e|er|es|en)?"
    r"|maßgeschneidert(?:e|er|es|en)?|Mehrwert|Synergien|Leuchtturmprojekt"
    r"|Quantensprung|Optimierungspotenzial(?:e|s)?|Handlungsfelder"
    r"|Erfolgsfaktoren|proaktiv(?:e|er|es|en)?|abzuholen|abholen"
    r"|auf Augenhöhe|in der heutigen schnelllebigen"
    r"|Herausforderungen meistern"
    r")\b",
    re.I,
)

#: How a deck should not end.  The last slide stays on screen through the whole
#: discussion, so it is the most-looked-at slide in the deck and the one most
#: often wasted.  Both languages, and "Q&A" without word boundaries because the
#: ampersand is not a word character.
CLOSING_WASTE = re.compile(
    r"^\s*(thank(s| you)?|questions\??|any questions\??|q\s*&\s*a"
    r"|danke|vielen dank|herzlichen dank|fragen\??|ende)\s*[.!?]*\s*$",
    re.I,
)

# What separates a title that makes a point from one that names a topic.
#
# A machine cannot see whether a sentence carries a claim -- that needs the
# evidence on the slide and the argument around it.  What it can see is the
# shape of the string, and three signals hold in German and English alike:
#
#   * a label is short.  "Marktgröße", "Next Steps", "Ausblick" are one to three
#     words; a sentence with a subject and a verb is rarely under four.
#   * a sentence ends with a period.  A label never does.  This is prescriptive
#     rather than descriptive -- language.md makes the period the convention
#     precisely so that the check is possible.
#   * a title that ends in a question mark is a question, and a title answers.
#
# Titles above the cap are the opposite defect: two lines at title size is the
# budget, and 14 words overruns it in English and sooner in German.
TITLE_MIN_WORDS = 4
TITLE_MAX_WORDS = 14


def _slide_title(page):
    """``(address, text)`` for the page's title, or ``None``.

    The ``Title`` placeholder first, because a slide built the way
    ``authoring.md`` asks has one.  A deck built from the company template
    often does not: its designs carry their headline in an ordinary shape that
    ``retext`` fills by key, and no placeholder is involved.

    So the fallback is the largest type in the top third of the slide -- which
    is what a headline is, on every design in the catalog.  It is a heuristic
    and it can pick the wrong shape on a slide whose biggest text is a giant
    number; the checks that use it are advisory for that reason among others.
    """
    shape = placeholder(page, "Title")
    if shape is not None:
        try:
            text = _norm(shape.getString())
        except Exception:
            text = ""
        if text:
            return "title", text

    try:
        height = page.Height
    except Exception:
        return None
    best = None
    for address, shape in walk_shapes(page):
        try:
            text = _norm(shape.getString())
        except Exception:
            continue
        box = _geometry(shape)
        if not text or not box or box[1] > height / 3:
            continue
        sizes = _text_sizes(shape)
        if not sizes:
            continue
        size = max(sizes)
        if best is None or size > best[0]:
            best = (size, address, text)
    return (best[1], best[2]) if best else None


def _title_findings(number, address, text):
    """Action-title defects in one title.  See ``TITLE_MIN_WORDS``.

    Length is not checked here.  A deck written to a house style is long in
    every title or none, and one line per slide saying so buries the four
    findings that are specific -- the same reason ``check`` groups a table's
    eleven off-slide cells into one line.  ``prose_findings`` reports it once.
    """
    out = []
    if text.endswith("?"):
        out.append(
            f"slide {number} [{address}]: the title asks a question -- "
            f"{text[:60]!r}. A title states the answer; the speaker asks"
        )
    elif len(text.split()) < TITLE_MIN_WORDS:
        out.append(
            f"slide {number} [{address}]: title {text[:60]!r} names a topic "
            f"rather than making a point. Correct on a section divider; "
            f"anywhere else, write the so-what as a sentence"
        )
    elif not text.endswith("."):
        out.append(
            f"slide {number} [{address}]: title {text[:60]!r} does not end in "
            f"a period, so it is probably a label rather than a claim "
            f"(references/language.md §1)"
        )
    return out


def repeated_titles(titles):
    """Findings for one headline used on more than one slide.

    This is the shape a deck takes when the approved titles were thrown away
    at build time and the branch name went into the title slot instead: the
    ghost deck reads as four labels, each repeated three times.  Unlike
    ``lost_plan_titles`` it needs no ``plan.md`` to see that, which matters
    because most decks have none.

    Dividers are ``None`` here and drop out, which is right -- two dividers
    named the same thing is a different mistake and a rarer one.
    """
    repeats = {}
    for number, text in titles:
        if text:
            repeats.setdefault(_norm(text).lower(), []).append((number, text))
    out = []
    for group in sorted(repeats.values(), key=len, reverse=True):
        if len(group) < 2:
            continue
        shown = ", ".join(str(number) for number, _ in group)
        text = group[0][1]
        out.append(
            f"slides {shown} share the title {text[:40]!r}. One title per "
            f"slide, each a claim the last leads to -- a headline repeated "
            f"is a section name sitting where the argument should be"
        )
    return out


def prose_findings(doc, wanted=None, baseline=frozenset()):
    """``(findings, titles)`` -- how the deck reads, and its titles in order.

    Not findings in the ``issues`` sense and deliberately not part of them: a
    banned word can have an honest use, a four-word title is right on a section
    divider, and a deck of one slide has no argument to flow.  Every line here
    is a judgement a person makes in a second and a machine cannot make at all,
    so this reports and does not fail a build -- the same bargain as
    ``sparse_slides``.

    ``titles`` is the ghost deck: the point is not that a machine grades it but
    that reading the titles alone, in order, tells you within ten seconds
    whether the deck has an argument.  Nothing else in this file replaces that.

    ``baseline`` is the template's own wording, so a tell that the brand chose
    for itself is not reported against every deck built on it.
    """
    findings, titles = [], []
    for number, page in enumerate(slides(doc), start=1):
        if wanted and number not in wanted:
            continue

        found = _slide_title(page)
        titles.append((number, found[1] if found else None))
        # Slide 1 is the title slide: a deck title is a name, not a claim.
        if found and number > 1:
            findings.extend(_title_findings(number, *found))
        elif not found:
            findings.append(
                f"slide {number}: no title found. Every content slide carries "
                f"its conclusion in the headline"
            )

        seen = set()
        for address, text in _page_texts(page):
            if text in baseline:
                continue
            for hit in BANNED.finditer(text):
                word = hit.group(0).lower()
                if word in seen:
                    continue
                seen.add(word)
                findings.append(
                    f"slide {number} [{address}]: {hit.group(0)!r} is on the "
                    f"banned list (references/language.md §3) -- in "
                    f"{text[:60]!r}. Say the specific thing, or keep it and "
                    f"move on"
                )

    # Same bargain for the period: a deck written without them -- the norm in
    # German business slides -- is one convention, not nine label findings.
    # Per slide only when the deck mixes the two, where the odd ones out are
    # the likely labels.
    no_period = [f for f in findings if "does not end in a period" in f]
    if len(no_period) >= 3:
        findings = [f for f in findings if f not in no_period]
        findings.append(
            f"{len(no_period)} titles end without a period. If they are "
            f"sentences, that is a convention, not a defect; if they are "
            f"labels, rewrite them as the slide's conclusion "
            f"(references/language.md §1)")

    # One line for the whole deck, naming the worst: a house style that runs
    # long runs long everywhere, and sixteen identical lines are read as
    # sixteen mistakes and then ignored.
    long_ones = [(len(text.split()), number) for number, text in titles
                 if text and len(text.split()) > TITLE_MAX_WORDS]
    if long_ones:
        words, worst = max(long_ones)
        findings.append(
            f"{len(long_ones)} of {len(titles)} titles run over "
            f"{TITLE_MAX_WORDS} words -- the longest is slide {worst} at "
            f"{words}. Over two lines at title size the headline stops being "
            f"scannable; cut each to its claim"
        )

    if not wanted:
        findings.extend(repeated_titles(titles))

    if titles and not wanted:
        number, last = titles[-1]
        if last and CLOSING_WASTE.match(last):
            findings.append(
                f"slide {number}: the deck ends on {last[:40]!r}. The last "
                f"slide stays up through the whole discussion -- put the ask "
                f"on it, with an owner and a date"
            )
    return findings, titles


#: A ``plan.md`` block's title, as stage 3 of ``references/storyline.md``
#: writes it: a ``## 4`` header, and under it ``title: <the action title>``.
PLAN_TITLE = re.compile(r"^\s*title\s*:\s*(.+?)\s*$", re.I)

#: How much of an approved title has to reach a slide before the difference is
#: an edit rather than a loss.  A build that tightened a long headline carried
#: the title over; a build that put the branch name there did not, and the two
#: need different words from this file.
TITLE_KEPT = 0.6


def plan_titles(deck):
    """``[(block, title)]`` -- the titles ``plan.md`` approved, in order.

    Stage 3 of ``references/storyline.md`` writes one ``## <n>`` block per
    slide with a ``title:`` line, **the user approves that file**, and stage 5
    carries those strings into the deck.  This reads them back so the transfer
    can be checked.  ``[]`` when there is no plan, which is most decks.

    **Beside the deck, and nowhere else.**  Not the working directory: run
    ``qa.py`` on someone else's deck from a storyline directory and every
    approved title is missing from it, which is sixteen confident findings
    about the wrong file.  These feed ``issues`` and fail a build, so the
    failure mode when the plan is somewhere unexpected has to be silence.
    Stage 5 writes the deck beside ``plan.md`` and this is one more reason to.

    Deliberately **not** matched to slide numbers: blocks are numbered from 0
    in some runs and 1 in others, the title and summary slides are prepended
    after the plan is approved, and none of that bears on the only question
    here -- did the sentence reach the deck.
    """
    plan = Path(deck).resolve().parent / "plan.md"
    if not plan.is_file():
        return []
    try:
        text = plan.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return []

    out, block = [], None
    for line in text.splitlines():
        if line.startswith("#"):
            block = line.lstrip("#").strip() or block
            continue
        hit = PLAN_TITLE.match(line)
        if hit:
            title = _norm(hit.group(1))
            if title:
                out.append((block if block else str(len(out)), title))
    return out


def _title_words(text):
    """A title reduced to the words worth comparing.

    Case and punctuation go.  A build that carried the sentence over and
    dropped the full stop transferred the title, and a check that calls that a
    lost title buries the case where the sentence is genuinely gone.
    """
    return {word for word
            in re.sub(r"[^\w\s]", " ", _norm(text).lower()).split() if word}


def lost_plan_titles(deck, texts):
    """``(lost, reworded)`` -- the approved titles, against the built deck.

    ``lost`` belongs in ``issues``: a sentence the user approved appears
    **nowhere** on any slide.  That is not a judgement and not a heuristic --
    the string is in the file or it is not -- which is why these fail a build
    when everything in ``prose_findings`` does not.  It is the defect the
    storyline layer exists to prevent and the one thing it could not see: a run
    carried an approved plan through stage 4, then wrote the branch name into
    the title slot on eleven of sixteen slides and passed every other check
    here.  The deck argued nothing and the report said so nowhere.

    ``reworded`` is prose: most of the title reached a slide but not all of it.
    Usually a build tightening a headline, occasionally the claim quietly
    changing after it was approved, never something to fail on.

    **Which slot the title landed in is not checked.** ``_slide_title`` cannot
    read the display title of designs 1, 5 and 7 at all, and on the title slide
    the governing thought belongs in the subtitle -- a placement check would
    report those every time it ran. Presence is what can be measured without
    lying, and presence is what was actually lost.
    """
    plan = plan_titles(deck)
    if not plan:
        return [], []

    per_slide = [(number, _title_words(text)) for number, text in texts]
    lost, reworded = [], []
    for block, title in plan:
        want = _title_words(title)
        if not want:
            continue
        best, where = 0.0, None
        for number, words in per_slide:
            if not words:
                continue
            share = len(want & words) / len(want)
            if share > best:
                best, where = share, number
        if best >= 1.0:
            continue
        if best >= TITLE_KEPT:
            reworded.append(
                f"plan.md block {block}: slide {where} carries a reworded "
                f"version of the approved title {title!r} -- fine if you "
                f"tightened it, not fine if the claim moved. The user "
                f"approved the sentence, not the topic"
            )
        else:
            lost.append(f"plan.md block {block}: {title!r}")

    # Twelve of these is one mistake made twelve times -- stage 5 writing
    # something other than the plan into the title slot -- so the reason is
    # given once and the rest are the payload: the exact sentences to copy.
    # They are not truncated for the same reason. A title is under fifteen
    # words by rule, and half of one cannot be pasted anywhere.
    if lost:
        head = "" if len(lost) == 1 else f"{len(lost)} approved titles are "
        lost[0] = (
            f"{lost[0]} -- {head or 'this approved title is '}on no slide of "
            f"this deck. Stage 5 carries the title from plan.md into the "
            f"title slot verbatim; a branch name or a topic label there "
            f"throws away the argument the user signed off "
            f"(references/storyline.md stage 5)"
        )
    return lost, reworded
