"""The briefs the vision-agent gets, and the parsing of what it answers.

The build model cannot see images, so every look at a render is delegated to a
vision model.  Left to write that brief itself, a build model writes "describe
these nine slides" and gets "all ok" back for a deck a designer would return.
Measured on gemma-4-31b against slides with planted defects, three things made
the difference, and every brief here keeps them:

* **one image per call** -- nine slides in one request are judged as a set,
  and the set is "fine";
* **describe, then judge** -- a 3x3 grid inventory first, so the answers are
  about what is on the slide rather than what slides usually have;
* **closed questions** with "unsure" allowed and a location required for every
  "yes" -- open "find problems" invites invented defects.

What the vision model is **never** asked: colours, sizes, counts, spelling, or
leftover template wording.  Those are measured from the file (``qa.py``,
``count_style.py``, ``structure.json``) and a vision model guessing them is
strictly worse -- it named a brand's cyan accent "yellow-orange" and the setup
run that trusted it picked the wrong accent.
"""

from __future__ import annotations

import json

#: What a design can carry.  One list, shared by the design brief, the
#: worksheet and the decision table, so a type the vision model proposes is a
#: type the table knows how to place.
CONTENT_TYPES = (
    "section-divider", "cover", "agenda", "one-big-idea", "bulleted-content",
    "comparison", "cards", "process", "timeline", "cycle", "kpis", "table",
    "chart", "matrix", "org-chart", "image", "quote", "team", "closing",
)

#: The ten questions of the slide review, in the order they are asked.  Kept
#: as data so the collected answers can be printed by name.
SLIDE_QUESTIONS = {
    "Q1": "HEADLINE",
    "Q2": "FOCAL POINT",
    "Q3": "EMPTY",
    "Q4": "OFF-CENTRE",
    "Q5": "OVERLAP",
    "Q6": "PEERS",
    "Q7": "HOLLOW",
    "Q8": "BROKEN WORD",
    "Q9": "LANGUAGE",
    "Q10": "UNFINISHED",
}

SLIDE_BRIEF = """\
You are a strict presentation designer checking ONE rendered slide before it is \
shown to senior management. Open the image with your read tool: {image}

A separate tool has already checked spelling, leftover template text, text \
overflow, colour contrast and font sizes from the file itself. Do NOT report \
those.

Deck language: {language}. Common English technical terms used inside another \
language (Backend, Team, Budget, Go-live, Release, Sprint) are normal. Only \
report a whole phrase or sentence in another language.
Measured by the tool for this slide: {measured}

STEP 1 - DESCRIBE, do not judge yet. Split the slide into a 3x3 grid \
(top-left, top-centre, top-right, middle-left, centre, middle-right, \
bottom-left, bottom-centre, bottom-right). In 3-6 short lines, say what sits \
in each non-empty cell and how big it is.

STEP 2 - ANSWER every question with yes, no or unsure. For each "yes" name the \
grid cell and the element. Never guess: if you cannot see it clearly, write \
unsure.
Q1 HEADLINE: Is there NO clear headline at the top - only a small label, or \
nothing, instead of a title in large type?
Q2 FOCAL POINT: Is there NO single element the eye goes to first (a diagram, \
chart, big number, image) - the slide is only small text?
Q3 EMPTY: Is a whole band (top, middle or bottom third) or a left/right third \
empty while the content is small or crowded elsewhere?
Q4 OFF-CENTRE: Is the content block clearly closer to one side, leaving the \
other side much emptier?
Q5 OVERLAP: Does any text or box sit on top of another text or box and cover \
part of it?
Q6 PEERS: Among repeated items (cards, boxes, steps, columns), does one look \
different - colour, size, height, alignment or text style?
Q7 HOLLOW: Do boxes or cards consist mostly of empty space, or is a diagram \
tiny compared with the space it has?
Q8 BROKEN WORD: Is a word split in the middle across two lines (e.g. \
"Manage / ment")?
Q9 LANGUAGE: Is a whole phrase or sentence in a language other than the deck \
language?
Q10 UNFINISHED: Would a professional designer send this slide back as \
unfinished or amateurish? Give the single main reason.

Finish with exactly one JSON object and nothing after it:
{{"verdict": "ok" or "fix", "findings": [{{"q": "Q1".."Q10", "where": \
"<grid cell>", "what": "<concrete element and problem>", "fix": "<one concrete \
suggestion>", "severity": "high"|"medium"|"low"}}]}}
"ok" only if every answer is "no". A clean slide is a normal, good result; do \
not invent problems.
"""

DESIGN_BRIEF = """\
You are cataloguing ONE slide design from a company's presentation template. \
Open the image with your read tool: {image}

A text-only assistant that cannot see images will later read your answer to \
decide when to use this design, so describe it in words it can match against \
content. The text on the slide is sample text: ignore what it says and judge \
the LAYOUT. Colours, fonts, sizes and shape counts are measured elsewhere - do \
not report them.

Measured facts for this design (trust these over your eyes for counts):
{facts}

STEP 1 - LOOK. In 3-5 short lines describe the layout: where the title sits, \
what the main visual structure is (columns, cards, circles, arrows, chart, \
table, photo, map...), how many repeated items there are and how they are \
arranged.

STEP 2 - ANSWER. Output exactly one JSON object and nothing after it:
{{
 "name": "<2-5 words a person would say, e.g. 'Three KPI circles'>",
 "purpose": "<one sentence: what message this slide makes visible>",
 "content_types": [one or more of: {types}],
 "content_shape": "<the exact content it needs, as a pattern: e.g. '3 numbers, \
each with a 1-2 word label and a 1-sentence explanation'>",
 "use_when": "<one sentence starting with 'You have ...'>",
 "avoid_when": "<one sentence: content that breaks this design, with a number, \
e.g. 'more than 4 steps' or 'labels longer than 3 words'>",
 "max_items": <integer: how many repeated items it holds, or 1>,
 "looks_best": "<one tip that keeps it looking designed when filled, e.g. \
'keep each label to one line'>",
 "imagery": [{{"depicts": "<what the picture concretely shows>", "role": \
"photo-placeholder|decorative|icon|logo|background"}}]
}}
"""

#: What the language line says when nobody named one.  Deliberately not a
#: guess at a language: the titles are on the slide, and the model can read
#: them.
DEFAULT_LANGUAGE = "the language of the slide's title"


def slide_brief(image: str, measured: str = "", language: str = "") -> str:
    return SLIDE_BRIEF.format(image=image,
                              measured=measured or "nothing unusual",
                              language=language or DEFAULT_LANGUAGE)


def design_facts(structure: dict) -> list[str]:
    """The measured lines a design brief carries, from ``structure.json``."""
    slots = structure.get("slots") or []
    texts = [slot for slot in slots if slot.get("sample_text")]
    columns = len((structure.get("reading_order") or {}).get("columns") or [])
    budgets = [b for b in structure.get("title_budgets") or []
               if isinstance(b, dict) and b.get("chars_per_line")]
    facts = [f"- design number {structure.get('number')}, master "
             f"{structure.get('master')}",
             f"- {len(texts)} text slot(s) carry sample text, in {columns} "
             f"reading-order column(s)",
             f"- {len(structure.get('images') or [])} picture(s)"]
    if budgets:
        facts.append(f"- the title holds about {budgets[0]['chars_per_line']} "
                     f"characters on one line")
    return facts


def design_brief(image: str, structure: dict) -> str:
    return DESIGN_BRIEF.format(
        image=image, facts="\n".join(design_facts(structure)),
        types=", ".join(f'"{kind}"' for kind in CONTENT_TYPES))


def last_json(text: str):
    """The last JSON object in a model's reply, or ``None``.

    Last, not first: the briefs ask for prose first and the object at the
    end, and a model that quotes the schema while thinking would otherwise be
    read as its own answer.
    """
    decoder = json.JSONDecoder()
    found, pos = None, 0
    while (start := text.find("{", pos)) != -1:
        try:
            value, end = decoder.raw_decode(text, start)
        except ValueError:
            pos = start + 1
            continue
        if isinstance(value, dict):
            found = value
        # Past the whole object: its inner findings are not answers.
        pos = end
    return found


def verdict_line(answer: dict | None) -> str:
    """One verdict-table cell from a parsed slide answer."""
    if not answer:
        return ""
    findings = [f for f in answer.get("findings") or [] if isinstance(f, dict)]
    if str(answer.get("verdict", "")).lower() == "ok" and not findings:
        return "ok (vision-agent)"
    parts = []
    for finding in findings:
        name = SLIDE_QUESTIONS.get(str(finding.get("q", "")).upper(), "")
        what = str(finding.get("what") or "").strip()
        fix = str(finding.get("fix") or "").strip()
        parts.append(f"{name.lower() or 'finding'}: {what}"
                     + (f" → {fix}" if fix else ""))
    text = "fix — " + "; ".join(parts) if parts else "fix"
    return text.replace("|", "/").replace("\n", " ")
