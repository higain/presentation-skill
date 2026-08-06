#!/usr/bin/env python3
"""
Look at the rendered slides with a vision model and report what it sees.

    scripts/review_renders.py render/ --openrouter --model google/gemma-4-31b-it
    scripts/review_renders.py render/ --model gemma3:27b        # local ollama
    scripts/review_renders.py render/ --dry-run                 # show the request

This is the one check in this skill that has eyes.  Everything in ``qa.py`` is
arithmetic on the file: it finds text that overruns its box, shapes off the
trim, colours that cannot be read, wording nobody replaced.  None of it can
tell you a diagram is off-centre, that four cards are styled three different
ways, or that a slide reads as unfinished -- and those are the defects that
actually shipped in the runs this script was written for.

**Advisory by construction.**  It exits 0 whether it finds ten things or none,
because a vision model's false-positive rate on slide layout is unknown and a
check that cannot be satisfied is a check that gets ignored -- along with the
measured ones around it, which are right.  Treat what it says as a second pair
of eyes with no authority: read the finding, look at the slide, decide.

Stdlib only, and it talks to any OpenAI-compatible ``/chat/completions``
endpoint -- ollama, vLLM, llama.cpp, LM Studio, OpenRouter.  Same transport as
``catalog_gemma.py``, which has been driving a vision model in this skill for
as long as the catalog has existed.

**The model must accept images.**  A text-only model (deepseek-v4-flash, say)
returns a confident review of a slide it never saw, which is worse than no
review at all: check the model's input modalities before pointing this at it.

Keys are read from the environment (``OPENROUTER_API_KEY``, then
``OPENAI_API_KEY``) or ``--api-key``.  This script does not go looking through
any tool's credential store.

Exit codes: 0 reviewed (findings or not), 1 nothing could be reviewed, 2 bad
usage.
"""

import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

OPENROUTER_URL = "https://openrouter.ai/api/v1"

#: Env vars, so ``make_deck.py`` can run this without the build agent having to
#: choose to.  The failure this whole script exists for is a model that was
#: told to look at the renders and did not; putting the reviewer behind a flag
#: an agent has to remember reproduces it exactly.
MODEL_ENV = "PRESENTATION_REVIEW_MODEL"
BASE_URL_ENV = "PRESENTATION_REVIEW_BASE_URL"

PROMPT = """You are reviewing one rendered slide from a finished business \
presentation, as the audience will see it. You are the only reviewer with eyes.

Report only what you can SEE in the image. Do not guess at intent, do not \
comment on whether the wording is persuasive, and do not invent text you cannot \
read. If the slide looks fine, say so -- a clean slide is the expected case and \
an empty findings list is a good answer.

NEVER report spelling, typos, grammar or punctuation. The exact text of this \
deck is checked by a tool that reads the file rather than looking at pixels, so \
it already knows every character. A letter-level defect you think you see is \
far more likely to be your own misreading of the render, and reporting one \
costs the reader more than it saves. (Saying a word is in the wrong LANGUAGE, \
or is obviously leftover placeholder furniture, is not a spelling claim and is \
welcome.)

These are ALREADY checked by measurement, so do not report them unless the \
image shows something much worse than a borderline case: text overflowing its \
box, shapes hanging off the edge of the slide, low colour contrast, two text \
blocks printed over each other.

Report what measurement cannot see:
- a diagram, chart or figure that is off-centre, lopsided, or not aligned with \
the text it belongs to
- elements that should line up and do not -- columns of different widths, cards \
at different heights, labels at different baselines
- inconsistent styling between things that are peers: four cards where one has \
a different fill, a heading in a different size or colour than its siblings
- a large empty region that reads as unfinished, or content crammed into part \
of the slide while the rest is bare
- a figure that fails to communicate: a Venn whose circles do not overlap, \
arrows pointing nowhere, a legend for a chart that is not there
- text left in a different language from the rest of the slide
- anything else that would make a reader think the deck was not finished

Return ONLY a JSON object, no prose:
{"verdict": "ok" | "look",
 "findings": [{"what": "<the defect, concretely -- name the element>",
               "where": "<where on the slide, e.g. 'the three cards, lower \
half' or 'the Venn, right lobe'>",
               "severity": "high" | "medium" | "low"}]}

"ok" with an empty findings list if nothing is wrong. "high" means a reader \
notices immediately and the deck looks broken; "low" means you would mention it \
only if asked. At most %d findings, worst first -- if there are more, report \
the worst and say so in the last one."""


def _extract_json(text: str):
    """Pull the first JSON object out of a model reply, fenced or not."""
    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, re.S)
    if fenced:
        text = fenced.group(1)
    start = text.find("{")
    if start == -1:
        raise ValueError(f"no JSON in reply: {text[:200]!r}")
    end = text.rfind("}")
    return json.loads(text[start:end + 1])


def _call(base_url, model, api_key, messages, max_tokens=1500, timeout=180):
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0,
        "max_tokens": max_tokens,
    }
    req = urllib.request.Request(
        base_url.rstrip("/") + "/chat/completions",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json",
                 **({"Authorization": f"Bearer {api_key}"} if api_key else {})},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = json.loads(resp.read())
    return body["choices"][0]["message"]["content"]


def _messages(png: Path, max_findings: int):
    b64 = base64.b64encode(png.read_bytes()).decode()
    return [{
        "role": "user",
        "content": [
            {"type": "text", "text": PROMPT % max_findings},
            {"type": "image_url",
             "image_url": {"url": f"data:image/png;base64,{b64}"}},
        ],
    }]


def slide_number(png: Path):
    """The slide number in a ``slide-07.png``, or None."""
    hit = re.search(r"(\d+)", png.stem)
    return int(hit.group(1)) if hit else None


def find_renders(outdir: Path, wanted=None):
    """The per-slide PNGs in render order, contact sheets excluded."""
    pngs = [p for p in outdir.glob("*.png") if slide_number(p) is not None]
    pngs.sort(key=lambda p: slide_number(p))
    if wanted:
        pngs = [p for p in pngs if slide_number(p) in wanted]
    return pngs


def review_one(png: Path, args):
    """``(number, verdict, findings, error)`` for one slide."""
    number = slide_number(png)
    try:
        reply = _call(args.base_url, args.model, args.api_key,
                      _messages(png, args.max_findings), timeout=args.timeout)
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:200]
        return number, None, [], f"HTTP {e.code}: {detail}"
    except Exception as e:                       # network, timeout, bad body
        return number, None, [], f"{type(e).__name__}: {e}"

    try:
        parsed = _extract_json(reply)
    except Exception as e:
        return number, None, [], f"unparseable reply ({e})"

    findings = parsed.get("findings") or []
    if not isinstance(findings, list):
        findings = []
    # A model that answers "look" with nothing to look at, or lists findings
    # under an "ok" verdict, has contradicted itself.  The findings are the
    # substance; the verdict is derived from them rather than trusted.
    clean = []
    for f in findings:
        if len(clean) >= args.max_findings:      # cap real findings, not junk
            break
        if not isinstance(f, dict):
            continue
        what = str(f.get("what", "")).strip()
        if not what:
            continue
        clean.append({
            "what": what,
            "where": str(f.get("where", "")).strip(),
            "severity": (str(f.get("severity", "")).strip().lower()
                         if str(f.get("severity", "")).strip().lower()
                         in ("high", "medium", "low") else "medium"),
        })
    return number, ("look" if clean else "ok"), clean, None


RANK = {"high": 0, "medium": 1, "low": 2}

# Words that carry no signal about *which* defect a finding describes.
_STOP = frozenset("""the a an of in on at to and or as is it its that this with
while rest text slide""".split())

# How much of two findings' wording must overlap before they are the same
# defect seen twice.  Every slide is reviewed in its own call, so a model has no
# way to know it already reported the footer four slides ago -- it just reports
# it again, in different words each time.
#
# Measured on two real gemma-4-31b-it runs over the same ten-slide deck: the
# template's German footer produced four findings worded four ways in the first
# run and three in the second, and 0.4 folds each set into one line.  Below that
# it over-merges -- at 0.35 the second run's "the table is not centered" joined
# the Venn-diagram group, and the first run's two typo findings joined the
# footer group.  0.5 splits the footer wordings apart again.
SAME_DEFECT = 0.4


def _tokens(text):
    return {w for w in re.findall(r"[a-z]+", text.lower())
            if w not in _STOP and len(w) > 2}


def _jaccard(a, b):
    return len(a & b) / len(a | b) if (a or b) else 0.0


def cluster(findings):
    """Group ``(slide, finding)`` pairs that describe the same defect.

    Single-link and greedy, which is right for the shape of the input: a
    repeated defect is a tight blob of near-identical wordings, and there are
    never more than a few dozen findings to compare.

    Keyed on ``what`` *and* ``where``, because the two are what a model varies
    when it re-reports one defect.  A live run described the same template
    footer as "placeholder text in a different language" on two slides and
    "placeholder text in footer" on a third; on wording alone those share one
    content word out of four and printed as two lines.  Adding "bottom left
    footer" to both makes them one.
    """
    groups = []
    for number, f in findings:
        tokens = _tokens("%s %s" % (f["what"], f.get("where") or ""))
        for g in groups:
            if any(_jaccard(tokens, t) >= SAME_DEFECT for _, _, t in g):
                g.append((number, f, tokens))
                break
        else:
            groups.append([(number, f, tokens)])
    out = []
    for g in groups:
        # The longest wording is the most specific one, and the worst severity
        # anyone gave it is the one to lead with.
        best = max(g, key=lambda nft: len(nft[1]["what"]))[1]
        worst = min(RANK.get(f["severity"], 1) for _, f, _ in g)
        out.append((sorted({n for n, _, _ in g}), best, worst))
    return sorted(out, key=lambda row: (row[2], row[0][0]))


def review(pngs, args):
    """Review every PNG, in parallel, and return the rows in slide order."""
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        rows = list(pool.map(lambda p: review_one(p, args), pngs))
    return sorted(rows, key=lambda r: (r[0] is None, r[0]))


def as_markdown(rows, model):
    """The advisory section, ready to print or paste into a report."""
    findings = [(n, f) for n, _, fs, _ in rows for f in fs]
    errors = [(n, e) for n, _, _, e in rows if e]
    lines = []
    if findings:
        groups = cluster(findings)
        severity = {v: k for k, v in RANK.items()}
        lines += [f"## {len(groups)} thing(s) {model} noticed", ""]
        for numbers, f, worst in groups:
            where = f" ({f['where']})" if f["where"] else ""
            if len(numbers) == 1:
                head = f"slide {numbers[0]}"
                tail = ""
            else:
                head = "slides " + ", ".join(str(n) for n in numbers)
                # Something on this many slides is nearly always one thing --
                # a master footer, a logo, a label the template repeats.
                tail = " — same thing on each, so probably one fix"
            lines.append(
                f"- {head} [{severity[worst]}]: {f['what']}{where}{tail}")
        lines += ["",
                  "Advisory: a vision model looked at the renders and said "
                  "this. It is not authoritative and it does not fail the "
                  "build -- open the slide, decide, and ignore what is wrong."]
    else:
        reviewed = len([r for r in rows if r[3] is None])
        lines += [f"## {model} saw nothing wrong", "",
                  f"{reviewed} slide(s) reviewed, no findings. This is a "
                  f"second opinion, not a guarantee."]
    if errors:
        lines += ["", f"{len(errors)} slide(s) could not be reviewed:"]
        lines += [f"- slide {n}: {e}" for n, e in errors]
    return "\n".join(lines)


def configured(args=None):
    """The model to review with, from ``--model`` or the environment."""
    if args is not None and getattr(args, "model", None):
        return args.model
    return os.environ.get(MODEL_ENV) or None


def resolve_endpoint(model=None, base_url=None, api_key=None, openrouter=False):
    """``(base_url, api_key)`` from flags, then environment, then defaults."""
    if openrouter:
        base_url = OPENROUTER_URL
    base_url = base_url or os.environ.get(BASE_URL_ENV)
    api_key = api_key or (os.environ.get("OPENROUTER_API_KEY")
                          or os.environ.get("OPENAI_API_KEY"))
    if not base_url:
        # A provider-qualified id ("google/gemma-4-31b-it") is an OpenRouter
        # name and means nothing to ollama, which calls the same model
        # "gemma3:27b".  Guess from the shape of the id rather than making
        # every caller spell out an endpoint they already implied.
        base_url = (OPENROUTER_URL if model and "/" in model
                    else "http://localhost:11434/v1")
    return base_url, api_key


def main():
    ap = argparse.ArgumentParser(
        description="Review rendered slides with a vision model (advisory)")
    ap.add_argument("outdir", help="Directory of slide-NN.png renders")
    ap.add_argument("--model", default=None,
                    help=f"Vision model to ask. Falls back to ${MODEL_ENV}.")
    ap.add_argument("--base-url", default=None,
                    help=f"OpenAI-compatible endpoint. Falls back to "
                         f"${BASE_URL_ENV}, then OpenRouter for a "
                         f"provider/model id, else local ollama.")
    ap.add_argument("--api-key", default=None,
                    help="Bearer token. Falls back to OPENROUTER_API_KEY, "
                         "then OPENAI_API_KEY.")
    ap.add_argument("--openrouter", action="store_true",
                    help="Shorthand for --base-url " + OPENROUTER_URL)
    ap.add_argument("--slides", help="Only these slides, e.g. 4,7-9")
    ap.add_argument("--workers", type=int, default=4,
                    help="Slides in flight at once (default 4)")
    ap.add_argument("--timeout", type=int, default=180,
                    help="Seconds per slide (default 180)")
    ap.add_argument("--max-findings", type=int, default=4,
                    help="Cap per slide (default 4)")
    ap.add_argument("--json", action="store_true",
                    help="Emit the rows as JSON instead of markdown")
    ap.add_argument("--dry-run", action="store_true",
                    help="Write one request body and call nothing")
    args = ap.parse_args()

    outdir = Path(args.outdir)
    if not outdir.is_dir():
        ap.error(f"{outdir} is not a directory -- point this at render/")

    wanted = None
    if args.slides:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from office.impress import parse_slide_spec
        wanted = set(parse_slide_spec(args.slides))

    pngs = find_renders(outdir, wanted)
    if not pngs:
        print(f"No slide-NN.png renders in {outdir}", file=sys.stderr)
        return 1

    args.model = configured(args)
    if not args.model and not args.dry_run:
        ap.error(f"name a vision model: --model, or set ${MODEL_ENV}. "
                 f"It must accept images -- a text-only model will review a "
                 f"slide it cannot see.")

    args.base_url, args.api_key = resolve_endpoint(
        args.model, args.base_url, args.api_key, args.openrouter)
    if not args.api_key and "openrouter.ai" in args.base_url and not args.dry_run:
        ap.error("OpenRouter needs a key: set OPENROUTER_API_KEY or pass "
                 "--api-key")

    if args.dry_run:
        preview = json.loads(json.dumps(_messages(pngs[0], args.max_findings)))
        preview[0]["content"][1]["image_url"]["url"] = \
            "data:image/png;base64,<...>"
        print(json.dumps(preview, indent=2, ensure_ascii=False))
        print(f"\n[dry-run] {len(pngs)} slide(s) would go to "
              f"{args.model or '<no model>'} at {args.base_url}",
              file=sys.stderr)
        return 0

    rows = review(pngs, args)
    if all(e is not None for *_, e in rows):
        print(f"Nothing could be reviewed -- every call failed. First: "
              f"{rows[0][3]}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(
            [{"slide": n, "verdict": v, "findings": f, "error": e}
             for n, v, f, e in rows], indent=2, ensure_ascii=False))
    else:
        print(as_markdown(rows, args.model))
    return 0


if __name__ == "__main__":
    sys.exit(main())
