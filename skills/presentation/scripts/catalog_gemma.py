#!/usr/bin/env python3
"""
Drive a local vision model over the catalog/ that catalog_extract.py produced,
turning each layout's render + structure into a catalogue entry, then aggregating
the entries into a decision table.

Stdlib only -- talks to any OpenAI-compatible /chat/completions endpoint
(ollama, vLLM, llama.cpp server, LM Studio, OpenRouter...). Nothing needs pip.

    # 1. one entry per layout (sends render.png + structure.json to the model)
    python3 scripts/catalog_gemma.py catalog/ \
        --base-url http://localhost:11434/v1 --model gemma3:27b

    # 2. aggregate the entries into a decision table
    python3 scripts/catalog_gemma.py catalog/ --aggregate \
        --base-url http://localhost:11434/v1 --model gemma3:27b

    # inspect the exact request bodies without calling the model
    python3 scripts/catalog_gemma.py catalog/ --dry-run

No local GPU? --openrouter points the same passes at OpenRouter, reusing the
OPENROUTER_API_KEY the machine already has for it:

    OPENROUTER_API_KEY=... python3 scripts/catalog_gemma.py catalog/ \
        --openrouter --model google/gemma-4-31b-it

The key is only ever read from the environment (or --api-key); this script does
not go looking through any tool's credential store for it. Pick a model whose
input modalities include images -- the per-layout pass sends a PNG, so a
text-only model (glm-5.2, say) can run --aggregate but not the first pass.

Writes to:
    catalog/entries/slide-NN.json   one per layout (per-layout pass)
    catalog/decision_table.json     the aggregate (--aggregate pass)

Then assemble the catalog:
    python3 scripts/catalog_build.py templates/acme/catalog/ \
        -o templates/acme/catalog.md

The human-readable spec for these prompts is references/gemma-catalog-prompt.md;
the strings below mirror it. If you tune one, tune the other.
"""

import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

OPENROUTER_URL = "https://openrouter.ai/api/v1"

CONTENT_TYPES = (
    "title, agenda, section-divider, one-big-idea, bulleted-content, two-column, "
    "comparison, process-steps, timeline, metrics-kpi, quote, team-people, "
    "image-showcase, data-chart, table, closing, contact"
)

PER_LAYOUT = f"""You are cataloguing one slide layout from our company's PowerPoint \
template so an automated deck-builder knows when to use it. You are given a \
rendered image of the layout and a JSON description of its shapes (indices, \
positions in 1/100 mm, sizes, fonts, sample text).

Look at the image first -- the JSON confirms capacities and names, it does not \
override what you can see.

Return ONLY a JSON object, no prose, with exactly these fields:
{{"name": "<short human name, e.g. 'Section divider'>",
 "purpose": "<one sentence: what this slide is FOR>",
 "content_types": ["<tags from the fixed list>"],
 "capacity": {{"headline": <bool>, "body_slots": <int>,
   "max_bullets_per_body": <int, honest read of what fits without crowding>,
   "images": <int>, "supports_chart": <bool>}},
 "use_when": "<one sentence: the trigger to reach for this layout>",
 "avoid_when": "<one sentence: content that does NOT belong here>",
 "imagery": [{{"depicts": "<what the picture actually shows, e.g. 'a bird in \
flight', 'a city skyline'; never just 'an image'>",
   "role": "<decorative|photo-placeholder|icon|logo|background>"}}],
 "notes": "<quirks a builder must respect, or empty string>"}}

content_types must come from exactly this fixed vocabulary (pick all that apply):
{CONTENT_TYPES}.

Describe every picture you can see. For each photograph, illustration, icon or \
logo on the slide, name what it actually depicts ("a bird", "the company logo"), \
not "an image". One entry per picture in "imagery"; role is photo-placeholder if \
it looks like sample art meant to be swapped, decorative/background if it is \
fixed brand styling, icon/logo otherwise. Empty list [] if there are no \
pictures. The shape list will not tell you this -- only the rendered image does.

Be honest about capacity -- if a body box is sized for one line, say
max_bullets_per_body 1. Overstating capacity is what makes decks overflow. Do
not invent slots the layout does not have.

Do NOT produce retext keys, shape indices, geometry, slot counts or character
budgets. Those are measured from the file and are already in the catalog;
anything you write about them is guesswork that will be believed. If a field
above needs one, describe it in words instead."""

AGGREGATE = """Here are the catalogue entries for every design in our template, \
as a JSON array. Each entry's "slide" field is that design's number in the \
template deck -- that is the only identifier a deck-builder can act on. Ignore \
"layout": it is the master's autolayout id and is usually identical on every \
slide, so it names nothing. Produce two things.

1. A decision_table: an ordered list of rules mapping content to a template \
slide, most specific first. Each rule is {"when": "<content situation>", \
"use_slide": <int, the "slide" value of one of the entries above>, \
"because": "<short>"}. Every use_slide MUST be a slide number that appears in \
the entries -- never invent one, never emit a layout number. Cover at least: \
deck opener, agenda, each section start, a single headline idea, standard \
bulleted content, comparing two things, a process or timeline, numbers/KPIs, a \
chart, a quote, people, and the closing slide. If two slides serve the same \
purpose, pick one as default and name the other's number in because.

2. A variety_rule: one sentence on how to avoid every slide being \
title-and-bullets, naming which template slides to interleave.

Return ONLY JSON: {"decision_table": [...], "variety_rule": "..."}. No prose."""


def _extract_json(text: str):
    """Pull the first JSON value out of a model reply, fenced or not."""
    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, re.S)
    if fenced:
        text = fenced.group(1)
    start = min((i for i in (text.find("{"), text.find("[")) if i != -1),
                default=-1)
    if start == -1:
        raise ValueError(f"no JSON in model reply: {text[:200]!r}")
    end = max(text.rfind("}"), text.rfind("]"))
    return json.loads(text[start:end + 1])


def _call(base_url, model, api_key, messages, max_tokens=6000):
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
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            body = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        raise SystemExit(f"{base_url}: HTTP {e.code} -- {detail}") from e
    except (urllib.error.URLError, OSError) as e:
        raise SystemExit(
            f"cannot reach {base_url}: {e}\n  Is the model server running? "
            f"--base-url names it; the default is ollama's OpenAI endpoint "
            f"on this machine.") from e
    try:
        return body["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as e:
        raise SystemExit(f"{base_url}: unexpected reply shape: "
                         f"{json.dumps(body)[:300]}") from e


def _per_layout_messages(slide_dir: Path):
    structure = (slide_dir / "structure.json").read_text()
    png = slide_dir / "render.png"
    b64 = base64.b64encode(png.read_bytes()).decode()
    return [{
        "role": "user",
        "content": [
            {"type": "text",
             "text": PER_LAYOUT + "\n\nstructure.json:\n" + structure},
            {"type": "image_url",
             "image_url": {"url": f"data:image/png;base64,{b64}"}},
        ],
    }]


def per_layout(catalog: Path, args):
    manifest = json.loads((catalog / "manifest.json").read_text())
    out_dir = catalog / "entries"
    out_dir.mkdir(exist_ok=True)
    for row in manifest["slides"]:
        n = row["number"]
        slide_dir = catalog / f"slide-{n:02d}"
        messages = _per_layout_messages(slide_dir)
        if args.dry_run:
            # drop the base64 blob so the dump is readable
            preview = json.loads(json.dumps(messages))
            preview[0]["content"][1]["image_url"]["url"] = "data:image/png;base64,<...>"
            (out_dir / f"slide-{n:02d}.request.json").write_text(
                json.dumps(preview, indent=2, ensure_ascii=False))
            print(f"[dry-run] slide {n:02d}: request written")
            continue
        reply = _call(args.base_url, args.model, args.api_key, messages)
        entry = _extract_json(reply)
        # Identity is stamped from the manifest, never taken from the model.
        # "slide" is what a builder can act on: it indexes the template deck
        # (slides(doc)[slide - 1]). "layout" is the master's autolayout id and
        # is usually the same number on every slide of a designed template --
        # keep it as metadata, but it identifies nothing.
        entry = {"slide": n, "layout": row["layout"], "master": row["master"],
                 **{k: v for k, v in entry.items()
                    if k not in ("slide", "layout", "master")}}
        (out_dir / f"slide-{n:02d}.json").write_text(
            json.dumps(entry, indent=2, ensure_ascii=False))
        print(f"slide {n:02d}: {entry.get('name')}")
    if args.dry_run:
        print(f"\nDry run: request bodies in {out_dir}/*.request.json (no model called)")


def _entry_slide(path: Path, entry: dict):
    """The template slide number an entry describes, filename as fallback."""
    if isinstance(entry.get("slide"), int):
        return entry["slide"]
    try:
        return int(path.stem.split("-")[-1])
    except ValueError:
        return None


def aggregate(catalog: Path, args):
    entries_dir = catalog / "entries"
    entries = []
    for p in sorted(entries_dir.glob("slide-*.json")):
        # --dry-run leaves slide-NN.request.json next to the real entries, and
        # that matches slide-*.json too. Feeding a request body to the aggregate
        # pass as though it were a catalogue entry yields a decision table built
        # from the prompt rather than from the template.
        if p.name.endswith(".request.json"):
            continue
        entry = json.loads(p.read_text())
        entry.setdefault("slide", _entry_slide(p, entry))
        entries.append(entry)
    if not entries:
        sys.exit(f"no per-layout entries in {entries_dir} -- run without "
                 f"--aggregate first")
    messages = [{
        "role": "user",
        "content": AGGREGATE + "\n\nentries:\n" + json.dumps(entries, indent=2),
    }]
    if args.dry_run:
        (catalog / "decision_table.request.json").write_text(
            json.dumps(messages, indent=2, ensure_ascii=False))
        print(f"[dry-run] aggregate request -> {catalog}/decision_table.request.json")
        return
    reply = _call(args.base_url, args.model, args.api_key, messages, max_tokens=6000)
    table = _extract_json(reply)
    rules = table.get("decision_table", [])

    # A rule is only usable if it names a slide the template actually has.
    # Models reach for the layout id here (identical on every slide, so every
    # rule would point at the same design) -- catch that rather than ship it.
    known = {e["slide"]: e.get("name", "") for e in entries
             if isinstance(e.get("slide"), int)}
    unknown = []
    for rule in rules:
        n = rule.get("use_slide")
        if n in known:
            rule["slide_name"] = known[n]   # readable without cross-referencing
        else:
            unknown.append(n)

    (catalog / "decision_table.json").write_text(
        json.dumps(table, indent=2, ensure_ascii=False))
    print(f"decision_table.json written: {len(rules)} rules")
    if unknown:
        print(f"WARNING: {len(unknown)} rule(s) cite a slide the catalogue does "
              f"not have: {unknown} -- known slides are {sorted(known)}. "
              f"Re-run --aggregate or fix those rows by hand.")
    unused = sorted(set(known) - {r.get("use_slide") for r in rules})
    if unused:
        print(f"note: no rule reaches for slide(s) {unused}")


def main():
    ap = argparse.ArgumentParser(
        description="Author the layout catalogue with a local vision model")
    ap.add_argument("catalog", help="Directory from catalog_extract.py")
    ap.add_argument("--aggregate", action="store_true",
                    help="Fold catalog/entries/*.json into a decision table")
    ap.add_argument("--base-url", default="http://localhost:11434/v1",
                    help="OpenAI-compatible endpoint (default: ollama)")
    ap.add_argument("--model", default="gemma3:27b", help="Model name to request")
    ap.add_argument("--api-key", default=None,
                    help="Bearer token, if your endpoint needs one. Falls back "
                         "to OPENROUTER_API_KEY, then OPENAI_API_KEY.")
    ap.add_argument("--openrouter", action="store_true",
                    help="Shorthand for --base-url " + OPENROUTER_URL)
    ap.add_argument("--dry-run", action="store_true",
                    help="Write the request bodies and call nothing")
    args = ap.parse_args()

    if args.openrouter:
        args.base_url = OPENROUTER_URL
    # Keys come from the environment, never from another tool's credential file.
    if not args.api_key:
        args.api_key = (os.environ.get("OPENROUTER_API_KEY")
                        or os.environ.get("OPENAI_API_KEY"))
    if not args.api_key and "openrouter.ai" in args.base_url and not args.dry_run:
        ap.error("OpenRouter needs a key: set OPENROUTER_API_KEY or pass --api-key")

    catalog = Path(args.catalog)
    if not (catalog / "manifest.json").exists():
        ap.error(f"{catalog}/manifest.json not found -- run catalog_extract.py first")

    if args.aggregate:
        aggregate(catalog, args)
    else:
        per_layout(catalog, args)


if __name__ == "__main__":
    main()
