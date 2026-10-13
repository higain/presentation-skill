#!/usr/bin/env python3
"""
Split a long markdown source into one file per slide, without reading it.

    scripts/split_source.py Folienspezifikation.md --out spec/
    scripts/split_source.py outline.md --out spec/ --heading '^## Slide (\\d+)'

A slide-by-slide specification for a big deck runs to hundreds of kilobytes --
a 58-slide one was 288 KB, half of a 200k-token context before a line of code
was written.  The agent that read it whole lost it again to context pruning,
re-read pieces at random, and fell back to parsing it with regular
expressions, which silently dropped a field on every slide.  Split, each
slide's text is a few kilobytes that whoever builds that slide reads in full.

Writes into ``--out``:

* ``slide-07.md`` -- each slide's section, heading included, verbatim;
* ``_shared.md``  -- everything that belongs to no single slide (the brief, the
  design rules, an overview table), in order: read it along with any slide;
* ``index.md``    -- number, title and size of every slide.

The slide heading is found on its own: of all markdown headings that carry a
number, the most frequent kind (same level, same words before the number --
"### Folie 7: ...", "## Slide 7") is the slide heading.  ``--heading`` takes a
regular expression with one group for the number when that guesses wrong.
A section ends at the next heading of the same or a higher level.

Exit codes: 0 written, 1 no slide headings or a number used twice, 2 usage.
"""

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

HEADING_RE = re.compile(r"^(#{1,6})\s+(.*?)\s*$")
NUMBER_RE = re.compile(r"\d+")


def _headings(lines):
    """``[(line index, level, text)]`` for every markdown heading outside code."""
    found, fenced = [], False
    for i, line in enumerate(lines):
        if line.lstrip().startswith("```"):
            fenced = not fenced
            continue
        m = None if fenced else HEADING_RE.match(line)
        if m:
            found.append((i, len(m.group(1)), m.group(2)))
    return found


def _kind(text):
    """The words before a heading's first number: 'Folie 7: Titel' -> 'Folie'."""
    m = NUMBER_RE.search(text)
    return None if m is None else text[:m.start()].strip().lower()


def find_slides(lines, pattern=None):
    """``[(number, start, end, title)]`` -- each slide's line range, end exclusive."""
    heads = _headings(lines)
    if pattern:
        rx = re.compile(pattern)
        picked = []
        for i, level, text in heads:
            m = rx.match(lines[i])
            if m:
                picked.append((i, level, int(m.group(1)), text))
        # A custom pattern may match plain lines too, not only headings.
        if not picked:
            for i, line in enumerate(lines):
                m = rx.match(line)
                if m:
                    picked.append((i, 0, int(m.group(1)), line.strip()))
    else:
        numbers = {}
        for _, level, text in heads:
            if _kind(text) is not None:
                numbers.setdefault((level, _kind(text)), []).append(
                    int(NUMBER_RE.search(text).group()))
        if not numbers:
            return [], heads
        # Numbers that never repeat beat a bigger count ("1.1", "1.2" are
        # section numbers, not slides), and a word before the number ("Folie
        # 7") beats a bare one.
        (level, kind), found = max(
            numbers.items(),
            key=lambda item: (len(set(item[1])) == len(item[1]),
                              bool(item[0][1]), len(item[1])))
        if len(found) < 2:
            return [], heads
        picked = [(i, lvl, int(NUMBER_RE.search(text).group()), text)
                  for i, lvl, text in heads
                  if lvl == level and _kind(text) == kind]

    slides = []
    for i, level, number, text in picked:
        end = len(lines)
        for j, lvl, _ in heads:
            if j > i and (level == 0 or lvl <= level):
                end = j
                break
        slides.append((number, i, end, text))
    return slides, heads


def main():
    parser = argparse.ArgumentParser(
        description="Split a long markdown source into one file per slide")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out", type=Path, required=True,
                        help="Directory for slide-NN.md, _shared.md, index.md")
    parser.add_argument("--heading",
                        help="Regex for a slide heading line, one group for "
                             "the number, e.g. '^### Folie (\\d+)'")
    args = parser.parse_args()

    if not args.source.is_file():
        print(f"Error: {args.source} is not a file", file=sys.stderr)
        return 2
    try:
        text = args.source.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        print(f"Error: {args.source} is not UTF-8 text -- convert it to "
              f"markdown first", file=sys.stderr)
        return 2
    lines = text.splitlines(keepends=True)
    try:
        slides, _ = find_slides(lines, args.heading)
    except re.error as e:
        print(f"Error: --heading is not a regular expression: {e}",
              file=sys.stderr)
        return 2
    if not slides:
        print("Error: found no numbered slide headings. Name the heading "
              "with --heading '^## Slide (\\d+)'", file=sys.stderr)
        return 1

    numbers = Counter(n for n, *_ in slides)
    twice = sorted(n for n, c in numbers.items() if c > 1)
    if twice:
        print(f"Error: slide number(s) {', '.join(map(str, twice))} head "
              f"more than one section -- pass --heading to pick the right "
              f"headings", file=sys.stderr)
        return 1

    args.out.mkdir(parents=True, exist_ok=True)
    width = max(2, len(str(max(numbers))))
    covered = [False] * len(lines)
    rows = []
    for number, start, end, title in slides:
        for k in range(start, end):
            covered[k] = True
        body = "".join(lines[start:end]).rstrip() + "\n"
        path = args.out / f"slide-{number:0{width}d}.md"
        path.write_text(body, encoding="utf-8")
        rows.append((number, title, len(body.encode("utf-8")), path.name))

    shared = "".join(line for line, used in zip(lines, covered) if not used)
    shared = re.sub(r"\n{3,}", "\n\n", shared).strip() + "\n"
    (args.out / "_shared.md").write_text(shared, encoding="utf-8")

    missing = sorted(set(range(1, max(numbers) + 1)) - set(numbers))
    index = [f"# Slides in {args.source.name}", "",
             "Read `_shared.md` together with any slide file.", "",
             "| slide | file | KB | heading |", "|---|---|---|---|"]
    for number, title, size, name in sorted(rows):
        heading = title.replace("|", "\\|")
        index.append(f"| {number} | {name} | {size / 1024:.1f} | {heading} |")
    if missing:
        index += ["", "No section for slide(s) "
                  + ", ".join(map(str, missing)) + "."]
    (args.out / "index.md").write_text("\n".join(index) + "\n",
                                       encoding="utf-8")

    total = sum(size for _, _, size, _ in rows)
    print(f"{len(rows)} slides, {total / 1024:.0f} KB, largest "
          f"{max(size for _, _, size, _ in rows) / 1024:.1f} KB; shared part "
          f"{len(shared.encode('utf-8')) / 1024:.1f} KB")
    print(f"written to {args.out}/: slide-NN.md, _shared.md, index.md")
    if missing:
        print(f"warning: no section for slide(s) "
              f"{', '.join(map(str, missing))}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
