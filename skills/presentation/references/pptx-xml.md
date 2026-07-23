# Editing .pptx XML directly


Only for OOXML details UNO cannot express. A UNO round trip is lossless for
everything listed above; reach for raw XML when you need PowerPoint-specific
markup that Impress has no model for.

```python
import sys, zipfile; sys.path.insert(0, "scripts")
from pathlib import Path
from office.helpers import rezip, safe_extract

with zipfile.ZipFile("deck.pptx") as zf:
    safe_extract(zf, Path("unpacked"))       # rejects path-traversal entries
# ... edit unpacked/ppt/slides/slideN.xml ...
rezip(Path("unpacked"), Path("out.pptx"))
```

Use `rezip`, not the `zip` command — `zip` is not installed, and `rezip` writes
the member order and compression a `.pptx` needs.

- **Parse with `defusedxml.minidom`.** Round-tripping OOXML through
  `xml.etree.ElementTree` rewrites namespace prefixes and corrupts the deck.
- One `<a:p>` per list item — never concatenate items into one paragraph. Copy
  the sibling `<a:pPr>` to keep spacing.
- Let bullets inherit from the layout; add `<a:buChar>`, `<a:buAutoNum>` or
  `<a:buNone>` only to override. Never a literal `•` in the text.
- Text with leading or trailing spaces needs `xml:space="preserve"` on `<a:t>`.
- **Never reorder the children of `<p:presentation>`** — PowerPoint reads the
  order it wrote; moving an element makes the deck unopenable.
- Re-run `validate.py` afterwards.

