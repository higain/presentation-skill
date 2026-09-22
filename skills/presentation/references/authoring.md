# Building and editing decks without a template

For a deck on the house template, follow the recipe in SKILL.md — this file is
the layer underneath it: making a document from nothing, editing one that
already exists, and the typography and environment facts that apply either way.

## Building a deck

```python
import sys; sys.path.insert(0, "scripts")
from office.impress import (
    new_document, add_slide, add_textbox, add_image, placeholder,
    set_notes, save_as, slides, mm, pt,
)

doc = new_document()
page = slides(doc)[0]                      # a new document ALREADY has one slide
page.Layout = 0                            # 0 = title + subtitle
placeholder(page, "Title").setString("Quartalsbericht Q3")
placeholder(page, "Subtitle").setString("Vertrieb & Marketing")
set_notes(page, "Mit der Umsatzzahl aufmachen.")

page = add_slide(doc)                      # blank; place everything yourself
box = add_textbox(doc, page, mm(20), mm(20), mm(240), mm(25), "Ergebnisse",
                  size=40, weight=150, color=0x1E2761)
add_image(doc, page, "chart.png", mm(20), mm(60), mm(120), mm(80))

save_as(doc, "deck.pptx")                  # the extension picks the filter
doc.close(False)
```

**Geometry is in 1/100 mm, not points or EMU.** Use `mm()` and `pt()` rather
than raw numbers — a value that looks like points is 25× too small and lands
off-slide with no error at all.

**Read the slide size; never hardcode it.** `slide_size(doc)` is the only answer
that is right for the deck in front of you. A LibreOffice default page is
28000 × 15750, but a brand template is usually its own size — the house
template is **32001 × 18000**, 14% wider, enough to shift every offset you
computed against the default. `D.page_box(page)` gives you the whole slide as a
`Box`.

**Style after the text, never before.** `setString()` on a shape re-inherits the
drawing default (Arial 18, black) and silently discards character formatting set
before it — this is why font sizes "don't stick". Passing `size=`/`color=`/
`weight=` to `add_textbox` gets the order right for you; if you call
`setString()` yourself, follow it with `office.draw.restyle(shape, ...)`.

`add_slide(doc)` gives you a blank slide. It takes an optional `layout=` —
Impress *autolayout* numbers, `0` title+subtitle, `1` title+content, `19` title
only, `20` blank (the default) — which have **nothing to do with the design
numbers in a layout catalog.** A catalog's design `20` is the template's
twentieth slide, whatever it holds; `layout=20` is "no placeholders". Confusing
the two puts a densely filled template slide where you wanted an empty one, and
the numbers collide often enough to make it a real mistake rather than a
theoretical one.

**`new_document()` already contains one slide.** Use `slides(doc)[0]` for the
title slide, as above — calling `add_slide()` for it leaves a stray blank slide
at the front of every deck. A document must also always keep at least one
slide: deleting the last one wedges LibreOffice, so `delete_slide()` refuses.

**Fill placeholders rather than adding text boxes** wherever a layout offers
one. Text in a placeholder inherits the master's fonts, sizes and position; a
free text box inherits nothing, and that is how decks end up looking unstyled.

## Editing an existing deck

```python
from office.impress import impress_document, slides, placeholder, set_notes

with impress_document("deck.odp") as doc:      # saved in place, still .odp
    page = slides(doc)[0]
    placeholder(page, "Title").setString("Neuer Titel")
    set_notes(page, "Angepasste Notiz.")
```

- `impress_document(path, save=False)` opens read-only — use it for inspection.
- An exception inside the block leaves the file on disk untouched.
- **`store()` is a no-op unless the document is modified.** `impress_document`
  handles that; if you call `doc.store()` yourself, call `doc.setModified(True)`
  first or nothing is written.
- `save_as(doc, path)` writes elsewhere and leaves the input alone, so a
  load → edit → `save_as` chain never overwrites the source.

## Reading a deck

```bash
scripts/dump.py deck.odp
```

```
## Slide 1
*layout 0, master «Candy»*

**title:** Quartalsbericht Q3
*empty slots: subtitle*

**Notes:**
> Mit der Umsatzzahl aufmachen.
```

`empty slots` is what makes a template readable — it names the placeholders a
layout offers to fill. Text is listed in z-order (the order shapes were added),
not top-to-bottom reading order.

## Typography

**Font names you write are rendered by the recipient's Office, not by this
machine.** QA renders through LibreOffice, which substitutes fonts it lacks. If
the substitute has different metrics, your overflow check is wrong.

Measured on Debian with the `fonts-crosextra-*` packages installed:

| You write | Renders here as | Overflow QA |
|---|---|---|
| Arial | Liberation Sans | trustworthy (metric-compatible) |
| Times New Roman | Liberation Serif | trustworthy |
| Courier New | Liberation Mono | trustworthy |
| Calibri | Carlito | trustworthy |
| Cambria | Caladea | trustworthy |
| Georgia, Trebuchet MS, Impact, Arial Black, Garamond, Consolas, Palatino | DejaVu family | **not trustworthy** — different widths |
| Aptos | DejaVu Sans | **avoid** — no metric-compatible substitute anywhere, and missing from older Office installs |

Calibri and Cambria are trustworthy only when `fonts-crosextra-carlito` and
`fonts-crosextra-caladea` are installed (both from apt). If `fc-match Calibri`
returns anything but Carlito, treat them as untrustworthy again.

For fonts outside the trustworthy list: use them where asked, size those
containers with ~10% slack, and do not trust the preview's text fit.

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

## Editing .pptx XML directly

Only for OOXML details UNO cannot express — a UNO round trip is lossless for
everything in `references/qa.md`. When you do need it, **`references/pptx-xml.md`**
has the unpack/rezip recipe and the rules that keep a deck openable.

## Where temporary files go

Every script here needs scratch space: a disposable LibreOffice profile per run,
a PDF between a deck and its PNGs, an unzipped `.pptx` for the validator.
Importing anything from `office` picks a writable directory once and points both
Python and LibreOffice at it — you do not normally have to think about it.

**If `/tmp` is not writable** — a locked-down or containerised machine — it falls
back to `.presentation-tmp/` in the working directory and removes it on exit.
Without that, LibreOffice aborts with *"User installation could not be
completed"* and converts nothing, so every render and every QA run fails at once
with no obvious cause. To put the scratch space somewhere specific:

```bash
PRESENTATION_TMPDIR=./build/tmp scripts/make_deck.py build.py out.pptx
```

That path must be writable; if it is not, the run stops immediately and says so,
rather than failing later inside LibreOffice.

## Notes on this environment

- **apt only.** Never run `pip` or `npm`, in a script or in a shell. Everything
  needed comes from apt (`libreoffice-impress python3-uno python3-lxml
  python3-defusedxml python3-pil poppler-utils`); if something is missing, name
  the package and stop.
- **`from com.sun.star... import ...` only works after `import uno`.** The
  `com` package is synthesised by the UNO bridge, so put those imports *below*
  your `office.*` imports or you get `ModuleNotFoundError: No module named 'com'`.
  Same for `createInstance`: services like `drawing.ShapeCollection` come from
  the remote `ServiceManager`, not from `doc`.
- The machine's locale decides how LibreOffice formats numbers and dates in the
  renders (`de_DE` gives `1.234,50`, `en_US` gives `1,234.50`). Content in any
  language is normal; write the deck in the language of the brief.
- Macros are never used here; documents open with macro execution disabled.
- Each script starts its own headless LibreOffice with a private profile and
  port, so parallel runs — and a LibreOffice the user has open — never collide.
