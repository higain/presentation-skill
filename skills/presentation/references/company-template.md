# delta — layout catalog

The house template this skill builds from when the user names no template of
their own. Generated from the real deck by `catalog_extract.py` + a vision
model (see *Provenance* at the foot), then reviewed.

**Asset:** `assets/company-template.pptx` — 21 designs, 16:9 (320 × 180 mm),
masters `AT_ST` (content), `Section_dark` (dividers), `Default 36` (slide 4).
Build with `clone_slide` + `retext` (see "Adapting a template deck" in SKILL.md).
QA every build with `qa.py --template assets/company-template.pptx`.

**Designs are named by their slide number in that asset.** Slide `N` is
`slides(doc)[N - 1]` — what you hand to `clone_slide`. Do *not* key on the
autolayout number `probe.py` reports as `layout`: every slide of this deck
reports `layout: 20`, so it distinguishes nothing.

## Decision table — content → template slide

Apply top to bottom; first match wins.

| When the content is… | Use slide | Design | Because |
|---|---|---|---|
| deck opener | `7` | Section Divider Dark | Dark section divider provides a high-impact start; slide 1 is a similar alternative. |
| agenda | `8` | Content slide with navigator | Content slide with navigator allows for a list of topics with a progress indicator. |
| section start | `1` | Section divider | Standard section divider for thematic shifts; slide 5 and 7 are alternatives. |
| single headline idea | `5` | Section Divider | Designed for a single centered title or short identifier. |
| standard bulleted content | `8` | Content slide with navigator | Standard content slide with a headline and a single body slot for bullets. |
| comparing two things | `12` | Two-column comparison with examples | Two-column comparison with examples; slide 6 is an alternative for analysis matrices. |
| process or timeline | `13` | Three-phase process flow | Three-phase process flow is the standard roadmap; slide 3, 10, 15, 16, 17, 18, 19, 21 offer specific variations. |
| numbers/KPIs | `4` | Project Timeline and Costing | Project Timeline and Costing is the only slide specifically for financial/effort metrics. |
| a chart | `11` | Phased Data Comparison | Phased Data Comparison is the only layout that explicitly supports a data chart. |
| a quote | `5` | Section Divider | Minimalist centered text slot is best for high-impact quotes. |
| people | `2` | Project Organization Chart | Project Organization Chart is designed for team roles and governance. |
| closing slide | `7` | Section Divider Dark | Dark section divider provides a clean, definitive end to the deck. |

**Variety rule:** Interleave standard bullet slides (8) with visual process flows (13, 18), analysis matrices (6, 12), and high-impact dividers (1, 7) to break the monotony of text.

## Design entries

### Slide 1 — Section divider
- **Purpose:** To introduce a new section or topic within the presentation.
- **Content types:** `section-divider`
- **Capacity:** headline ✓ · body slots 0 · max bullets 0 · images 0 · chart ✗
- **Use when:** A new chapter or major thematic shift in the deck is required.
- **Avoid when:** Detailed content, bullet points, or data visualizations are needed.
- **Imagery:** a large white outlined triangle (decorative); the delta company logo (logo)
- **Notes:** The layout is designed for a single centered title; avoid adding additional text boxes to maintain the minimalist aesthetic.

### Slide 2 — Project Organization Chart
- **Purpose:** To visualize the project governance structure, key roles, and communication flow.
- **Content types:** `bulleted-content`, `team-people`
- **Capacity:** headline ✓ · body slots 2 · max bullets 6 · images 2 · chart ✗
- **Use when:** When showing a project org chart with associated communication guidelines.
- **Avoid when:** When a simple list of team members or a complex multi-level corporate hierarchy is needed.
- **Imagery:** a speech bubble icon (icon); a yellow square placeholder for a customer logo (photo-placeholder); the delta company logo (logo)
- **Notes:** The left side is a fixed-structure org chart with 6 specific role boxes; the right side is for communication bullet points.

### Slide 3 — Process Flow with Logos
- **Purpose:** To visualize a linear sequence of steps or milestones involving partner logos.
- **Content types:** `process-steps`
- **Capacity:** headline ✓ · body slots 4 · max bullets 2 · images 4 · chart ✗
- **Use when:** Use when showing a 4-step chronological progression or a roadmap where each step is associated with a specific entity or logo.
- **Avoid when:** Avoid for complex non-linear processes, detailed data tables, or lists with more than four steps.
- **Imagery:** the delta company logo (logo); placeholder text 'Kundenlogo' (photo-placeholder)
- **Notes:** The layout uses a specific visual flow of interlocking circles and arrows; text must be kept concise to avoid overlapping the graphic elements.

### Slide 4 — Project Timeline and Costing
- **Purpose:** To present a detailed project schedule with associated effort (BT) and total cost calculations.
- **Content types:** `timeline`, `metrics-kpi`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 4 · images 0 · chart ✗
- **Use when:** When providing a detailed breakdown of project phases, milestones, and the resulting financial quote.
- **Avoid when:** When a high-level roadmap is needed without granular task-level effort or pricing.
- **Imagery:** none
- **Notes:** The central table is a manual construction of shapes; the builder must ensure the total cost box (bottom right) is updated in sync with the table totals.

### Slide 5 — Section Divider
- **Purpose:** To signal the start of a new section or chapter within the presentation.
- **Content types:** `section-divider`, `one-big-idea`
- **Capacity:** headline ✗ · body slots 1 · max bullets 1 · images 0 · chart ✗
- **Use when:** A new high-level topic is introduced that requires a visual break.
- **Avoid when:** Detailed content, lists, or data-heavy information are needed.
- **Imagery:** a white outlined stylized triangle (decorative); the delta company logo (logo)
- **Notes:** The text slot is centered within the decorative triangle and is intended for short identifiers or section titles only.

### Slide 6 — Analysis Matrix
- **Purpose:** To present two parallel analysis streams (e.g., competitor vs. target group) with their respective findings.
- **Content types:** `bulleted-content`, `comparison`
- **Capacity:** headline ✓ · body slots 6 · max bullets 3 · images 2 · chart ✗
- **Use when:** When you need to show a structured breakdown of two different analysis categories and their resulting outcomes.
- **Avoid when:** When presenting a linear process, a single deep-dive analysis, or data-heavy charts.
- **Imagery:** a laptop showing a chart on screen (photo-placeholder); a hand holding a pen writing on a document (photo-placeholder); crossed swords icon (icon); a target/bullseye icon (icon); a speech bubble icon (icon); a document/checklist icon (icon); the delta company logo (logo)
- **Notes:** The layout uses a mirrored structure: left-side context, center image, and right-side detailed findings with icons.

### Slide 7 — Section Divider Dark
- **Purpose:** To introduce a new section or chapter within the presentation.
- **Content types:** `section-divider`
- **Capacity:** headline ✓ · body slots 0 · max bullets 0 · images 0 · chart ✗
- **Use when:** A new high-level topic or section is starting and requires a visual break.
- **Avoid when:** Detailed content, bullet points, or data-heavy information are needed.
- **Imagery:** a large white outlined triangle (decorative); the delta company logo (logo)
- **Notes:** The layout features a dark background with a single centered text slot for the section title.

### Slide 8 — Content slide with navigator
- **Purpose:** To present detailed text content while indicating the current progress or phase within a larger process.
- **Content types:** `bulleted-content`, `one-big-idea`
- **Capacity:** headline ✓ · body slots 1 · max bullets 6 · images 0 · chart ✗
- **Use when:** Use when providing detailed explanations or instructions that are part of a multi-step sequence.
- **Avoid when:** Avoid when the slide requires heavy data visualization, multiple columns, or large imagery.
- **Imagery:** a sequence of five chevron arrows (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout includes a specific area for a progress navigator (chevrons) in the top right.

### Slide 9 — Navigator Showcase
- **Purpose:** To demonstrate or explain different navigation elements and frameworks used throughout a presentation.
- **Content types:** `bulleted-content`, `image-showcase`
- **Capacity:** headline ✓ · body slots 1 · max bullets 3 · images 6 · chart ✗
- **Use when:** When explaining the presentation's navigation logic or showcasing specific framework icons.
- **Avoid when:** When presenting primary data, complex tables, or long-form text documents.
- **Imagery:** a grey puzzle piece (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a sequence of four document icons with a grey arrow (photo-placeholder); a simplified Business Model Canvas diagram with blue and green blocks (photo-placeholder); a chevron process arrow sequence in grey and green (decorative); the delta company logo (logo)
- **Notes:** The layout is designed for a mix of explanatory text on the left and a gallery of visual examples on the right.

### Slide 10 — Process Flow with Detail
- **Purpose:** To visualize a linear progression of phases with a detailed breakdown of a specific step or result below.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 3 · images 0 · chart ✗
- **Use when:** Use when showing a high-level 3-step process and providing a detailed explanation or sub-process for one of those steps.
- **Avoid when:** Avoid for complex timelines with many dates or slides requiring heavy data tables.
- **Imagery:** a grey right-pointing arrow (decorative); a grey right-pointing arrow (decorative); a sequence of connected blue hexagonal shapes (photo-placeholder); an orange slanted rectangle label (decorative); an orange slanted rectangle label (decorative); the delta company logo (logo)
- **Notes:** The layout contains a fixed graphic of connected hexagons that serves as a visual placeholder for a sub-process or detailed breakdown.

### Slide 11 — Phased Data Comparison
- **Purpose:** To show a progression of phases alongside a corresponding data chart and explanatory notes.
- **Content types:** `process-steps`, `data-chart`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 5 · max bullets 2 · images 0 · chart ✓
- **Use when:** Use when you need to link specific project phases or steps to a quantitative chart result.
- **Avoid when:** Avoid for simple lists or complex tables that require full-slide width.
- **Imagery:** a hand holding a coin with a dollar sign (icon); a magnifying glass (icon); a calendar with a growth chart line (icon); the delta company logo (logo)
- **Notes:** The layout contains three fixed phase labels on the left and a chart area on the right; ensure text in 'Erklärung' (explanation) remains brief to avoid overlapping the chart.

### Slide 12 — Two-column comparison with examples
- **Purpose:** To compare two parallel workstreams or categories, each supported by a visual example and a list of detailed activities.
- **Content types:** `two-column`, `comparison`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 4 · images 2 · chart ✗
- **Use when:** When you need to contrast two distinct processes or deliverables that both require a visual aid and a supporting list of bullet points.
- **Avoid when:** When there is only one main topic or when the content requires a complex data chart instead of a simple screenshot/diagram.
- **Imagery:** a person icon (icon); a flowchart/process icon (icon); a document layout showing a role description (photo-placeholder); a process flow diagram with blue boxes (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout includes specific 'Beispiel' (Example) labels over the image areas that should be maintained or updated.

### Slide 13 — Three-phase process flow
- **Purpose:** To visualize a linear project timeline or process divided into three distinct phases with associated tasks.
- **Content types:** `process-steps`, `timeline`
- **Capacity:** headline ✓ · body slots 3 · max bullets 3 · images 0 · chart ✗
- **Use when:** Use when presenting a high-level project roadmap or a three-step methodology with detailed action items for each phase.
- **Avoid when:** Avoid for complex processes with more than three steps or for content that does not follow a linear chronological sequence.
- **Imagery:** a target with an arrow (icon); a magnifying glass (icon); a checklist on a clipboard (icon); a red diamond (decorative); the company logo (logo)
- **Notes:** The layout uses a chevron-style header for phases; ensure phase titles are concise to avoid text overflow in the blue shapes.

### Slide 14 — Five-pillar analysis
- **Purpose:** To present five distinct dimensions or categories of a strategic analysis surrounding a central core concept.
- **Content types:** `bulleted-content`, `process-steps`
- **Capacity:** headline ✓ · body slots 5 · max bullets 3 · images 0 · chart ✗
- **Use when:** Use when you need to break down a central topic into five specific analytical dimensions or question sets.
- **Avoid when:** Avoid when you have more or fewer than five key points, or when the content requires a linear sequence rather than a hub-and-spoke relationship.
- **Imagery:** four arrows pointing outwards (icon); a bar chart on a screen (icon); three small people figures (icon); a speech bubble with a person silhouette (icon); a bar chart (icon); the delta company logo (logo)
- **Notes:** The layout is strictly balanced around a central circle; ensure text in the five surrounding boxes is concise to avoid overlapping the center element.

### Slide 15 — Process Steps with Details
- **Purpose:** To visualize a sequential process or list of steps where specific steps are highlighted with detailed explanatory text.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 7 · max bullets 3 · images 0 · chart ✗
- **Use when:** Use when presenting a multi-step process where one or more steps require expanded descriptions or examples.
- **Avoid when:** Avoid for simple lists without detailed call-outs or for complex timelines with specific dates.
- **Imagery:** a pie chart icon (icon); a magnifying glass over a document icon (icon); a list with a magnifying glass icon (icon); a human head profile icon (icon); a bar chart with an upward arrow icon (icon); the delta company logo (logo)
- **Notes:** The layout uses a bracket-style visual connector to link specific process steps to their corresponding detail boxes on the right.

### Slide 16 — Process List
- **Purpose:** To outline a sequential process or a series of steps with corresponding detailed descriptions.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 5 · max bullets 3 · images 0 · chart ✗
- **Use when:** Use when presenting a linear workflow or a list of five key phases that require brief explanatory bullets.
- **Avoid when:** Avoid for complex diagrams, large data tables, or processes with more than five steps.
- **Imagery:** a data table (photo-placeholder); a funnel (photo-placeholder); a bar chart with an upward arrow (photo-placeholder); a funnel (photo-placeholder); a document with charts (photo-placeholder); a blue vertical arrow shape (decorative); the delta company logo (logo)
- **Notes:** The layout features a central column of placeholder graphics that should be replaced with icons or visuals representing each process step.

### Slide 17 — Iterative Process Flow
- **Purpose:** To visualize a cyclical development process that transforms initial specifications into a final model through iterative feedback loops.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 5 · images 0 · chart ✗
- **Use when:** Use when showing a non-linear workflow where inputs are refined through a central iterative cycle to reach a set of outcomes.
- **Avoid when:** Avoid for simple linear timelines or slides requiring heavy data visualization.
- **Imagery:** a circular arrow diagram representing a cycle (decorative); the company logo 'delta' (logo)
- **Notes:** The central cycle contains three fixed text areas; the left and right columns are for bulleted lists of inputs and outputs respectively.

### Slide 18 — Process flow with images
- **Purpose:** To visualize a linear multi-step process where each phase is supported by a representative image and a descriptive label.
- **Content types:** `process-steps`, `image-showcase`
- **Capacity:** headline ✓ · body slots 8 · max bullets 1 · images 4 · chart ✗
- **Use when:** You need to show a 4-step sequence where each step requires a visual aid and a short caption.
- **Avoid when:** The process has more or fewer than four steps, or requires detailed bulleted explanations for each phase.
- **Imagery:** a pie chart and line graph icon (icon); a compass icon (icon); a lightbulb icon (icon); a strategy map with X and O markers icon (icon); a magnifying glass over documents (photo-placeholder); a hand drawing arrows on a blackboard (photo-placeholder); people looking at maps or blueprints on a table (photo-placeholder); a movie clapperboard (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout is strictly locked to 4 columns; the chevron headers and bottom labels must be kept concise to avoid text overflow.

### Slide 19 — Three-column process analysis
- **Purpose:** To present three parallel analysis streams, each with a visual representation, descriptive bullets, and a resulting outcome.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 2 · images 3 · chart ✗
- **Use when:** When comparing three distinct strategic modules or analysis steps that lead to specific foundations or conclusions.
- **Avoid when:** When there are more or fewer than three categories, or when long-form text is required.
- **Imagery:** a colored grid matrix (photo-placeholder); a value chain diagram with arrows (photo-placeholder); a Business Model Canvas diagram (photo-placeholder); three blue chevron arrows (decorative); the delta company logo (logo)
- **Notes:** The layout relies on a strict three-column symmetry; ensure text in the bottom 'Grundlage' slots is concise to avoid overlapping the blue accent lines.

### Slide 20 — Two-column process detail
- **Purpose:** To explain a specific module or process step by splitting the approach and the sources into two distinct lists.
- **Content types:** `bulleted-content`, `two-column`
- **Capacity:** headline ✓ · body slots 2 · max bullets 5 · images 1 · chart ✗
- **Use when:** When you need to present a 'how-to' (Vorgehen) alongside the 'where from' (Recherchequellen) for a specific project phase.
- **Avoid when:** When you have a single long list of items or need to present complex data charts.
- **Imagery:** a person wearing a VR headset seen through a magnifying glass (photo-placeholder); a gear and document icon (icon); a book/library icon (icon); three blue chevron arrows (decorative); the delta company logo (logo)
- **Notes:** The layout features a central circular image placeholder that overlaps the two text columns; ensure text does not wrap into the image area.

### Slide 21 — Process Diagram with Detail
- **Purpose:** To visualize a non-linear process or conceptual framework with a detailed explanation of one specific component.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 5 · images 0 · chart ✗
- **Use when:** When showing a conceptual flow or a set of interconnected methods where one specific step requires a detailed bulleted list.
- **Avoid when:** When a simple linear timeline or a standard text-heavy slide is needed.
- **Imagery:** a blue chevron arrow (decorative); a document icon (icon); the delta company logo (logo)
- **Notes:** The layout uses a custom diagram of interlocking polygons; the builder must ensure text fits within these specific geometric shapes.
## Working with this asset

Facts established by inspecting the file itself (`probe.py`, zip inspection),
not by reading the renders. A vision model cannot see any of these.

### think-cell is embedded throughout

17 of the 21 slides carry a hidden OLE shape literally named
`think-cell data - do not delete`; the deck holds 62 `oleObject*.bin`
embeddings in total. **Do not delete these shapes**, and expect them not to
survive a UNO round trip intact — a rebuilt deck loses the think-cell linkage
even when the visible slide looks identical. If a build must keep live
think-cell charts, edit the original in PowerPoint instead of rebuilding here.

Slide 11 is the exception worth knowing: its waterfall is a genuine OOXML
`barChart` (`ppt/charts/chart1.xml`) with an embedded workbook — two series,
the lower one invisible to fake the waterfall — so it *is* editable as a real
chart. It carries no think-cell object.

### Cloning dividers — fixed, but know what it cost

Cloning a `Section_dark` divider used to produce a white slide with black text.
Two separate faults, both now fixed in `office/template.py`:

1. **The master was not carried.** A new page comes up on its neighbour's
   master, so the divider landed on `AT_ST` and lost the dark background.
   `clone_slide` now assigns `src_page.MasterPage` before copying anything.
2. **A cloned placeholder loses its direct run colour on export.** This one is
   the subtle half. The write lands — read the portion back in the same session
   and it is white — but LibreOffice's `.pptx` export discards direct run
   formatting on a *placeholder* and re-inherits the master's text style, so
   the run leaves as `<a:schemeClr val="dk1"/>` where the source had `lt1`.
   Reassigning the master, rebinding `Style` to the source master's family, and
   forcing a non-theme colour (`0xFEFEFE`) all failed to prevent it.

`clone_slide` now rebuilds such placeholders as plain `drawing.TextShape`s,
which have no style to re-inherit from, so the colour exports as an explicit
`<a:srgbClr val="ffffff"/>`. The trigger is narrow: a placeholder is flattened
only when a run's colour differs from its own presentation style's. Across this
template that is the divider titles and the blue topic label on every content
slide — the black action titles keep their placeholder.

**The blue topic label was silently going black too**, on all 18 content
slides; the same fix repairs it.

Verified against the renders: clones of slides 1, 5 and 7 now match their
sources exactly (identical near-white pixel counts, RMS ≈ 3, antialiasing
only), and slides 13 and 20 are byte-for-byte unchanged from before the fix.

The cost of flattening: those shapes are no longer placeholders, so they will
not pick up later layout changes, and per-run formatting *within* a paragraph
collapses to that paragraph's first run.

### Still unfixed: inline formatting in body text boxes

Separate from the above and **not** fixed. The body text on these slides sits
in ordinary `ooxml-rect` shapes, not placeholders, so `clone_shape` handles it
— and it loses bullet markers and inline run colour, flattening a paragraph
that mixed blue and black to a single colour. Compare a clone of slide 20
against the original and the bullets are visibly plainer. Budget for fixing
body text by hand after a build, or check the render before shipping.

### Round-trip validation noise

`validate.py` reports the untouched asset clean, but a UNO round trip of it
raises 22 findings: 6 broken `customXml/item*.xml` relationships (reported
CRITICAL) and 16 undeclared `.wmf` content types. The resulting file opens and
renders correctly regardless, so treat these as round-trip noise on this
template rather than as build failures — but they will mask any *real*
validator finding, so compare against a baseline run rather than expecting zero.

### The AT_ST text grid

Every `AT_ST` content slide (2–4, 6, 8–21) repeats the same three text slots:

| Slot | Position | Size | Role |
|---|---|---|---|
| Action title | y = 498 | 16 pt, black | The slide's **assertion** — a full sentence, not a label |
| Topic label | y = 2302 | 16 pt, blue | Short chapter/section name |
| Source line | y = 16003 | 14 pt | Empty footnote/source line, usually left blank |

The action title carries the argument; German consulting decks read as a chain
of these sentences. When retexting, keep it a full assertive sentence.

### Authoring annotations that must be deleted

The deck is a pattern collection, so several slides carry rotated coloured
flags that are instructions to the author, never content:

- Orange/green/magenta `Beispiel`, `Layout`, `Alternatives Design zum Master`
  flags — slides 6, 10, 11, 12, 15.
- Slide 4's yellow `Vorgehen:` block (four numbered build steps) and its orange
  toolbox tip box.
- Yellow highlight marks `Kundenlogo`, `Kunde`, `XX.XXX`, `XX`, `450` — fill-in
  markers, replace with real values or the customer's logo.

Delete these on use. `qa.py` will not flag them; they are ordinary shapes.

### Gaps in the collection

**There is no cover slide and no closing/contact slide.** Slides 1, 5 and 7 are
all section dividers on `Section_dark`. The decision table's "deck opener" and
"closing slide" rows therefore both point at a divider — workable, but if a
brief needs a real title slide or a contact page, build one rather than
assuming the template has one.

**Slides 8 and 9 are documentation, not designs.** They explain how the
Navigator element works and what it is for. The decision table reaches for
slide 8 as the default bulleted-content and agenda layout; that works because
its body is a plain text box, but the sample text on it is meta-commentary
about the template and must be fully replaced.

## Provenance

Generated 2026-07-28 from `delta-sammlung-21f.pptx` (21 slides):

1. `scripts/catalog_extract.py` — per-slide `render.png` + `structure.json`,
   no model involved.
2. `scripts/catalog_gemma.py --openrouter --model google/gemma-4-31b-it` —
   per-layout pass, then the aggregation pass for the decision table.

To regenerate after the template changes, repeat both passes (see *Maintaining
the default* in SKILL.md) and re-review every row against the renders.
