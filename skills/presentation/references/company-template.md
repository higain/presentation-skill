# delta — layout catalog

The house template this skill builds from when the user names no template of
their own. Generated from the real deck by `catalog_extract.py` + a vision
model (see *Provenance* at the foot), then reviewed.

**Asset:** `assets/company-template.pptx` — 21 designs, 16:9, **32001 × 18000**
in 1/100 mm (*not* the LibreOffice default 28000 × 15750 — read
`slide_size(doc)` rather than assuming, or every offset you compute lands 14%
too far left). Masters `AT_ST` (content), `Section_dark` (dividers),
`Default 36` (design 4).

**Build with `build_deck`** — see "Building from a template at full fidelity" in
SKILL.md. It duplicates rather than rebuilds, so this template's bullets, inline
colour and think-cell objects survive; `clone_slide` loses all three here. Check
every build with
`make_deck.py build.py out.pptx --template assets/company-template.pptx`, and
open the renders it prints.

**Designs are named by their slide number in that asset.** Design `N` is
`slides(doc)[N - 1]`, and `N` is what you put in `build_deck`'s design list. Do
*not* key on the autolayout number `probe.py` reports as `layout`: every slide of
this deck reports `layout: 20`, so it distinguishes nothing — and **catalog
design 20 is a real, densely filled two-column slide, not a blank one.**

**Much of this template's text lives inside groups.** Eleven of the 21 designs
carry text in grouped shapes, which `shapes(page)` cannot reach — design 10 has
16 grouped text shapes against 11 top-level. `retext`, `find_shapes` and `qa.py`
all walk groups, and `probe.py` lists them indented as `[6.2]`. Each entry below
notes its count. Retext by wording, never by index.

## Decision table — content → template slide

Apply top to bottom; first match wins.

| When the content is… | Use slide | Design | Because |
|---|---|---|---|
| **a diagram, chart, matrix, comparison — anything no row below fits** | `8` | **Canvas** | Strip design 8 back with `draw.canvas(page, keep=(0, 1, 2, 3))` — think-cell stub, title, topic label, source line — and draw in the area the body text vacates: **2031, 3545 to 30003, 15702**. You keep the master, background, chrome, fonts and margins. **Reach for this rather than forcing the content into the nearest bullet list** — see "The canvas pattern" in SKILL.md. |
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

**The divider titles hold exactly one line.** Designs 1, 5 and 7 set their
title in the gap between the two edges of a big white triangle, at 60pt. One
line lands in the gap; a **second line crosses the triangle**, and the words
come out with a stroke drawn through them. Nothing on the page collides, so
`qa.py` cannot see it — the triangle is master decoration, not a shape on the
slide. Budgets at 60pt, measured:

| Design | Title box width | One line holds |
|---|---|---|
| `1` | 18188 | ~18 characters |
| `5` | 17872 | ~18 characters |
| `7` | 24575 | ~26 characters |

Over budget, in order of preference: **shorten the title**; or
`shrink_to_fit(doc, page, shape, max_lines=1)` from `office.template`, which
reduces the type until it is one line again — a 38-character title on design 7
settles at 47pt and looks right. Do not add a second text box beside the
placeholder: that is how a title and its subtitle both end up written through
the triangle.

**There is no cover design and no closing design.** Design 7 is a dark section
divider pressed into service for both, and a real deck title with a subtitle and
a date does not fit it — see the budget above. A title slide has to be built:
design 8 as a canvas, with title, subtitle and date drawn in the area the body
text vacates. Do not leave the deck without one.

**Neither is there a design for a diagram.** That is what the canvas row at the
top of the table is for, and on a deck with any analytical content you should
expect to use it several times.

## Design entries

Each entry's **Retext keys** are the strings this design actually carries, in
`retext`/`find_shapes` form. Copy a key from here rather than guessing one:
every failed `retext` in testing was a key remembered from a *different*
design. Any distinctive substring of a listed string works, and text is shown
truncated — `"Potentielle Navigatoren"` matches. `[3.1]` addresses a shape
*inside* the group at `[3]`. A design's title or source slot that is **absent
from its key list is empty in the template** — `retext` cannot find an empty
shape, so write it with `set_text(shapes(page)[i], ...)` or draw your own.


### Slide 1 — Section divider
- **Purpose:** To introduce a new section or topic within the presentation.
- **Content types:** `section-divider`
- **Capacity:** headline ✓ · body slots 0 · max bullets 0 · images 0 · chart ✗
- **Retext keys:** `[0]` "Projektorganisation"
- **Use when:** A new chapter or major thematic shift in the deck is required.
- **Avoid when:** Detailed content, bullet points, or data visualizations are needed.
- **Imagery:** a large white outlined triangle (decorative); the delta company logo (logo)
- **Notes:** The layout is designed for a single centered title; avoid adding additional text boxes to maintain the minimalist aesthetic. **One line only** — the title sits in the gap in the decorative triangle and ~18 characters fill it at 60pt; see the divider budget above.

### Slide 2 — Project Organization Chart
- **Purpose:** To visualize the project governance structure, key roles, and communication flow.
- **Content types:** `bulleted-content`, `team-people`
- **Capacity:** headline ✓ · body slots 2 · max bullets 6 · images 2 · chart ✗
- **Retext keys:** `[1]` "Ein zentraler Ansprechpartner und der Austau" · `[2]` "Projektorganisation" · `[5]` "Kontinuierliche Abstimmung (z.B. telefonisch" · `[6]` "Kommunikation" · `[8]` "Hr. QMB (QM-Betreuer)" · `[9]` "Hr. PL" · `[13]` "Hr. PM" · `[16]` "Fr. Ansprechpartner" · `[17]` "Hr. Ansprechpartner" · `[21]` "Kundenlogo" · `[22]` "Zwei Projektmitglieder"
- **Use when:** When showing a project org chart with associated communication guidelines.
- **Avoid when:** When a simple list of team members or a complex multi-level corporate hierarchy is needed.
- **Imagery:** a speech bubble icon (icon); a yellow square placeholder for a customer logo (photo-placeholder); the delta company logo (logo)
- **Notes:** The left side is a fixed-structure org chart with 6 specific role boxes; the right side is for communication bullet points.

### Slide 3 — Process Flow with Logos
- **Purpose:** To visualize a linear sequence of steps or milestones involving partner logos.
- **Content types:** `process-steps`
- **Capacity:** headline ✓ · body slots 4 · max bullets 2 · images 4 · chart ✗
- **Retext keys:** `[1]` "Nach einer Entscheidung zur Fortsetzung der" · `[2]` "Nächste Schritte" · `[7]` "&" · `[8]` "Kundenlogo" · `[14]` "Entscheidung zur Fortsetzung der Zusammenarb" · `[15]` "Vorlage eines Beratungsvertrags" · `[16]` "Abschluss des Beratungsvertrags mit der delt" · `[17]` "Erfolgreiche gemeinsame Umsetzung des Projek"
- **Use when:** Use when showing a 4-step chronological progression or a roadmap where each step is associated with a specific entity or logo.
- **Avoid when:** Avoid for complex non-linear processes, detailed data tables, or lists with more than four steps.
- **Imagery:** the delta company logo (logo); placeholder text 'Kundenlogo' (photo-placeholder)
- **Notes:** The layout uses a specific visual flow of interlocking circles and arrows; text must be kept concise to avoid overlapping the graphic elements.

### Slide 4 — Project Timeline and Costing
- **Purpose:** To present a detailed project schedule with associated effort (BT) and total cost calculations.
- **Content types:** `timeline`, `metrics-kpi`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 4 · images 0 · chart ✗
- **Grouped text:** 8 text shape(s) sit inside groups (26 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Bei einer aufwandsbasierten Schätzung belauf" · `[2]` "Zeit- und Aufwandsplanung" · `[5]` "Für das vorliegende Angebot wurde ein Berate" · `[6]` "Σ XX BT zu XX.XXX €" · `[20]` "Klausurzeiten" · `[21.0]` "2" · `[21.1]` "3" · `[21.2]` "1" · `[21.3]` "4" · `[21.4]` "6" · `[21.5]` "5" · `[21.6]` "7" · `[22.0]` "Meilensteine: 1 Projekt-Kick-Off 2 Termin/Ab" · `[23]` "Vorgehen: 1. Lösche die Spalten der nicht be" · `[25]` "Viele nützliche Informationen rund um das Th"
- **Use when:** When providing a detailed breakdown of project phases, milestones, and the resulting financial quote.
- **Avoid when:** When a high-level roadmap is needed without granular task-level effort or pricing.
- **Imagery:** none
- **Notes:** The central table is a manual construction of shapes; the builder must ensure the total cost box (bottom right) is updated in sync with the table totals.

### Slide 5 — Section Divider
- **Purpose:** To signal the start of a new section or chapter within the presentation.
- **Content types:** `section-divider`, `one-big-idea`
- **Capacity:** headline ✗ · body slots 1 · max bullets 1 · images 0 · chart ✗
- **Retext keys:** `[1]` "AZ"
- **Use when:** A new high-level topic is introduced that requires a visual break.
- **Avoid when:** Detailed content, lists, or data-heavy information are needed.
- **Imagery:** a white outlined stylized triangle (decorative); the delta company logo (logo)
- **Notes:** The text slot is centered within the decorative triangle and is intended for short identifiers or section titles only. **One line only** — ~18 characters at 60pt; a second line crosses the triangle. See the divider budget above.

### Slide 6 — Analysis Matrix
- **Purpose:** To present two parallel analysis streams (e.g., competitor vs. target group) with their respective findings.
- **Content types:** `bulleted-content`, `comparison`
- **Capacity:** headline ✓ · body slots 6 · max bullets 3 · images 2 · chart ✗
- **Grouped text:** 2 text shape(s) sit inside groups (21 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Der Direktvertrieb und das Erschließen einer" · `[2]` "Ausgangssituation und Zielsetzung" · `[6]` "Wettbewerbsanalyse" · `[7]` "Zielgruppenanalyse" · `[11]` "Ausgewählte Marktteilnehmer sind analysiert" · `[12]` "Der eigene USP ist geschärft und Handlungsem" · `[13]` "Mittels Umfrage sind Anforderungen von KMUs" · `[14]` "Unterschiede zu Start-Ups sind erfasst und H" · `[18.2]` "Der Einstieg in den Direktvertrieb bedarf ak" · `[19.2]` "KMUs wurden als neue Zielgruppe identifizier" · `[20]` "Alternatives Design zum Master"
- **Use when:** When you need to show a structured breakdown of two different analysis categories and their resulting outcomes.
- **Avoid when:** When presenting a linear process, a single deep-dive analysis, or data-heavy charts.
- **Imagery:** a laptop showing a chart on screen (photo-placeholder); a hand holding a pen writing on a document (photo-placeholder); crossed swords icon (icon); a target/bullseye icon (icon); a speech bubble icon (icon); a document/checklist icon (icon); the delta company logo (logo)
- **Notes:** The layout uses a mirrored structure: left-side context, center image, and right-side detailed findings with icons.

### Slide 7 — Section Divider Dark
- **Purpose:** To introduce a new section or chapter within the presentation.
- **Content types:** `section-divider`
- **Capacity:** headline ✓ · body slots 0 · max bullets 0 · images 0 · chart ✗
- **Retext keys:** `[1]` "Formelemente, Navigator"
- **Use when:** A new high-level topic or section is starting and requires a visual break.
- **Avoid when:** Detailed content, bullet points, or data-heavy information are needed.
- **Imagery:** a large white outlined triangle (decorative); the delta company logo (logo)
- **Notes:** The layout features a dark background with a single centered text slot for the section title. **One line only** — ~26 characters at 60pt; a second line crosses the triangle, and no automated check sees it. See the divider budget above.

### Slide 8 — Content slide with navigator
- **Purpose:** To present detailed text content while indicating the current progress or phase within a larger process.
- **Content types:** `bulleted-content`, `one-big-idea`
- **Capacity:** headline ✓ · body slots 1 · max bullets 6 · images 0 · chart ✗
- **Retext keys:** `[1]` "Potentielle Navigatoren entsprechend der ver" · `[2]` "Navigator (1/2)" · `[4]` "Funktion Der Navigator soll es dem Betrachte"
- **Use when:** Use when providing detailed explanations or instructions that are part of a multi-step sequence.
- **Avoid when:** Avoid when the slide requires heavy data visualization, multiple columns, or large imagery.
- **Imagery:** a sequence of five chevron arrows (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout includes a specific area for a progress navigator (chevrons) in the top right.

### Slide 9 — Navigator Showcase
- **Purpose:** To demonstrate or explain different navigation elements and frameworks used throughout a presentation.
- **Content types:** `bulleted-content`, `image-showcase`
- **Capacity:** headline ✓ · body slots 1 · max bullets 3 · images 6 · chart ✗
- **Grouped text:** 6 text shape(s) sit inside groups (13 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Potentielle Navigatoren entsprechend der ver" · `[2]` "Navigator (2/2)" · `[5.0.2]` "2" · `[5.0.3]` "1" · `[5.0.4]` "5" · `[5.0.6]` "4" · `[5.0.8]` "3" · `[5.2]` "Business Model Canvas" · `[8]` "Verwendung Die Navigatoren werden entspreche"
- **Use when:** When explaining the presentation's navigation logic or showcasing specific framework icons.
- **Avoid when:** When presenting primary data, complex tables, or long-form text documents.
- **Imagery:** a grey puzzle piece (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a puzzle piece partially filled with blue (photo-placeholder); a sequence of four document icons with a grey arrow (photo-placeholder); a simplified Business Model Canvas diagram with blue and green blocks (photo-placeholder); a chevron process arrow sequence in grey and green (decorative); the delta company logo (logo)
- **Notes:** The layout is designed for a mix of explanatory text on the left and a gallery of visual examples on the right.

### Slide 10 — Process Flow with Detail
- **Purpose:** To visualize a linear progression of phases with a detailed breakdown of a specific step or result below.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 3 · images 0 · chart ✗
- **Grouped text:** 16 text shape(s) sit inside groups (11 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[3]` "Phase 1" · `[4]` "Phase 3" · `[5]` "Phase 2" · `[8]` "Lorem ipsum dolor sit amet, consetetur sadip" · `[9.0.0]` "Smarte Bewer-bung" · `[9.0.1]` "HSG-Partner-Schaft" · `[9.0.2]` "Campus-Events" · `[9.0.3]` "Flyer" · `[9.0.4]` "Sponso-ring" · `[9.0.5]` "Stipen-dium" · `[9.0.6]` "Case" · `[9.1]` "✓" · `[9.2]` "✗" · `[9.8]` "Beispiel" · `[9.9]` "Betrachtung der generierten Ideen unter Berü" · `[10]` "Layout"
- **Use when:** Use when showing a high-level 3-step process and providing a detailed explanation or sub-process for one of those steps.
- **Avoid when:** Avoid for complex timelines with many dates or slides requiring heavy data tables.
- **Imagery:** a grey right-pointing arrow (decorative); a grey right-pointing arrow (decorative); a sequence of connected blue hexagonal shapes (photo-placeholder); an orange slanted rectangle label (decorative); an orange slanted rectangle label (decorative); the delta company logo (logo)
- **Notes:** The layout contains a fixed graphic of connected hexagons that serves as a visual placeholder for a sub-process or detailed breakdown.

### Slide 11 — Phased Data Comparison
- **Purpose:** To show a progression of phases alongside a corresponding data chart and explanatory notes.
- **Content types:** `process-steps`, `data-chart`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 5 · max bullets 2 · images 0 · chart ✓
- **Retext keys:** `[3]` "Phase 1" · `[5]` "Phase 2" · `[7]` "Phase 3" · `[10]` "Erklärung" · `[11]` "Beispiel" · `[12]` "Charttitel" · `[13]` "Layout"
- **Use when:** Use when you need to link specific project phases or steps to a quantitative chart result.
- **Avoid when:** Avoid for simple lists or complex tables that require full-slide width.
- **Imagery:** a hand holding a coin with a dollar sign (icon); a magnifying glass (icon); a calendar with a growth chart line (icon); the delta company logo (logo)
- **Notes:** The layout contains three fixed phase labels on the left and a chart area on the right; ensure text in 'Erklärung' (explanation) remains brief to avoid overlapping the chart.

### Slide 12 — Two-column comparison with examples
- **Purpose:** To compare two parallel workstreams or categories, each supported by a visual example and a list of detailed activities.
- **Content types:** `two-column`, `comparison`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 4 · images 2 · chart ✗
- **Retext keys:** `[1]` "Die Verantwortungen und Tätigkeiten der Mita" · `[2]` "Phase II: Prozessdokumentation (1/2)" · `[4]` "Modellierung" · `[5]` "Rollenbeschreibungen" · `[9]` "Beispiel" · `[11]` "Definieren der Verantwortungs- und Aufgabenb" · `[12]` "Modellieren der Prozesse und Schnittstellen"
- **Use when:** When you need to contrast two distinct processes or deliverables that both require a visual aid and a supporting list of bullet points.
- **Avoid when:** When there is only one main topic or when the content requires a complex data chart instead of a simple screenshot/diagram.
- **Imagery:** a person icon (icon); a flowchart/process icon (icon); a document layout showing a role description (photo-placeholder); a process flow diagram with blue boxes (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout includes specific 'Beispiel' (Example) labels over the image areas that should be maintained or updated.

### Slide 13 — Three-phase process flow
- **Purpose:** To visualize a linear project timeline or process divided into three distinct phases with associated tasks.
- **Content types:** `process-steps`, `timeline`
- **Capacity:** headline ✓ · body slots 3 · max bullets 3 · images 0 · chart ✗
- **Grouped text:** 7 text shape(s) sit inside groups (17 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[3.1]` "Review S.AG" · `[5.1]` "Abschluss" · `[6.1]` "Eine Woche" · `[6.2]` "Zusammengefasste Rechercheergebnisse" · `[7.1]` "Longlist in abgestimmter Struktur" · `[8]` "Phase II: Detailrecherche und Analyse" · `[9]` "Phase I: Identifizieren von Themenfeldern" · `[10]` "Phase III: Aufbereiten und Abschluss" · `[14]` "Festlegen von Themenfeldern pro Zukunftsfeld" · `[15]` "Analyse der Themenfelder hinsichtlich der Tä" · `[16]` "Zusammenfassen der Rechercheergebnisse inklu"
- **Use when:** Use when presenting a high-level project roadmap or a three-step methodology with detailed action items for each phase.
- **Avoid when:** Avoid for complex processes with more than three steps or for content that does not follow a linear chronological sequence.
- **Imagery:** a target with an arrow (icon); a magnifying glass (icon); a checklist on a clipboard (icon); a red diamond (decorative); the company logo (logo)
- **Notes:** The layout uses a chevron-style header for phases; ensure phase titles are concise to avoid text overflow in the blue shapes.

### Slide 14 — Five-pillar analysis
- **Purpose:** To present five distinct dimensions or categories of a strategic analysis surrounding a central core concept.
- **Content types:** `bulleted-content`, `process-steps`
- **Capacity:** headline ✓ · body slots 5 · max bullets 3 · images 0 · chart ✗
- **Grouped text:** 6 text shape(s) sit inside groups (16 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Bei der strategischen Marktanalyse werden fü" · `[2]` "Phase II: Aspekte der strategischen Marktana" · `[5.1]` "Strategische Marktanalyse" · `[6]` "Welche kritischen Faktoren sind zu beachten?" · `[7]` "Wie kann der Zielmarkt definiert werden? Wie" · `[8]` "Wie sind ausgewählte Wettbewerber aufgestell" · `[9.2.0]` "Segmentierung des Zielmarkts" · `[10.2.0]` "Analyse des Marktvolumens" · `[11.2.0]` "Potentialanalyse des Zielmarkts" · `[12.2.0]` "Wettbewerberanalyse" · `[13.2.0]` "Kundenanalyse" · `[14]` "Wie groß ist das Marktvolumen des Nischenmar" · `[15]` "Wie gestaltet sich der (potentielle) Kundenk"
- **Use when:** Use when you need to break down a central topic into five specific analytical dimensions or question sets.
- **Avoid when:** Avoid when you have more or fewer than five key points, or when the content requires a linear sequence rather than a hub-and-spoke relationship.
- **Imagery:** four arrows pointing outwards (icon); a bar chart on a screen (icon); three small people figures (icon); a speech bubble with a person silhouette (icon); a bar chart (icon); the delta company logo (logo)
- **Notes:** The layout is strictly balanced around a central circle; ensure text in the five surrounding boxes is concise to avoid overlapping the center element.

### Slide 15 — Process Steps with Details
- **Purpose:** To visualize a sequential process or list of steps where specific steps are highlighted with detailed explanatory text.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 7 · max bullets 3 · images 0 · chart ✗
- **Grouped text:** 1 text shape(s) sit inside groups (29 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Die Marktanalyse wird in fünf Schritten durc" · `[2]` "Phase II: Durchführen der Marktanalyse (2/2)" · `[4]` "Potentialanalyse des Zielmarktes" · `[5]` "Kundenanalyse" · `[6]` "Wettbewerberanalyse" · `[7]` "Analyse des Marktvolumens" · `[8]` "Segmentierung des Zielmarktes" · `[10]` "Basierend auf dem Marktvolumen des Nischenma" · `[11]` "Kritische Faktoren sowie aktuelle Trends sin" · `[23.1]` "Die Marktanalyse dient als Grundlage für das" · `[26]` "Beispiel"
- **Use when:** Use when presenting a multi-step process where one or more steps require expanded descriptions or examples.
- **Avoid when:** Avoid for simple lists without detailed call-outs or for complex timelines with specific dates.
- **Imagery:** a pie chart icon (icon); a magnifying glass over a document icon (icon); a list with a magnifying glass icon (icon); a human head profile icon (icon); a bar chart with an upward arrow icon (icon); the delta company logo (logo)
- **Notes:** The layout uses a bracket-style visual connector to link specific process steps to their corresponding detail boxes on the right.

### Slide 16 — Process List
- **Purpose:** To outline a sequential process or a series of steps with corresponding detailed descriptions.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 5 · max bullets 3 · images 0 · chart ✗
- **Grouped text:** 2 text shape(s) sit inside groups (21 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Die Ergebnisse der Umfrage werden analysiert" · `[2]` "Zielgruppenanalyse – Phase 3: Umfrageauswert" · `[4]` "Umfragedaten auswerten" · `[6]` "Statistisches Auswerten der Ergebnisse Vergl" · `[7]` "Anforderungen und Trends visualisieren" · `[9]` "Visualisieren der Analysen in Diagrammen Abl" · `[10]` "Handlungs-empfehlungen ableiten" · `[12]` "Ausarbeiten von Handlungsergebnissen auf Bas" · `[13]` "Ergebnis-präsentation" · `[15]` "Zusammenfassen der Analysen Ergebnispräsenta" · `[16.2]` "Umfang an Daten" · `[16.3]` "Nutzen aus Daten" · `[17]` "Aussortieren von unvollständigen Antworten H" · `[18]` "Ergebnisse bereinigen"
- **Use when:** Use when presenting a linear workflow or a list of five key phases that require brief explanatory bullets.
- **Avoid when:** Avoid for complex diagrams, large data tables, or processes with more than five steps.
- **Imagery:** a data table (photo-placeholder); a funnel (photo-placeholder); a bar chart with an upward arrow (photo-placeholder); a funnel (photo-placeholder); a document with charts (photo-placeholder); a blue vertical arrow shape (decorative); the delta company logo (logo)
- **Notes:** The layout features a central column of placeholder graphics that should be replaced with icons or visuals representing each process step.

### Slide 17 — Iterative Process Flow
- **Purpose:** To visualize a cyclical development process that transforms initial specifications into a final model through iterative feedback loops.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 5 · images 0 · chart ✗
- **Grouped text:** 3 text shape(s) sit inside groups (36 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Die Markttreiber und -hemmnisse werden itera" · `[2]` "Vorgehen zum Bilden eines Modells von Marktt" · `[4]` "Konzeptionierung" · `[9]` "Projektverlauf" · `[10]` "Umfassendes Marktmodell" · `[11]` "Erkenntnisse aus Projektphasen" · `[14]` "delta" · `[15]` "Initiale Problemspezifikation Entwickeln des" · `[16]` "Iterative Entwicklung parallel zu Projektpha" · `[17]` "Gemeinsames Verständnis der Marktsituation K" · `[19]` "A" · `[20]` "C" · `[21]` "A B" · `[23]` "A C" · `[25]` "C A" · `[28]` "delta & Häcksler" · `[30.1]` "Konsens-bildung über das Modell" · `[31.1]` "Struktur-validierung des Modells" · `[32.1]` "Qualitativer Entwurf des Modells"
- **Use when:** Use when showing a non-linear workflow where inputs are refined through a central iterative cycle to reach a set of outcomes.
- **Avoid when:** Avoid for simple linear timelines or slides requiring heavy data visualization.
- **Imagery:** a circular arrow diagram representing a cycle (decorative); the company logo 'delta' (logo)
- **Notes:** The central cycle contains three fixed text areas; the left and right columns are for bulleted lists of inputs and outputs respectively.

### Slide 18 — Process flow with images
- **Purpose:** To visualize a linear multi-step process where each phase is supported by a representative image and a descriptive label.
- **Content types:** `process-steps`, `image-showcase`
- **Capacity:** headline ✓ · body slots 8 · max bullets 1 · images 4 · chart ✗
- **Retext keys:** `[1]` "Auf Basis einer Unternehmens- und einer Tren" · `[2]` "Phasen mit Bildern" · `[4]` "Transformation" · `[5]` "Strategie- entwicklung" · `[6]` "Unternehmens- analyse" · `[7]` "Externer/neutraler Blick" · `[8]` "Trendbewertung" · `[9]` "Produktportfolio inkl. strategischer Einordn" · `[10]` "Begleitung der Transformation" · `[16]` "Trendanalyse"
- **Use when:** You need to show a 4-step sequence where each step requires a visual aid and a short caption.
- **Avoid when:** The process has more or fewer than four steps, or requires detailed bulleted explanations for each phase.
- **Imagery:** a pie chart and line graph icon (icon); a compass icon (icon); a lightbulb icon (icon); a strategy map with X and O markers icon (icon); a magnifying glass over documents (photo-placeholder); a hand drawing arrows on a blackboard (photo-placeholder); people looking at maps or blueprints on a table (photo-placeholder); a movie clapperboard (photo-placeholder); the delta company logo (logo)
- **Notes:** The layout is strictly locked to 4 columns; the chevron headers and bottom labels must be kept concise to avoid text overflow.

### Slide 19 — Three-column process analysis
- **Purpose:** To present three parallel analysis streams, each with a visual representation, descriptive bullets, and a resulting outcome.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 3 · max bullets 2 · images 3 · chart ✗
- **Grouped text:** 12 text shape(s) sit inside groups (18 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Um relevante Stärken und Schwächen, Potenzia" · `[2]` "Modul 1: Unternehmensanalyse – Ergebnisse (B" · `[4]` "Kompetenzanalyse" · `[8.1.1]` "mittel" · `[8.1.2]` "niedrig" · `[8.1.3]` "hoch" · `[10.0.0]` "Value Proposition" · `[10.0.1]` "Key Partners" · `[10.0.2]` "Key Activities" · `[10.0.3]` "Customer Relationship" · `[10.0.4]` "Customer Segments" · `[10.0.5]` "Cost Structure" · `[10.0.6]` "Revenue Streams" · `[10.0.7]` "Key Resources" · `[10.0.8]` "Channels" · `[11]` "Grundlage für die Entwicklung und Bewertung" · `[12]` "Grundlage für aus der Produktstrategie resul" · `[13]` "Grundlage für mögliche Anpassungen des Gesch" · `[14]` "Erfassen relevanter Kernkompetenzen Bewerten" · `[15]` "Veranschaulichen der Wertschöpfung der OG Er" · `[16]` "Abbilden des aktuellen Geschäftsmodells mitt"
- **Use when:** When comparing three distinct strategic modules or analysis steps that lead to specific foundations or conclusions.
- **Avoid when:** When there are more or fewer than three categories, or when long-form text is required.
- **Imagery:** a colored grid matrix (photo-placeholder); a value chain diagram with arrows (photo-placeholder); a Business Model Canvas diagram (photo-placeholder); three blue chevron arrows (decorative); the delta company logo (logo)
- **Notes:** The layout relies on a strict three-column symmetry; ensure text in the bottom 'Grundlage' slots is concise to avoid overlapping the blue accent lines.

### Slide 20 — Two-column process detail
- **Purpose:** To explain a specific module or process step by splitting the approach and the sources into two distinct lists.
- **Content types:** `bulleted-content`, `two-column`
- **Capacity:** headline ✓ · body slots 2 · max bullets 5 · images 1 · chart ✗
- **Retext keys:** `[1]` "In Modul 2 werden Trends mit Auswirkungen au" · `[2]` "Modul 2: Trendanalyse" · `[4]` "Vorgehen" · `[5]` "Recherchequellen" · `[12]` "Identifizieren von Trends mit Auswirkungen a" · `[13]` "Durchführen einer Literatur- und Desktoprech"
- **Use when:** When you need to present a 'how-to' (Vorgehen) alongside the 'where from' (Recherchequellen) for a specific project phase.
- **Avoid when:** When you have a single long list of items or need to present complex data charts.
- **Imagery:** a person wearing a VR headset seen through a magnifying glass (photo-placeholder); a gear and document icon (icon); a book/library icon (icon); three blue chevron arrows (decorative); the delta company logo (logo)
- **Notes:** The layout features a central circular image placeholder that overlaps the two text columns; ensure text does not wrap into the image area.

### Slide 21 — Process Diagram with Detail
- **Purpose:** To visualize a non-linear process or conceptual framework with a detailed explanation of one specific component.
- **Content types:** `process-steps`, `bulleted-content`
- **Capacity:** headline ✓ · body slots 2 · max bullets 5 · images 0 · chart ✗
- **Grouped text:** 7 text shape(s) sit inside groups (16 top-level) — `retext`/`find_shapes`/`qa.py` reach them, `shapes(page)` does not.
- **Retext keys:** `[1]` "Abhängig von den Analyseergebnissen wird ein" · `[2]` "Modul 3: Strategieentwicklung – (1) Kreativp" · `[7.0]` "Longlist möglicher Produkte" · `[8]` "Offene Ideenfindung Freies und unvoreingenom" · `[10.1]` "Kreativmethoden Ideengenerierung mit Mitarbe" · `[11.1]` "Customer Needs Identifikation der Kundenbedü" · `[12.1]` "Radical Game Ideengenerierung basierend auf" · `[13.1]` "Strategie-AK Diskussion der Produktideen im" · `[14.1]` "Dokumentation Erfassen der Produktideen für" · `[15.1]` "Szenariotechnik Planspiel basierend auf der"
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

### Build with `build_deck`, not `clone_slide`

**Use `office.deckbuild.build_deck` for this template.** It duplicates designs
instead of rebuilding their shapes, then sets the slide order by permuting
`<p:sldIdLst>` in the saved file. Measured over designs 7, 8, 13, 20 and 1
against their source renders: mean RMS **7.8 versus 22.8** for `clone_slide`,
better on every slide.

That matters here specifically because this template leans on three things a
shape rebuild destroys: **bullet markers**, **per-run colour inside a paragraph**
(the blue-on-black body text is everywhere), and **think-cell's OLE objects** —
6 references survive a `build_deck` build, none survive `clone_slide`. The
divider footer line survives too.

The rest of this section describes `clone_slide`'s behaviour, which still
applies if you use it — for an `.odp`, or when editing a deck in place.

### Cloning dividers with `clone_slide` — fixed, but know what it cost

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

### Inline formatting in body text boxes — avoided, not repaired

The body text on these slides sits in ordinary `ooxml-rect` shapes, not
placeholders, so `clone_shape` handles it — and it loses bullet markers and
inline run colour, flattening a paragraph that mixed blue and black to one
colour. Compare a `clone_slide` copy of slide 20 against the original and the
bullets are visibly plainer.

This was never fixed in `clone_slide`; `build_deck` sidesteps it by not
rebuilding shapes at all (slide 20: RMS 2.5 via `build_deck`, 31.5 via
`clone_slide`). If you do use `clone_slide` on this template, budget for fixing
body text by hand and check the render before shipping.

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
  flags — designs 6, 10, 11, 12, 15. Design 10 keeps its `Beispiel` **inside a
  group**.
- Design 4's yellow `Vorgehen:` block (four numbered build steps) and its orange
  toolbox tip box.
- Yellow highlight marks `Kundenlogo` (designs 2, 3), `Kunde`, `XX.XXX`, `XX`,
  `450` — fill-in markers, replace with real values or the customer's logo.

**`strip_annotations(page)` deletes them for you** — call it in `fill`, before
retexting:

```python
from office.template import strip_annotations, retext

def fill(doc, page, slot, design):
    strip_annotations(page)          # returns what it removed
    retext(page, {...})
```

It matches whole strings, so design 17's title *Vorgehen zum Bilden eines
Modells* and design 20's `Vorgehen` column header are correctly left alone. The
fill-in markers (`XX.XXX`, `450`) are deliberately **not** deleted — they are
values to replace, not flags to drop.

`qa.py` will not flag any of this; they are ordinary shapes carrying short
strings.

### Gaps in the collection

**No cover, no closing/contact slide, no diagram design** — see the notes under
the decision table. Designs 1, 5 and 7 are all section dividers on
`Section_dark`, which is why "deck opener" and "closing slide" both point at
one.

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
