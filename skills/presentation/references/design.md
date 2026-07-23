# Design: making a deck that doesn't look generated

Loaded from the presentation skill's SKILL.md. Read this when you are inventing
a look — a deck from scratch, or a restyle. Filling a brand template does not
need it: the template already made these decisions, and you should follow it.


**Don't create boring slides.** Plain bullets on white won't impress anyone.

### Before starting

- **Pick a bold, content-informed palette.** If swapping your colors into an
  unrelated deck would still "work", the choice wasn't specific enough.
- **Dominance over equality**: one color at 60-70% visual weight, 1-2
  supporting tones, one sharp accent. Never equal weight.
- **Dark/light contrast**: dark title and closing slides, light content
  ("sandwich"), or commit to dark throughout for a premium feel.
- **Commit to one motif** and repeat it — rounded image frames, icons in
  colored circles. Not a color bar or accent stripe (see Avoid).

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

In UNO, colors are integers: `0x1E2761`, not `"1E2761"`.

### For each slide

**Every slide needs a visual element** — image, chart, icon, or shape.
Text-only slides are forgettable.

- Two-column (text left, illustration right)
- Icon + text rows (icon in a colored circle, bold header, description below)
- 2×2 or 2×3 grid of content blocks
- Half-bleed image down one side with content overlaid
- Large stat callouts (60-72pt numbers, small labels below)
- Comparison columns (before/after, pros/cons)
- Timeline or process flow (numbered steps, arrows)

### Spacing

0.5" minimum margins · 0.3-0.5" between blocks · leave breathing room.

### Avoid

- **Don't repeat one layout** — vary columns, cards and callouts
- **Don't center body text** — left-align paragraphs and lists; center only titles
- **Don't skimp on size contrast** — titles 36pt+ against 14-16pt body
- **Don't default to blue** — pick colors that reflect the topic
- **Don't mix spacing randomly** — pick 0.3" or 0.5" and hold to it
- **Don't style one slide and leave the rest plain**
- **Don't create text-only slides**
- **Don't use low-contrast elements** — icons *and* text need contrast
- **NEVER use accent lines under titles** — a hallmark of AI-generated slides;
  use whitespace or background color instead
- **NEVER add decorative color bars or accent stripes** — header/footer bars,
  vertical sidebar stripes, thin stripes along a card edge, single-side borders.
  To set a card apart use a subtle background tint, a shadow, or an icon
- **Don't default to cream/beige backgrounds** — use white or the brand palette
- **Don't ship text that overflows its shape** — shrink, split, or enlarge

