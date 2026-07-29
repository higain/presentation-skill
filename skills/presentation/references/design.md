# Design: choosing a look for a deck that has no template

Loaded from the presentation skill's SKILL.md. Read this when you are **inventing
a look** — a deck from scratch, or a restyle. Filling a brand template does not
need it: the template already made these decisions, and you should follow it.

**What goes *on* a slide is a different question, and this file does not answer
it.** The visual-per-slide rule, layout variety, spacing and the list of tells
that make a deck look generated live in SKILL.md's **Composition** section, and
they apply to every deck — template or not. This file is only about the look.

### Choosing the palette

- **Pick a bold, content-informed palette.** If swapping your colors into an
  unrelated deck would still "work", the choice wasn't specific enough.
- **Dominance over equality**: one color at 60-70% visual weight, 1-2
  supporting tones, one sharp accent. Never equal weight.
- **Dark/light contrast**: dark title and closing slides, light content
  ("sandwich"), or commit to dark throughout for a premium feel.
- **Don't default to blue** — pick colors that reflect the topic.
- **Don't default to cream or beige backgrounds** — use white or the palette.

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

In UNO, colors are integers: `0x1E2761`, not `"1E2761"`. `office.draw` takes
them in that form.

### Committing to a motif

**Pick one motif and repeat it** across the deck — rounded image frames, icons
in colored circles, a consistent card shape. A motif is what makes a set of
slides read as one deck.

Not a colour bar and not an accent stripe: those are the tells, and SKILL.md's
Composition section says why.

### Consistency

- **Don't style one slide and leave the rest plain.** Whatever you choose here,
  apply it everywhere — a half-styled deck looks worse than an unstyled one.
- Fonts: pick from the trustworthy list in SKILL.md's Typography section, or
  accept that the overflow check cannot be relied on.
