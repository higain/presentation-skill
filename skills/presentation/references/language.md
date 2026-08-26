# Language: what the words on a slide have to do

Loaded from SKILL.md whenever slide text is being **written** rather than copied from
what the user supplied. Their wording is theirs — fix it only when asked.

The gap between a mediocre deck and a good one is mostly title quality, and title
quality is the part that automates. `scripts/qa.py` checks what a machine can check
here; §5 says which parts it cannot, and those stay yours.

---

## 1. Action titles

**Every content slide's title is a complete sentence stating the conclusion** — the
"so what" the audience should draw — not the topic the slide is about.

| Rule | |
|---|---|
| A full sentence | subject and an active verb |
| **Ends with a period** | this is the convention that makes the rest checkable |
| One idea | the title asserts it; the slide proves it |
| Quantified | wherever the evidence allows a number |
| ≤ 14 words | two lines at title size, no more |
| Never a label | "Marktgröße", "Q3 Overview", "Nächste Schritte" |
| Never a question | a title answers; the speaker asks |
| Never a hedge | "Einige Überlegungen zu…", "Some considerations regarding…" |

| Label — buried meaning | Action title — the so-what |
|---|---|
| "Q3 Financial Overview" | "Q3 revenue beat forecast by 15%, driven by APAC." |
| "Market Size" | "The €50B market is fragmented, with no national player." |
| "Growth Trends" | "50% growth with falling margins is not sustainable." |
| "Implementation Timeline" | "Three projects deliver the strategy over six quarters." |
| "Abhängigkeiten" | "Auf 42 geprüfte Bibliotheken kommen 1.447 installierte." |
| "Nächste Schritte" | "Drei Maßnahmen starten im Oktober, je mit benanntem Owner." |

**German titles follow the same rules**, period included. German runs ~20% longer than
English, so the 14-word cap bites sooner — cut the noun stacks first
("Optimierungspotenzialanalyse" is not a word anybody says out loud).

**Exempt from all of this:** the title slide, section dividers, and a single-statement
slide whose whole content *is* the sentence. A divider reading "01 — Wo wir stehen" is
correct as a label.

### The ghost deck test

Read **only the titles, in order.** They must tell the whole argument on their own.
`qa.py` prints them as a block for exactly this reason — a machine cannot judge whether
they flow, but it can put them in front of you in the ten seconds it takes to tell.

If they don't flow, the fix is in `plan.md` (see `storyline.md`), not in the slides.

---

## 2. Copy rules for everything else

- **Active voice.** "Wir schätzen…", not "Es wurde geschätzt…".
- **Their vocabulary**, not ours. Use the audience's words for their own business.
- **No stacked buzzwords.** One plain verb beats three impressive ones.
- **Don't state the obvious.** Every line adds a fact, a number, or a consequence.
- **Parallel structure** within any list: all fragments or all sentences, all verb-first
  or all noun-first. Never mixed.
- **Bullets ≤ 8 words**, one level deep, and never a bullet that restates its header.
- **Sources on the slide** that carries the claim, not in a notes field nobody opens.

---

## 3. Banned words and phrases

These are the tells. `qa.py` matches them literally, so this list is the specification,
not an illustration — add to it in the file, not only here.

**English:** delve · leverage (use "use") · robust · seamless · holistic · synergy ·
cutting-edge · state-of-the-art · empower · unlock · unleash · supercharge · elevate ·
transformative · game-changing · best-in-class · world-class · next-generation ·
paradigm shift · deep dive (as copy) · low-hanging fruit · win-win · actionable
insights · key learnings · drive value · a wide range of · various stakeholders ·
"in today's fast-paced world" · "it's important to note" · "at the end of the day"

**German:** ganzheitlich · nahtlos · zukunftsweisend · wegweisend · maßgeschneidert ·
Mehrwert (say which) · Synergien · Leuchtturmprojekt · Quantensprung · Optimierungs­-
potenzial · Handlungsfelder · Erfolgsfaktoren · proaktiv · "die Mitarbeiter abholen" ·
"auf Augenhöhe" · "in der heutigen schnelllebigen Zeit" · "Herausforderungen meistern" ·
"an dieser Stelle sei erwähnt"

Some have honest uses — *robust* standard errors, a real platform *ecosystem*,
a literal *deep dive* session. That is why these are reported as prose to look at and
**never fail a build.** Overriding a finding is a decision you are allowed to make; not
noticing it is not.

---

## 4. Banned patterns

| Pattern | What it looks like |
|---|---|
| **Bullet soup** | three bullets restating the header with different filler |
| **Empty triads** | "faster, better, stronger" with no number behind any of them |
| **Fragment chains** | "Growth → margin pressure → risk" in place of a sentence |
| **Symmetric fluff** | padding a short column so both have exactly four bullets |
| **Hedge stacking** | "could potentially help enable" → "cuts", or delete the claim |
| **Title Case Everywhere** | sentence case, except defined proper headers |

---

## 5. The four rewrite tests

Apply to every line. Each is a yes/no about one line, which is why they work as a final
pass even under a small model.

1. **Partner test** — would a senior colleague say this sentence out loud to the person
   it is aimed at? If it reads like a press release, rewrite it.
2. **Evidence test** — does the line carry a fact, number, name or consequence? If not,
   cut it or merge it into the line above.
3. **Verb test** — is the main verb concrete ("cuts", "doubles", "wins") rather than
   abstract ("enables", "supports", "drives")?
4. **Deletion test** — if deleting the line changes nothing, delete it. Whitespace beats
   filler.

**What `qa.py` can check:** the period, the word count, questions as titles, the banned
list, a missing title, a "Thank you" ending, and the ghost-deck listing.

**What it cannot, and you must:** whether the titles flow, whether the claim is true,
whether the proof proves it, and all four tests above. A clean prose report means nothing
was mechanically wrong — not that the deck argues.

---

## Attribution

The action-title rules, the banned list, the banned patterns and the four rewrite tests
are adapted from `executive-slide-skill/reference/language.md` (Apache-2.0); the ghost
deck test from `academic-pptx-skill` (MIT). The German list is ours. See `THIRD-PARTY.md`.
