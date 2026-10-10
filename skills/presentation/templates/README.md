# Template packs

**This skill ships with no template.** A pack is what teaches it yours, and
this directory is where packs live. One directory per template:

```
templates/
  README.md          this file — the only thing here a published skill contains
  default            optional: one line naming the pack to use by default
  acme/              a pack, named whatever you call the brand
    template.pptx      the master itself; build_deck duplicates slides out of it
    profile.json       palette, type scale, and the template's authoring notes
    catalog.md         which design carries which content  (generated)
    designs/           one file per design: keys, slots    (generated)
    overrides.md       the prose no script can measure     (hand-written)
    catalog/           extraction dossiers and the judged entries
```

## Making one

Run the **`presentation-setup`** skill and point it at your template. It reads
the file, counts the palette and the type sizes off it, asks you which colour
is the accent and which designs mean what, and writes the pack. Twenty minutes
for a 20-design master, and the last stage builds three real slides so you can
see whether it worked.

`../README.md` is the same pipeline as a manual, for anyone who would rather
drive the scripts directly.

## Which pack a deck is built on

1. The template the user names, or `--template` on the command line.
2. Otherwise the pack named in `templates/default`.
3. Otherwise, if exactly one pack is installed, that one.

Two packs and no `default` file is not guessed at — picking one of a company's
two brands at random is worse than asking.

## What lives in a pack, and what does not

A pack holds **facts about one template**: its design numbers, its retext keys,
its colours, its title budgets. Anything true of every template — how to draw a
Venn, what makes a title an action title, what QA measures — belongs in
`references/` and never in here.

The line matters because it is the one that decays. A design number written
into `references/` is wrong for everybody else's template, and nothing fails
when it happens: `tests/test_generic_docs.py` is what fails instead.

## Privacy

`templates/*` is gitignored, this README excepted. A company template, its
renders and its catalog are brand-internal and stay on the machine that built
them. Nothing in the extraction pipeline sends a slide anywhere; the two steps
that can call a model are opt-in and say so.
