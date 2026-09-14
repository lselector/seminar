---
name: gslides-update-toc
description: Rebuild the table of contents on the first slide of the live Google Slides seminar deck from the slides as they are now, and write the week's epigraph if its placeholder is still empty. Leaves a contents list or epigraph the author filled or wrote alone. Use when asked to update the contents, TOC, agenda or epigraph of the Google Slides deck.
---

# Update the contents and epigraph

Run from `work/`. The rules are in `work/README.md`.

## Contents

```bash
python3 g3_update_toc.py          # or: g3_update_toc.py 2026-09-25
```

It lists the first line of every topic box from slide 1 to
the separator, in the current slide order, including boxes
the author added. Benchmarks, the YouTube page, About and
Thank You are left out, as in the archive decks.

If the log says `left alone, you filled ...`, the author has
taken over the contents. Tell them; do not work around it.

## Epigraph

Only when the epigraph box on slide 1 still shows
`[epigraph: ...]`. Read the topics first (`python3
g2_add_news.py --list`), then write one line and pass it:

```bash
echo '{"epigraph": "Agents got their own computers. Oversight is still catching up."}' > /tmp/epi.json
python3 g3_update_toc.py --json /tmp/epi.json
```

What makes a good one, from the archive:

- First we taught machines to answer. Now we teach them to act.
- Reviewing and validating is the new bottleneck.
- The frontier is shifting from larger models to better
  scaffolding around the models we already have.

Rules: 70 characters or fewer; about this week, so a reader
could guess the headlines; one idea stated flat; no hype and
no exclamation marks; attribute any quote. Follow the
`humanize` skill.

Offer it to the user as a suggestion. Once it is written the
script never replaces it, so they can change it by hand.

## Report

Say how many headlines are listed, whether the contents or
epigraph were frozen by the author, and give the epigraph
you wrote, if any.
