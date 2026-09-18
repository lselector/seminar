---
name: gslides-epigraphs
description: Extract the text of the live Google Slides seminar deck and write 30 candidate epigraphs for the week (short, energetic, funny, elegant) into epigraphs.txt in the current directory, for the author to pick one. Use when asked for epigraphs, taglines, a motto or one-liners for this week's deck, or to extract the text of the deck.
---

# Write 30 epigraphs for the deck

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

## Step 1: extract the text

```bash
python3 g9_deck_text.py --out /tmp/deck-text.md   # or: ... 2026-09-25 --out ...
```

It writes the talk (slide 1 to the separator) as Markdown:
each slide's title, then its boxes top to bottom, tables one
row per line. It changes nothing in the deck. Read the whole
file.

## Step 2: find the week

The news topics are the week. Note each headline and its
most striking fact: a number, a name, an odd detail. The
standing pages (benchmarks, the Intelligence Index, YouTube,
layoffs) are background; use them for a line or two, not
more. If the deck has few news topics yet, say so in the
report: the epigraphs can only be as specific as the deck.

## Step 3: write 30 epigraphs

Each one:

- **Short**: 70 characters or fewer, most well under.
- **Energetic**: strong verbs, a clear turn, a rhythm that
  lands. Energy comes from the words, not from exclamation
  marks, capitals or emoji.
- **Funny**: wit, irony, a surprising pairing, an
  understatement. Laugh with the audience, never at people
  who lost jobs or at any group.
- **Elegant**: one idea, stated flat and clean. Two short
  sentences at most.
- **About this week**: a reader could guess at least one
  headline from it. No line that would fit any week.

Across the 30: vary the targets and the forms (a contrast,
a twist on a saying, a deadpan fact, a question, a mock
proverb). No two lines on the same joke. Follow the
`humanize` skill: plain words, straight quotes, no
em-dashes, none of its banned vocabulary. Attribute any real
quote.

## Step 4: write the file

Write `epigraphs.txt` in the current directory (`work/`),
one epigraph per line, no numbers, no blank lines, no other
text. Overwrite the file if it exists. Check every line is
70 characters or fewer.

## Step 5: report

Give the file path, how many lines, and your three
favourites. The deck itself is not changed. To use one, the
author types it into the epigraph box on slide 1, or asks
for it to be passed to step 3:

```bash
echo '{"epigraph": "..."}' > /tmp/epi.json
python3 g3_update_toc.py --json /tmp/epi.json
```

That works only while the box still shows its placeholder.
