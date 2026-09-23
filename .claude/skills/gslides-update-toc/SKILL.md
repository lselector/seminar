---
name: gslides-update-toc
description: Rebuild the table of contents on the first slide of the live Google Slides seminar deck from the slides as they are now, and write the week's epigraph if its placeholder is still empty. Leaves a contents list or epigraph the author filled or wrote alone. Use when asked to update the contents, TOC, agenda or epigraph of the Google Slides deck.
---

# Update the contents and epigraph

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

## Contents

```bash
python3 g3_update_toc.py          # or: g3_update_toc.py 2026-09-25
```

It lists the first line of every topic box from slide 1 to
the separator, in the current slide order, including boxes
the author added. Benchmarks, the YouTube page, About and
Thank You are left out, as in the archive decks.

The contents is four boxes, two a side, filled by the
script: left upper light yellow, left lower light green,
right upper light blue, right lower light yellow. Items run
in slide order and fill one box before the next, in that
order; each box holds 8 to 12 lines (as many as fit its
side), and an empty box shows `xxx`. The left side starts
under the title and the right side under the epigraph, so
they never overlap; when the epigraph is
written or grows, the right side moves down on the next run.

The font is 14 pt. When the items do not all fit, the script
drops to 13 or 12 pt, the largest size at which they do, and
logs it (`table of contents: font 12 pt`). If even 12 pt is
too big, the last box runs past the bottom of the slide:
tell the author, and suggest leaving minor items out with
`"-"` labels.

Because the script paints those colours, a fill is not how
the author freezes the contents. If the log says `left
alone, you recoloured ...`, the author gave a contents box
another colour or cleared its fill and has taken over the
contents. Tell them; do not work around it.

### Short labels

Every item is bold blue and bulleted, and must fit on one
line and carry no links or
details. The log prints each headline and the item it
became (`item  <=  headline`). A headline without a label is
cut at a word to fit, which often reads badly, so write a
label for every long or messy one and pass them:

```bash
cat > /tmp/toc.json <<'EOF'
{"toc": {
  "English - https://lmarena.ai/leaderboard/text": "Crowd-sourced LM Arena",
  "Coding - https://lmarena.ai/leaderboard/text/coding": "Crowd-sourced LM Arena",
  "OpenAI shelves its 2026 IPO as Anthropic heads to Nasdaq": "OpenAI Delays IPO, Anthropic to Nasdaq",
  "Tech Layoffs by year (US only):": "Jobs & Layoffs",
  "Sept 10": "-"
}}
EOF
python3 g3_update_toc.py --json /tmp/toc.json
```

Rules for labels, as in the archive decks:

- Keys are the headlines exactly as the log prints them.
- About 40 characters or fewer, Title Case, the concrete
  thing: `OpenAI GPT-6 Astra`, `DeepSeek V4.1 Flash Release`.
- Give several boxes of one slide the same label to list
  them once (benchmarks, layoffs).
- `"-"` leaves out a line that is not a topic, such as a
  date or a sub-heading.

Labels are appended to the contents slide's speaker notes,
so later plain runs keep them. Only new headlines need one.
The same JSON file can also carry `"epigraph"`.

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

For a list of 30 candidates to choose from, written to
`work/epigraphs.txt`, use the `gslides-epigraphs` skill.

## Report

Say how many headlines are listed, whether the contents or
epigraph were frozen by the author, and give the epigraph
you wrote, if any.


