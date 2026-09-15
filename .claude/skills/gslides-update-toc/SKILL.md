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

### Short labels

The contents is a bold blue bulleted list in two columns.
Every item must fit on one line and carry no links or
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

## Report

Say how many headlines are listed, whether the contents or
epigraph were frozen by the author, and give the epigraph
you wrote, if any.
