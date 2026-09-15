---
name: gslides-add-news
description: Find this week's most important AI news and add it as new topics to the live Google Slides seminar deck (YYYY-MM-DD-AI-News), without ever re-adding a topic the author parked or deleted, and without touching anything the author filled or wrote. Creates the deck first if it does not exist. Use when asked to add news to the Google Slides deck, fill this week's slides with news, or update the Friday deck in Google Slides.
---

# Add news to the live Google Slides deck

The deck in Google Slides is the source of truth. You gather
and write the news; `work/g2_add_news.py` decides what is new
and places it. Read `work/README.md` once if you have not:
it holds the three rules (a fill freezes a box, nothing past
the separator is written, the author's own boxes are theirs).

Run every command from `work/`.

## Step 1: make sure the deck exists

```bash
python3 g1_new_deck.py            # or: g1_new_deck.py 2026-09-25
```

It creates the deck for the next Friday (today, on a Friday)
or prints the link of the one that already exists. It never
replaces a deck, so it is always safe to run. Use the same
date argument in every later command if the user named one.

## Step 2: see what the deck already knows

```bash
python3 g2_add_news.py --list
```

Every line is a topic already in the talk (`deck`), parked
by the author (`parked`), or added before and possibly
deleted by the author (`ledger`). **None of these is a
candidate**, however it is worded. Hold the list in mind
while choosing stories.

## Step 3: gather the news

If the user gave stories, use those. Otherwise run the
`ai-news-digest` skill for the last 7 days and pick the
strongest stories from its digest that are not on the list
from step 2. Aim for 6 to 8 topics.

## Step 4: write the topics as JSON

Write `/tmp/news-<date>.json`:

```json
{"topics": [
  {"headline": "Mistral raises EUR 3 billion Series D",
   "bullets": [
     "EUR 3 billion led by Samsung, with Nvidia and ASML",
     "Valuation above **EUR 21 billion**",
     "Le Chat is being renamed Vibe"
   ],
   "url": "https://mistral.ai/news/series-d",
   "image": {"shot": "https://mistral.ai/news/series-d"}}
]}
```

Rules for each topic, the same as the archive decks:

- **Headline**: plain, specific, under about 70 characters.
- **Bullets are a list, not a paragraph.** 4 to 7 bullets,
  one fact each, 80 to 100 words for the topic in total.
- Lead with the concrete thing: who did what, how much, how
  fast. Keep every number the source gives.
- `**bold**` for the key fact, `==highlight==` for a number,
  `` `code` `` for commands or model ids. Nothing else.
- `url` is the source page. It is added as the last line, a
  small blue link.
- `image`: `{"shot": url}` screenshots a page, and is the
  usual choice. `{"src": url}` downloads an image file.
  Leave it out and the source page is screenshotted.
- Follow the `humanize` skill: plain sentences, no hype, no
  em-dashes.

## Step 5: dry run, then add

```bash
python3 g2_add_news.py --json /tmp/news-<date>.json --dry-run
python3 g2_add_news.py --json /tmp/news-<date>.json
```

The dry run lists each topic as `new` or `skipped` with the
reason. A skipped topic is not an error: it means the author
already has it, parked it, or deleted it. Do not reword a
skipped topic to get it past the check.

## Step 6: refresh the contents

```bash
python3 g3_update_toc.py
```

Then give each new topic a short one-line label, and write
the epigraph if it is still a placeholder. Both are in the
`gslides-update-toc` skill.

## Step 7: report

Tell the user:

- the deck name and link
- the topics added, and the slides they went on
- the topics skipped, with the reason for each
- that each new box is unfilled until they accept it, and
  that filling it protects it from every later run
- to run `python3 g8_preflight.py` before presenting

## Never

- Never edit the deck through any path other than the
  `g*_` scripts.
- Never delete or rewrite a slide or box yourself.
- Never re-add a topic the dry run skipped.
