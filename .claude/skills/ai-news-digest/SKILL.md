---
name: ai-news-digest
description: Extract the last N days of AI/tech newsletters from the local Postbox mail folder and write a deduplicated markdown digest of the most important events, each with a short description and 1-2 source URLs. Use when the user asks for an AI news summary, a newsletter digest, "what happened in AI this week", or to summarize their ___Medium_Twitter email folder.
---

# AI news digest from Postbox newsletters

Turn a week of AI newsletters sitting in Postbox into one
ranked, deduplicated list of events.

The mail folder holds roughly 10 newsletters per day from
TLDR AI, AlphaSignal, The Rundown AI, AINews, AI Breakfast,
Medium Daily Digest and others. They cover the same stories,
so **deduplication is the main job**, not summarization.

## Step 1: extract the messages

```bash
python3 .claude/skills/ai-news-digest/tools/extract_emails.py \
    --days 7 --out /tmp/ai_news_extract.txt
```

Run it from the repository root. Defaults point at the
`___Medium_Twitter` folder of the `lev.selector@gmail.com`
account, so normally no other flags are needed.

The script prints one JSON line like this:

```json
{"count": 68, "first_date": "2026-08-26",
 "last_date": "2026-09-02", "out": "/tmp/ai_news_extract.txt",
 "chars": 553580}
```

Keep `first_date` and `last_date`. They name the output file
in step 4. If `count` is 0, stop and report the error rather
than inventing a digest.

Useful flags: `--days N` for a different window, `--mbox
PATH` for another folder, `--max-chars N` for longer bodies.

## Step 2: read the extract

A 7-day window is about 6,300 lines, far more than one
`Read` returns. `Read` caps at 25,000 tokens per call and
this content runs ~25 tokens per line, so use `offset` and
`limit` with **800 lines per chunk**, about 8 calls. Larger
chunks fail outright and waste a turn. Do not stop after
the first chunk: the last day of news is at the bottom of
the file.

Each message appears as:

```
=== MESSAGE 8 ===
DATE: 2026-08-27 13:44 UTC
FROM: TLDR AI <dan@tldrnewsletter.com>
SUBJECT: GLM-5.3 Flash, Claudeforce, NVIDIA revenue surge
--- TEXT ---
...body as plain text...
--- LINKS ---
- anchor text | https://...
```

As you read, keep a running list of candidate events rather
than trying to hold whole newsletters in mind.

## Step 3: build the deduplicated event list

**Merge aggressively.** One real-world event gets exactly one
entry, no matter how many newsletters covered it. A model
launch reported by five newsletters is one event, not five.
Merge on the underlying fact (same company, same release,
same funding round), not on matching headline wording.

**Rank by importance.** Signals, strongest first:

1. How many separate newsletters covered it. This is the
   best available proxy for significance, so count it.
2. Whether it changes what practitioners can do: model
   releases, API and pricing changes, major open weights.
3. Money and structural moves: large funding rounds,
   acquisitions, regulation that takes effect.
4. Research results with named benchmarks or artifacts.

**Drop the noise.** Exclude sponsor blocks and ads, "join our
webinar" items, generic listicles and personal-brand Medium
posts, job listings, and repeat coverage of events older than
the window.

**Pick 1-2 URLs per event.** Prefer the primary source (the
company blog, paper, or repo) over a newsletter's rewrite. If
you only have a newsletter link, use it. Never invent a URL:
use only links present in the extract, and if an event has
none, write "No primary source link appeared in this week's
emails" instead of guessing.

TLDR AI and Substack senders yield the best primary links,
because the extractor decodes their tracking wrappers back
to the real destination. beehiiv and AlphaSignal wrappers
are opaque and get dropped, so events covered only by those
newsletters often have no citable URL. That is expected.

Aim for 15-25 events. Fewer means you merged too hard; many
more means sponsor content leaked in.

## Step 4: write the digest

Save into the **`data` subdirectory** of the repository root
as:

```
data/ai_news_<first_date>_to_<last_date>.md
```

for example `data/ai_news_2026-08-26_to_2026-09-02.md`, using
the dates from the step 1 JSON. Create `data/` if it does not
exist yet.

Structure:

```markdown
# AI news digest: 2026-08-26 to 2026-09-02

Compiled from 68 newsletter emails across 21 senders.

## 1. <Event headline>

<Two or three sentences: what happened, who did it, why it
matters. Include concrete numbers where the sources give
them.>

Covered by 5 newsletters.
Sources: [Company blog](https://...), [TLDR AI](https://...)

## 2. <Next event>
...
```

Order sections by importance, most significant first. Group
under `##` themes (Models & releases, Funding & business,
Policy & regulation, Research, Tools) only if that genuinely
helps; a flat ranked list is usually clearer.

## Writing style

The descriptions are prose, so follow the `humanize` skill at
`.claude/skills/humanize/SKILL.md`. Write plainly from the
first draft. State what happened and what it changes. Skip
hype adjectives, and do not pad an event to fill space when
one sentence covers it.

Stay faithful to the sources. If newsletters disagree on a
number, give the range or attribute it. If something is a
rumor or unconfirmed, label it.

## Notes

- Reading is read-only and safe while Postbox is running. A
  live folder can end mid-message; that block is skipped.
- Never parse these files with Python's `mailbox.mbox`. It
  splits on any line starting with `From `, which shreds 288
  real messages into 463 fragments. The script splits on the
  Postbox separator `From - <date>` instead.
- The window is measured back from the newest message in the
  folder, not from today, so a stale mailbox still returns a
  full week of news.
