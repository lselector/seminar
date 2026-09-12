---
name: slides-update
description: Create or update the weekly AI News seminar deck by editing its Markdown file and regenerating the PPTX. Resolves the seminar date (next Friday, or today when today is Friday) unless a date is given, creates YYYY-MM-DD-AI-News.md from a skeleton when it does not exist, adds or revises news items, then builds YYYY-MM-DD-AI-News-generated.pptx. Never changes text the author locked, and never re-adds a topic the author moved into YYYY-MM-DD-AI-News-deleted.md. Use when asked to make, update, refresh, or rebuild the weekly seminar slides, this week's AI news deck, or Friday's slides.
---

# Update the weekly seminar deck

One Markdown file per week is the source of truth. The PPTX
is a projection built from it. This skill edits the
Markdown and regenerates the PPTX, without ever undoing
what the human author did by hand.

Read `work/ADD.md` first if you have not: it holds the
grammar, the layout rules, and the reasoning.

## The rule that matters most

The author owns this file. You are a guest in it.

A human makes three kinds of edit. Each is recorded, so
each can be respected:

| Human did | How it is recorded | What you must do |
|---|---|---|
| **Added** a topic | it is simply there | Keep it. Never remove it |
| **Updated** a topic | `<!-- locked -->` | Never change a single character of it |
| **Deleted** a topic | moved to the deleted file | Never write that topic back into the deck |

`<!-- locked -->` goes on its own line. Under a `###`
heading it covers that news item. Directly under a
`## slide:` heading it covers the whole section.

```markdown
### A story the author rewrote himself
<!-- locked -->
- his wording, not yours
```

An item with no marker is yours to revise. Everything else
is not.

### Deletion happens in a separate file

Each deck has a sidecar beside it:

    work/2026-09-18-AI-News.md            the deck
    work/2026-09-18-AI-News-deleted.md    thrown out

Every `###` heading in the sidecar is a topic that must
never appear in the deck again.

The author has two ways to put one there, and both are
normal:

- cut the item out of the deck and paste it into the sidecar
- mark it in the deck with `<!-- deleted -->` and run
  `python3 move_deleted.py` from `work`, which sweeps every
  marked item across

So a `<!-- deleted -->` flag you find in the deck means the
author marked it and has not swept yet. Treat it exactly
like a sidecar entry: leave it alone, and never revive it.
If the sweep is clearly what they want, offer to run
`move_deleted.py --dry-run` first.

When the user asks you to drop a story, mark it
`<!-- deleted -->` and run `move_deleted.py`. That is the
one case where removing lines from the deck is correct,
because the script puts them in the sidecar rather than
throwing them away. Never simply delete an item's lines
yourself.

**Read the sidecar before you add any news.** A story from
the digest that matches something in there is not a
candidate, no matter how important it looks. Matching is on
the topic, not the exact wording, so rephrasing a headline
does not make it a new story.

## Step 1: work out which deck

```bash
python3 .claude/skills/slides-update/tools/deck_init.py --create
```

Run it from the repository root. Add a date first if the
user named one: `deck_init.py 2026-10-02 --create`.

With no date it picks the next Friday, and on a Friday it
picks today rather than a week later. `--create` writes the
deck from `reference/deck_skeleton.md` and the sidecar from
`reference/deleted_skeleton.md` when they do not exist yet,
and never overwrites one that does.

It prints one JSON line. Keep `md`, `pptx`, `deleted`,
`title`, and `created`. If `created` is true you are
starting a fresh deck; if false you are updating an
existing one, so read it before changing anything.

## Step 2: snapshot before you touch it

```bash
cp work/2026-09-18-AI-News.md /tmp/deck_before.md
```

This is what step 5 checks against. Do not skip it. Without
the snapshot there is no way to prove you preserved the
author's work.

## Step 3: gather the news

**Read the deleted file first.** Every `###` heading in
`work/<date>-AI-News-deleted.md` is a topic the author
threw out. Hold that list in mind while picking stories, and
drop any candidate that matches one, however strong it
looks.

For a fresh deck, or when the user asks for this week's
news, use the `ai-news-digest` skill to produce
`data/ai_news_<from>_to_<to>.md`. Pick the strongest
stories from it that are not on the thrown-out list.

If the user supplies stories directly, use those and skip
the digest. Tell them if one of their stories is in the
deleted file rather than quietly adding it.

## Step 3b: write the epigraph

Every deck carries one short line in red at the top right
of the contents page. Write it **after** choosing the
stories, because it should capture what the week was
actually about.

It lives as a blockquote under the deck title:

```markdown
# AI News - Sept 18, 2026

> Agents got their own computers. Oversight is still catching up.
```

What a good one looks like, from the archive:

- First we taught machines to answer. Now we teach them to act.
- Reviewing and validating is the new bottleneck.
- "Manual" work is becoming a legacy.
- The frontier is shifting from larger models to better
  scaffolding around the models we already have.
- Vulkan and Mojo are challenging Nvidia CUDA.

Rules that make them work:

- **Short.** Aim for 70 characters or fewer. The build
  warns if it wraps past two lines.
- **About this week**, not about AI in general. A reader
  should be able to guess the headlines from it.
- **Original.** Write it yourself. If you quote someone,
  attribute them, as the Schwarzenegger line does.
- **One idea, stated flat.** The good ones name a shift
  that just happened. Two short sentences beat one long
  one.
- **No hype.** "Revolutionary", "game-changing" and
  exclamation marks are wrong for this deck.

Offer it to the user rather than assuming it is final. It
is the one line on the deck that is pure editorial voice,
and they may well want their own words.

## Step 4: edit the Markdown

Read the whole file first. Then:

- **Add** new items as `###` blocks under the section they
  belong to. The two `## slide: AI News` sections are
  where news goes.
- **Revise** an unmarked item only when the user asks.
- **Never** touch a `<!-- locked -->` item, or one already
  marked `<!-- deleted -->`.
- **Never** add a topic listed in the deleted file.
- **Never** delete an item's lines yourself. To drop a
  story, and only when asked, mark it `<!-- deleted -->`
  and run `python3 move_deleted.py`.
- **Do not rewrite an epigraph the author changed.** If it
  no longer matches the week, say so and propose a new one
  instead of replacing it silently.

Every news item needs a headline, one image, and bullets:

```markdown
### Headline of the story
![](images/short-slug.jpg)
<!-- shot: https://the-source-page.example.com/article -->
- Each bullet is one short factual statement
- Lead with the number: 88% fewer tokens, $3 billion, 11 days
- **bold** for the key fact, ==highlight== for a number
- https://the-source-page.example.com/article
```

**Bullets are a list, not a paragraph.** Every bullet
renders with a red dot, so each one has to stand on its own
as a single fact. Aim for 5 to 7 bullets of one sentence
each, 80 to 100 words for the item in total.

- One fact per bullet. If a bullet has two, split it.
- Start with the concrete thing: who did what, how much,
  how fast. Not "It is worth noting that Google reports".
- Keep every number the source gives. Numbers are the
  reason these slides are worth reading.
- No bullet should wrap past two lines at 12 pt in the
  6.1 inch column. If it does, it is a paragraph, so cut it
  down or split it.
- The last bullet is usually the source URL, which renders
  as a small blue link rather than a bullet.

Image sources, pick one per item:

| Comment | When |
|---|---|
| `<!-- src: URL -->` | the URL is a real image file |
| `<!-- shot: URL -->` | the source is a web page. This is the usual case |
| none | a local original already in `images_raw/` |

`shot:` is the default choice. Screenshotting the source
page always yields a relevant picture and needs no hunting.

**Write no layout.** No coordinates, no widths, no font
sizes, no slide breaks. A section holds as many items as
the week needs and the generator paginates it. Putting
seven stories under one heading is correct.

Follow the `humanize` skill for the bullet prose. Plain
sentences, real numbers, no hype.

## Step 5: prove you preserved the author's work

```bash
python3 .claude/skills/slides-update/tools/check_preserved.py \
    /tmp/deck_before.md work/2026-09-18-AI-News.md
```

Exit code 0 means every protected item survived. Anything
else lists what broke:

- `CHANGED` you edited a locked item
- `REMOVED` an item that existed is gone
- `RECREATED` you added a topic from the deleted file
- `RESURRECTED` you took an old in-file tombstone off

**Fix every violation before building.** Restore the
original text from `/tmp/deck_before.md`. Do not argue with
the checker and do not build a deck that fails it.

## Step 6: build the PPTX

From the `work` directory:

```bash
cd work
python3 leaderboard.py                       # benchmark page
python3 s1_fetch_images.py 2026-09-18-AI-News.md
python3 s2_clean_images.py
python3 s3_make_pptx.py 2026-09-18-AI-News.md
```

If anything in the deck is marked `<!-- deleted -->`, sweep
it out first, before building:

```bash
python3 move_deleted.py --dry-run
python3 move_deleted.py
```

Skip `leaderboard.py` if it already ran today. Skip `s1`
and `s2` if you added no pictures. `s3` alone is enough for
a text-only change, and it takes under a second.

Output is always `<deck>-generated.pptx`. The hand-made
decks cannot be overwritten: `s3` refuses any target it did
not produce.

## Step 7: report

Tell the user:

- the date and the two file names
- how many slides came out, and how many news items
- what you added, and what you revised
- any missing image, overflowing slide, or failed fetch
- **which items are still unmarked**, and that adding
  `<!-- locked -->` to anything they rewrite will protect it
  from the next run

That last point matters. An unmarked item a human edited
looks exactly like an item you wrote, so the marker is the
only thing that tells the two apart.

## Limits worth stating plainly

The checker protects marked items, items that exist, and
thrown-out topics. It cannot protect an edit the author made
without adding a marker, because nothing records that the
change was theirs. Tell them so, rather than implying a
guarantee that is not there.

Committing the Markdown to git before a run is the wider
safety net, and costs one command.

## Troubleshooting

| Problem | Cause and fix |
|---|---|
| `ERROR in ...: line N` | Grammar mistake. Nothing was written. Fix that line |
| Every image missing | Run `s1_fetch_images.py` then `s2_clean_images.py` |
| Benchmark page empty | `resources_raw/leaderboard.json` is absent. Run `leaderboard.py` |
| `was not produced by this script` | The target is a hand-made deck. Use the `-generated.pptx` name |
| Screenshot is a cookie banner | Replace the file in `images_raw/` by hand, then re-run `s2` |
| Overflowing slides reported | One item has too much text even at 8 pt. Split it into two `###` items |
