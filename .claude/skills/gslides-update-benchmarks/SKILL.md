---
name: gslides-update-benchmarks
description: Refresh the Benchmarks slide of the live Google Slides seminar deck with the latest LM Arena leaderboards (English and Coding), only when newer votes are available, keeping any column the author filled. Use when asked to update the benchmarks, leaderboard or Arena slide in the Google Slides deck.
---

# Update the benchmarks slide

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

```bash
python3 g4_update_bench.py          # or: g4_update_bench.py 2026-09-25
```

What it does: runs `sources/leaderboard.py`, which reads the
top 25 of each board and the vote cutoff date the Arena
site publishes, then updates the slide.

Page 2 is the author's design: title `"LM Arena" Leaderboard`,
a `Data for Sept 13` box, red captions `English - <link>` and
`Coding - <link>`, two tables headed Code | Model | Score,
a colour legend, the `Starting Elo rating = 1000` note and
the list of model sizes. `g1_new_deck.py` gives every new
deck a copy of last week's page 2, so it always looks like
this; its numbers are last week's until this step runs.

- **Table page** (the usual case): fills each table in
  place, matched by the caption above it (English, Coding).
  A changed row gets the new model name and link, the vendor
  colour in the Code cell, and the score. The box holding a
  date gets the cutoff date. Nothing else is touched: the
  Elo note, the model sizes and all styling stay as the
  author left them, so edit those by hand when they change.
- **Script column page** (`s-bench`): only in a deck made
  with no earlier deck to copy. Rebuilds the legend, cutoff
  note and two text columns when the cutoff is newer. It
  cannot be turned into the table design by this step, as
  the Slides API cannot draw a table that compact.

| Log line | Meaning |
|---|---|
| `already current` | Same data as the slide. Nothing changed |
| `no Benchmarks slide in the talk` | No `s-bench` and no Code/Model/Score tables before the separator |
| `left alone (filled by you): s-bench-k1` | The author froze that column |
| `updated (N requests)` | Rebuilt |
| `sources/leaderboard.py failed; using the cached JSON` | The Arena site did not answer. Say so |

Add `--force` only when the user asks to redraw the page
anyway, and `--no-fetch` to reuse the last download.

Report the cutoff date now shown, and anything left alone.
Table row heights are the author's: a row that grew for a
long model name stays tall after that name moves.
