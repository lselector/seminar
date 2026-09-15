---
name: gslides-update-benchmarks
description: Refresh the Benchmarks slide of the live Google Slides seminar deck with the latest LM Arena leaderboards (English and Coding), only when newer votes are available, keeping any column the author filled. Use when asked to update the benchmarks, leaderboard or Arena slide in the Google Slides deck.
---

# Update the benchmarks slide

Run from `work/`. The rules are in `work/README.md`.

```bash
python3 g4_update_bench.py          # or: g4_update_bench.py 2026-09-25
```

What it does: runs `sources/leaderboard.py`, which reads the
top 25 of each board and the vote cutoff date the Arena
site publishes, then updates the slide.

- **Script slide** (`s-bench`): rebuilds the legend, cutoff
  note and two columns only when the cutoff is newer.
- **Hand-made slide** with two tables headed Code | Model |
  Score: fills each table in place, matched by the caption
  above it (English, Coding). A changed row gets the new
  model name and link, the vendor colour in the Code cell,
  and the score. The box holding a date ("Data for
  Sept 02") gets the cutoff date. Nothing else is touched.

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
