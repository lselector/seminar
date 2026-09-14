---
name: gslides-update-benchmarks
description: Refresh the Benchmarks slide of the live Google Slides seminar deck with the latest LM Arena leaderboards (English and Coding), only when newer votes are available, keeping any column the author filled. Use when asked to update the benchmarks, leaderboard or Arena slide in the Google Slides deck.
---

# Update the benchmarks slide

Run from `work/`. The rules are in `work/README.md`.

```bash
python3 g4_update_bench.py          # or: g4_update_bench.py 2026-09-25
```

What it does: runs `sources/leaderboard.py`, compares the vote cutoff
with the one on the slide, and rebuilds the legend, cutoff
note and two columns only when the data is newer.

| Log line | Meaning |
|---|---|
| `already current` | Same cutoff as the slide. Nothing changed |
| `left alone (filled by you): s-bench-k1` | The author froze that column |
| `updated (N requests)` | Rebuilt |
| `sources/leaderboard.py failed; using the cached JSON` | The Arena site did not answer. Say so |

Add `--force` only when the user asks to redraw the page
anyway, and `--no-fetch` to reuse the last download.

Report the cutoff date now shown, and anything left alone.
