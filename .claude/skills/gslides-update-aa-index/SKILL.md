---
name: gslides-update-aa-index
description: Refresh the Artificial Analysis Intelligence Index slide of the live Google Slides seminar deck with a new screenshot and, when asked, new bullets about the current leaders, leaving it alone if the author filled the box. Use when asked to update the Intelligence Index or Artificial Analysis slide in the Google Slides deck.
---

# Update the Intelligence Index slide

Run from `work/`. The rules are in `work/README.md`.

## Picture only

```bash
python3 g5_update_aa_index.py          # or: ... 2026-09-25
```

Retakes the screenshot of
https://artificialanalysis.ai/evaluations/artificial-analysis-intelligence-index
and swaps it into the frame.

## With new bullets

Read that page (WebFetch). Write two to four factual
bullets: the index version, the top models and their
scores, anything that changed this week. Keep the numbers.
Follow the `humanize` skill.

```json
{"bullets": [
  "Index v4.3 combines ten evaluations into one score",
  "Claude Fable 5 leads at **73**, ahead of GPT-6 Astra at 71",
  "Useful for comparing **cost per unit of intelligence**"
]}
```

```bash
python3 g5_update_aa_index.py --json /tmp/aa.json
```

The page link is kept as the last line automatically. The
headline stays "Artificial Analysis Intelligence Index"
unless the JSON has a `headline`.

If the log says `left alone (filled by you)`, the author
has accepted this slide. Tell them rather than forcing it.
