---
name: gslides-update-aa-index
description: Refresh the Artificial Analysis Intelligence Index slide of the live Google Slides seminar deck with a new screenshot of the Cost per Intelligence Index Task chart, today's date in its date box, and, when asked, new bullets about the current leaders. Use when asked to update the Intelligence Index or Artificial Analysis slide in the Google Slides deck.
---

# Update the Intelligence Index slide

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

## Chart and date

```bash
python3 g5_update_aa_index.py          # or: ... 2026-09-25
```

Screenshots only the "Cost per Intelligence Index Task"
chart on
https://artificialanalysis.ai/evaluations/artificial-analysis-intelligence-index
and swaps it into the picture frame. A box on the slide that
holds only a date (`Sept 10`) is set to today's date in the
same style.

Page 3 is the author's own design: the title, a notes box
with the page link, a date box and the chart marked
`auto: aa-index-chart`. `g1_new_deck.py` gives every new
deck a copy of last week's page 3, mark included, so this
is the slide the step normally finds; its chart and date
are last week's until this step runs.

The slide is found by its title, so it works on the
author's hand-made slide too. There it changes only the date
box and the one picture whose alt text description contains
`auto: aa-index-chart`, even if they are filled. Pictures
without that mark are never replaced.

| Log line | Meaning |
|---|---|
| `no Intelligence Index slide in the talk` | The title was changed or the slide was parked |
| `no picture has the alt text 'auto: aa-index-chart'` | Tell the author to add that mark to the chart picture (right-click, Alt text, Description) |

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

Bullets go only into the script's own text box. On a
hand-made slide they are not written.

If the log says `left alone (filled by you)`, the author
has accepted this slide. Tell them rather than forcing it.
