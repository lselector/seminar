---
name: gslides-update-layoffs
description: Refresh the Jobs and Layoffs slide of the live Google Slides seminar deck with current totals from layoffs.fyi and TrueUp, today's date, and new screenshots of both layoff charts. Use when asked to update the layoffs or jobs slide in the Google Slides deck.
---

# Update the layoffs slide

Run from `work/`. The rules are in `work/README.md`.

## Hand-made slide (the usual case)

```bash
python3 g7_update_layoffs.py          # or: ... 2026-09-25
```

The script reads the numbers itself and changes four things:

- the TrueUp box: `In 2026: 187,604 people laid off (727 per
  day)` lines get TrueUp's totals
- the layoffs.fyi box: `128.9K in 2026 (as of Sept 14, 2026)`
  lines get layoffs.fyi's yearly totals, rounded the way each
  line already is, and the date becomes today
- two pictures, found only by their alt text description:
  `auto: trueup-chart` gets TrueUp's "Tech Employees
  Impacted by Layoffs" chart, `auto: layoffs-fyi-chart`
  gets layoffs.fyi's "Recent Tech Layoffs" chart. Any other
  picture on the slide is never replaced

A year a source no longer reports (TrueUp currently states
only this year and last) keeps its line unchanged. TrueUp is
read in a real Chrome window parked off-screen, because its
Cloudflare check stops headless Chrome; a Chrome icon may
appear in the Dock for a few seconds.

| Log line | Meaning |
|---|---|
| `layoffs.fyi: not read` / `TrueUp: not read` | That source failed; its lines stay as they are |
| `could not shoot ...` | The chart did not load; the old picture stays |
| `no picture has the alt text 'auto: ...'` | Tell the author to add that mark to the chart picture (right-click, Alt text, Description) |
| `updated (N requests)` | Done |

Report the numbers now shown and anything that stayed old.

## Script-made topics (t-layoffs-fyi, t-trueup)

Only when the slide still has the script's own boxes. Read
both pages (WebFetch) and pass the words as JSON:

```json
{"layoffs_fyi": {"bullets": [
   "Tech layoffs by year, US only",
   "128.9K in 2026, 123K in 2025, 153K in 2024, 266K in 2023"]},
 "trueup": {"bullets": [
   "In 2026: 187,604 people laid off (727 per day)",
   "In 2025: 245,953 people laid off (674 per day)"]}}
```

```bash
python3 g7_update_layoffs.py --json /tmp/lay.json
```

A filled box is frozen: `left alone (filled by you)`.
