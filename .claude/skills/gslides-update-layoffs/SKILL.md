---
name: gslides-update-layoffs
description: Refresh the Jobs and Layoffs slide of the live Google Slides seminar deck with current totals from layoffs.fyi and TrueUp and new screenshots, updating each of the two topics independently and never touching one the author filled. Use when asked to update the layoffs or jobs slide in the Google Slides deck.
---

# Update the layoffs slide

Run from `work/`. The rules are in `work/README.md`.

## Step 1: read the numbers

Open both pages (WebFetch):

- https://layoffs.fyi : tech layoffs per year, US only
- https://trueup.io/layoffs : people laid off this year and
  the last two, with the per-day rate

If a page will not load, leave its part out of the JSON. The
script then only retakes that topic's picture.

## Step 2: write the JSON

```json
{"layoffs_fyi": {"bullets": [
   "Tech layoffs by year, US only",
   "128.5K in 2026, 124K in 2025, 153K in 2024, 264K in 2023"]},
 "trueup": {"bullets": [
   "In 2026: 187,160 people laid off (737 per day)",
   "In 2025: 245,953 people laid off (674 per day)",
   "In 2024: 238,461 people laid off (653 per day)"]}}
```

Same wording pattern as the archive, exact numbers, no
commentary. The source link is added as the last line.

## Step 3: run it

```bash
python3 g7_update_layoffs.py --json /tmp/lay.json
```

| Log line | Meaning |
|---|---|
| `left alone (filled by you): t-trueup-b` | The author froze that topic. The other still updates |
| `is behind a bot check; not shooting it` | TrueUp's Cloudflare wall. The old picture stays |
| `picture ...: not available` | The page did not load. The old picture stays |

Report the numbers written, and which picture could not be
refreshed. For TrueUp, suggest the author replace the
picture by hand if the old one is out of date.
