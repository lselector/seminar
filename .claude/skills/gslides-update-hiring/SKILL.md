---
name: gslides-update-hiring
description: Refresh the Tech Hiring slide of the live Google Slides seminar deck - software job postings statistics from Indeed via FRED (read by the script), plus the skills and qualifications in demand this month (researched by you) with Dice's Top 50 Skills table. Draws the slide after Jobs and Layoffs if the deck has none. Use when asked to update the hiring, jobs demand, job postings or skills-in-demand slide in the Google Slides deck.
---

# Update the Tech Hiring slide

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

The slide sits right after Jobs and Layoffs and has two
script topics:

| Topic | Words | Picture |
|---|---|---|
| Software job postings on Indeed | the script, from Indeed's postings indexes (FRED) and Indeed's AI share of postings | FRED chart: software vs all postings |
| Skills in demand | you, as JSON (step 2 below) | Dice's "Top 50 Skills" table, current month |

If the deck has no Tech Hiring slide, the step draws it
from `config/skeleton.json` first and says so.
`g1_new_deck.py` carries last week's slide into each new
deck, like the layoffs page.

## Step 1: research the skills in demand

Read the current Dice Tech Jobs Report (WebFetch):
https://www.dice.com/hiring/recruitment/reports/dice-tech-job-report

Take its report month, the year-over-year change in tech
postings, AI/ML postings growth, the fastest-growing skills
and job titles. Add one or two facts on qualifications
(seniority, experience, degrees, AI skills) from Indeed
Hiring Lab (https://www.hiringlab.org) or another named,
dated source found by WebSearch. Keep every number the
source gives. Do not use a source older than about three
months.

## Step 2: write the JSON

`/tmp/hiring-<date>.json`:

```json
{"skills": {
  "bullets": [
    "Dice, 7M+ US tech postings (Aug 2026): **+18%** year over year",
    "AI and machine learning postings ==+101%== year over year",
    "Fastest-growing skills, each 200%+: responsible AI, AI agents, agentic AI",
    "Indeed: 71% of the software postings rebound is senior roles",
    "https://hiringlab.indeed.com/2026/07/08/ai-and-job-postings-from-destruction-to-creation/"
  ],
  "url": "https://www.dice.com/hiring/recruitment/reports/dice-tech-job-report"
}}
```

- 4 to 6 bullets, one fact each, 60 to 90 words in total.
- `**bold**` for the key fact, `==highlight==` for a number.
- A second source goes in as its own bare URL line. `url`
  is the main source and becomes the last line.
- The headline stays "Skills in demand" unless the JSON has
  a `headline`.
- `image` is optional: a picture spec (`{"shot": url}`,
  `{"src": url}`) replacing the Dice table.
- Follow the `humanize` skill: plain sentences, no hype, no
  em-dashes.

## Step 3: run it

```bash
python3 g10_update_hiring.py --json /tmp/hiring-<date>.json
```

Without `--json` only the job postings topic is refreshed;
the skills words and picture always change together.

| Log line | Meaning |
|---|---|
| `no 'Tech Hiring' slide; drawing it` | The slide was missing and is added after Jobs and Layoffs |
| `job postings: not read` | FRED failed; that topic stays as it is |
| `Dice table not found; old picture stays` | The report page changed; the words are still written |
| `left alone (filled by you)` | The author accepted that box; tell them rather than forcing it |
| `the slide is gone or parked` | Nothing was written |

## Step 4: report

The deck name, the numbers now on the slide, the skills
bullets, and anything that stayed old. If the slide was
drawn, suggest `/gslides-update-toc` so the contents lists
it.
