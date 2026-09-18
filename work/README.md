# Building the weekly seminar deck

The Google Slides deck is the source of truth. There is no
Markdown file and no PPTX in this workflow. Scripts read the
live deck, change only what they own, and leave everything
else alone, so you can edit the deck by hand at any time,
before, between or after any script run.

This page is how to use it. The design, the measurements
behind it and the reasons are in `ADD.md`.

## The three rules you need to know

**1. A fill freezes a box.** Scripts draw text boxes with a
red border and no background. Give a box any background
colour and no script will touch it again, or its picture.
This is how you accept a topic: edit it if you like, then
fill it yellow.

**2. The separator slide is a wall.** Slide 10 says "Not in
the presentation". Nothing after it is ever written. To drop
a topic, copy its box onto the "Parked topics" slide and
delete the original. It will never be added again, even if
next week's news words it differently.

**3. Your own boxes are always yours.** Anything you add by
hand, including a copy of a script box, belongs to you. No
script changes it or removes it. Four exceptions, by
request, even if you made the shapes:

- Intelligence Index slide: step 5 replaces the chart
  picture and the date box (`Sept 10`).
- Benchmarks slide: step 4 fills the two Code | Model |
  Score tables and the date box (`Data for Sept 02`) with
  the latest LM Arena data and its vote cutoff date.
- Jobs and Layoffs slide: step 7 updates the numbers and
  date in the TrueUp and layoffs.fyi boxes and replaces the
  two chart pictures.
- YouTube slide: step 6 rewrites the one line holding
  `subscribers`, in whichever box on that slide holds it,
  filled or not. The rest of that box is left alone, and if
  the promo box is gone the step draws it again from
  `config/skeleton.json`.

**Pictures you add are never deleted or replaced**, on any
slide. A script replaces a picture you placed only when its
alt text description names it as a chart (right-click the
picture, Alt text, Description):

| Description contains | Refreshed by |
|---|---|
| `auto: aa-index-chart` | step 5, the Cost per Intelligence Index Task chart |
| `auto: trueup-chart` | step 7, TrueUp's Tech Employees Impacted chart |
| `auto: layoffs-fyi-chart` | step 7, layoffs.fyi's Recent Tech Layoffs chart |

The marks travel with a slide copied into next week's deck.
To keep a marked picture as it is, delete the mark.

That is all. Everything else below is detail.

## One-time setup

Already done on this machine. For another one:

    python3 g_auth.py <FOLDER_ID>

It opens a browser to sign in and saves
`credentials/token.json`, which git ignores. The folder id
and the source URLs live in `config/gslides.json`.

## The weekly routine

Run these from `work/`. Every command takes an optional date
(`2026-09-18`). Without one it works on this week's seminar:
the coming Friday, which on a Friday is today until 3 pm US
Eastern time, when the seminar is over. From 3 pm on Friday
every step moves to the next week's deck, so Friday evening
work lands in the next deck. Pass a date to reach any other
deck, for example today's after 3 pm.

| Step | Command | Skill |
|---|---|---|
| 1. Create the deck | `python3 g1_new_deck.py` | |
| 2. Add news | `python3 g2_add_news.py --json news.json` | `/gslides-add-news` |
| 3. Contents and epigraph | `python3 g3_update_toc.py --json toc.json` | `/gslides-update-toc` |
| 4. Benchmarks | `python3 g4_update_bench.py` | `/gslides-update-benchmarks` |
| 5. Intelligence Index | `python3 g5_update_aa_index.py` | `/gslides-update-aa-index` |
| 6. YouTube counts | `python3 g6_update_youtube.py` | `/gslides-update-youtube` |
| 7. Layoffs | `python3 g7_update_layoffs.py --json lay.json` | `/gslides-update-layoffs` |
| Before presenting | `python3 g8_preflight.py` | |
| Deck text, epigraph ideas | `python3 g9_deck_text.py --out deck.md` | `/gslides-epigraphs` |

Steps 2 to 7 can run in any order, as often as you like.
Running one twice either refreshes it or does nothing. Steps
that need judgement (finding news, reading numbers off a web
page) are skills: Claude does the reading and hands the
result to the script as JSON.

Every text box a script writes is as tall as its text.
Slides' "resize shape to fit text" cannot be switched on
through the API (probed 2026-09-14: only `NONE` is
accepted), so the scripts measure the text and set the
height themselves. A box a script creates or rewrites gets
exactly its text's height. A box a script only edits (a
number, a date, one line) grows or shrinks by what the edit
did to the text, so the padding you set by hand stays.

The contents on slide 1 is a bold blue bulleted list in two
columns, one short line per topic, no links. Step 3 takes
short labels in `toc.json` (`{"toc": {headline: label}}`,
`"-"` leaves a line out) and keeps them in that slide's
speaker notes; a headline with no label is cut to fit.

## The deck

| # | Slide | Filled by |
|---|---|---|
| 1 | Title, contents, epigraph | Step 3 |
| 2 | Benchmarks: a copy of last week's page 2 (see below) | Step 4 |
| 3 | Artificial Analysis Intelligence Index: a copy of last week's page 3 (see below) | Step 5 |
| 4 | AI News (placeholder) | Step 2 |
| 5 | Weekly videos every Friday: a copy of last week's page 5 (see below) | Step 6 |
| 6 | AI News (placeholder) | Step 2 |
| 7 | Jobs and Layoffs: a copy of last week's page 7 (see below) | Step 7 |
| 8 | About the Speaker | fixed |
| 9 | Thank You! | fixed |
| 10 | Not in the presentation | never touched |
| 11 | Parked topics | never touched |

Pages 2, 3, 5 and 7 are your own designs. Page 2 has two
Code | Model | Score tables, the captions, legend, Elo note
and model sizes; page 3 has your notes, the date box and the
chart marked `auto: aa-index-chart`; page 5 has your promo
box and channel screenshots; page 7 has your TrueUp and
layoffs.fyi boxes and the two charts marked
`auto: trueup-chart` and `auto: layoffs-fyi-chart`. Step 1
makes each new deck as a Drive copy of the latest earlier
deck, deletes every slide but those four, and draws the rest
around them (the Slides API cannot draw a table that
compact). So whatever you change on those pages this week
carries into next week's deck, alt text marks included.
Steps 4 to 7 then put this week's numbers, dates and charts
in; on page 5 that is the counts line only, and your
screenshots stay as you placed them. Only a deck with no
earlier deck to copy gets the script's plain versions.

News goes onto slides 4 and 6 first. Further news slides are
inserted just before Jobs and Layoffs, wherever that slide
is at the time. If you put anything of your own on slide 4 or
6, that slide is no longer used for news.

## What each step does to your edits

| You did this | What scripts do |
|---|---|
| Filled a box | Never change it or its picture, apart from the four by-request exceptions above |
| Typed in an unfilled script box | May overwrite it on the next run. Fill it to keep it |
| Added your own box or slide | Never touch it |
| Copied a script box | The copy is yours |
| Moved or resized a script box | Keep its position and width. Its height follows its text when rewritten, growing only into free space |
| Reordered slides | Follow the new order: the contents list it, news goes before layoffs wherever it is |
| Parked a topic | Never add it again |
| Deleted a topic outright | Never add it again (the ledger remembers) |
| Filled one contents column | Leave the whole contents alone |
| Wrote your own epigraph | Keep it |
| Deleted a fixed slide | That step reports it and does nothing |

The **topic ledger** lives in the speaker notes of the
separator slide. Step 2 adds a line for every topic it
places. You can read it; leave it in place.

## Pictures

Google fetches every picture from a web address. The scripts
take the screenshot on this machine, upload it to the deck's
Drive folder for a few seconds with link sharing on, insert
it, then delete the upload. The picture stays in the deck
because Slides keeps its own copy. Nothing has to be pushed
to GitHub.

TrueUp puts a Cloudflare check in front of automated
browsers. When the check is there, Step 7 keeps the old
picture instead of inserting the "security verification"
page. Replace that picture by hand if it matters.

## Before you present

    python3 g8_preflight.py

It changes nothing and lists what still wants a look: script
boxes you have not filled yet, placeholders still showing, a
contents list that no longer matches the slides, boxes whose
text probably overflows, topics without a picture, and
benchmarks more than a week old. "Ready to present." means
the list is empty.

## Finding a deck

Each deck is tagged with its date in Drive. You can rename it
or move it to another folder and every step still finds it.
**A copy you make of the whole deck is invisible to the
scripts**, because they are only allowed to see files they
created. Keep working in the deck Step 1 made.

## When something goes wrong

| Message | What to do |
|---|---|
| `no deck for 2026-09-18` | Run `python3 g1_new_deck.py 2026-09-18` |
| `A deck for ... already exists` | Step 1 never replaces a deck. Open the link it prints |
| `the deck kept changing` | You were typing while it wrote. Nothing was half-written; run it again |
| `left alone (filled by you)` | Working as intended. Clear the fill if you want the script to update it |
| `slide ... is parked` | The slide is past the separator. Move it back to update it |
| `credentials/token.json not found` | Run `python3 g_auth.py <FOLDER_ID>` |
| Picture unchanged after a run | The page was unreachable or behind a bot check. The log names it |

## How the directory is organised

Run everything from `work/`. The commands are the only Python
files at the top; everything else lives in a directory with
one job and a `README.md` of its own.

```
work/
  g1_new_deck.py ... g9_deck_text.py   the commands you run
  g_auth.py                            one-time Google sign-in
  README.md, ADD.md                    this guide, the design

  gslides/      reads and writes the live deck
  layout/       topics, markup, geometry; no network
  sources/      leaderboards, downloads, screenshots
  config/       gslides.json (folder id, URLs),
                skeleton.json (wording of standing slides)
  assets/       pictures a new deck starts with
  data/         generated files: leaderboard.json, xlsx
  tests/        offline tests and the saved decks they use
  probes/       tools for checking things against Google
  credentials/  OAuth client and token, never committed
```

Dependencies point one way: the commands use `gslides`,
`gslides` uses `sources` and `layout`, `sources` uses
`layout`, and `layout` uses nothing.

Tests: `python3 -m tests` runs them all, or
`python3 -m tests.test_gdeck` runs one file.

`data/images` and `data/images_raw` hold pictures from the
earlier Markdown decks; nothing reads them. `ADD.md`
(Appendix E) lists the code that remains from that workflow.
