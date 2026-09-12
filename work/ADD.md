# Architecture Design Document: Weekly AI News Deck

Status: Draft (pending approval) - Date: 2026-09-12

## Executive summary

Every Friday a new `YYYY-MM-DD-AI-News.pptx` has to
exist, carrying 20 to 40 news items with pictures and
links. Today that deck is built by hand in PowerPoint, and
the work that takes the longest is not thinking about the
news. It is typing, pasting, cropping, dragging boxes
around, and re-typing the same titles into the table of
contents.

This design keeps one text file as the source of truth for
a week's deck. A Markdown file holds the titles, the
bullets, the links, and the names of the images. Four small
Python scripts sit around it: one downloads or screenshots
the pictures the file references, one normalizes them, one
fetches the benchmark numbers, and one builds the PPTX.

**The author never lays out a slide.** There are no
coordinates in the deck file and no way to put one there.
The generator decides how many slides a section needs,
which news items share a slide, how big the text is, and
the exact rectangle for every box and every picture. The
table of contents is generated from the headings, so it
cannot drift out of sync with the slides.

Each piece of news is one headline, a few bullets, and one
illustrative picture. That uniformity is what makes
automatic layout possible: the generator is never guessing
what a slide is supposed to look like.

Measured on the Sept 18 deck: 13 news items across 11
slides, 12 pictures, all text placed at 12 pt with zero
overflow, built in 0.34 seconds.

## Contents

- [The problem](#the-problem)
- [Non-goals](#non-goals)
- [The pipeline](#the-pipeline)
- [The deck file: the one contract](#the-deck-file-the-one-contract)
- [Protecting what a human wrote](#protecting-what-a-human-wrote)
- [Automatic layout](#automatic-layout)
- [The three pages that are not news](#the-three-pages-that-are-not-news)
- [Slide inventory](#slide-inventory)
- [Images](#images)
- [Why the deck is small by construction](#why-the-deck-is-small-by-construction)
- [The weekly run](#the-weekly-run)
- [Verification](#verification)
- [Failure handling](#failure-handling)
- [Known limits and scaling gates](#known-limits-and-scaling-gates)
- [The shape of the thing](#the-shape-of-the-thing)
- [Appendix A: Requirements](#appendix-a-requirements)
- [Appendix B: Technology choices](#appendix-b-technology-choices)
- [Appendix C: Risks](#appendix-c-risks)

## The problem

The archive in `2026/` holds 41 decks from this year alone.
They are good decks. Producing one costs an evening, and
most of that evening goes on mechanics:

- The same headline gets typed twice, once on the news
  slide and once in the table of contents on slide 1. When
  a story is cut late, slide 1 still lists it.
- Every text box and every picture is positioned by hand.
  Slide 14 of the Sept 11 deck has three text boxes and
  three pictures at six different hand-chosen positions.
  None of that work carries forward to next week.
- Images arrive at whatever size and format the source
  served. Some come in with transparent backgrounds that
  turn black, some are 4000 px wide, some are 200 px wide
  and go fuzzy when stretched.
- There is no text version of a deck. Searching last
  month's slides means `grep` against `seminar_ppt.txt`, a
  generated 8 MB text dump.
- Nothing carries forward. Week 2's deck starts as a copy
  of week 1's file, and the stale slides have to be found
  and deleted by hand.

Constraints that shape the design:

- One person builds this, on a Mac, in a few hours a week.
- The output must stay PPTX. Five years of archive are in
  that format, and it is what gets presented.
- Slides are 16:9 landscape, 10 x 5.625 inches.
- No network dependency at presentation time.

## Non-goals

The first version deliberately does not do these:

- It does not write the news. Selecting and describing
  stories is the human's job, assisted by the existing
  `/ai-news-digest` skill.
- It does not accept layout instructions. No coordinates,
  no widths, no slide breaks, no font sizes in the deck
  file. If a slide looks wrong the fix belongs in
  `deck_layout.py`, where it fixes every future deck too.
- It does not reproduce the hand-tuned collages of the
  older decks. It produces one consistent shape.
- It does not produce HTML slides, PDF slides, or video.
- It does not publish anything. Committing to git and
  uploading to YouTube stay manual.

## The pipeline

Seven modules. Each owns one job and talks to the next only
through files on disk or a documented function.

```
 /ai-news-digest skill
        |
        v
 data/ai_news_<from>_to_<to>.md      (ranked pool of events)
        |
        |  human curates and writes prose
        v
 work/YYYY-MM-DD-AI-News.md       <== SOURCE OF TRUTH
        |
        +--> s1_fetch_images.py --> images_raw/
        |         (download or screenshot)   |
        |                          s2_clean_images.py
        |                                    v
        |                                images/
        |
        |    leaderboard.py --> resources_raw/leaderboard.json
        |                                    |
        +--> s3_make_pptx.py <---------------+
                  |
                  |  deck_parser.py   grammar  -> data
                  |  deck_layout.py   data     -> rectangles
                  v
      work/YYYY-MM-DD-AI-News-generated.pptx
                  |
             copy to 2026/
```

Module ownership, public API, and allowed dependencies:

| Module | Owns | Public API | Depends on |
|---|---|---|---|
| `ai-news-digest` (skill) | `data/ai_news_*.md` | the digest file | mail folder |
| deck file (human-authored) | slide content and order | the grammar below | nothing |
| `deck_parser.py` | the grammar | `parse_deck`, `image_manifest`, `all_headlines` | nothing |
| `text_metrics.py` | measuring text | `text_width`, `wrap_lines` | Pillow |
| `deck_layout.py` | all geometry | `paginate(sections)` | `deck_parser`, `text_metrics` |
| `pptx_text.py` | the palette and text boxes | `add_textbox`, `style_run` | python-pptx |
| `bench_page.py` | the benchmarks slide | `render_benchmarks_page` | `deck_layout`, `pptx_text` |
| `s1_fetch_images.py` | `images_raw/` | reads deck file, writes raw files | `deck_parser`, network, Chrome |
| `s2_clean_images.py` | `images/` | reads `images_raw/`, writes `images/` | ImageMagick |
| `leaderboard.py` | benchmark data | XLSX + `leaderboard.json` | network, openpyxl |
| `s3_make_pptx.py` | the `.pptx` | reads deck file, images, JSON | all of the above |
| `move_deleted.py` | moving flagged items to the sidecar | the deck file and its sidecar | `deck_parser` |

Dependencies run one way. `deck_parser` imports nothing of
ours. `deck_layout` imports only `deck_parser`, and knows
nothing about PowerPoint. `s3_make_pptx` is the only module
that touches python-pptx, and it is the composition root:
it reads the files, calls the two pure modules, and writes
the result.

That split is the reason the layout arithmetic can be
tested at all. `deck_layout` returns numbers, so a test can
assert that no rectangle leaves the slide without opening a
PPTX or looking at a picture.

Why separate scripts rather than one program with
subcommands: the stages fail for different reasons and get
re-run at different times. Screenshots fail because a site
is slow. The renderer fails because a heading is
malformed. Separating them means a renderer fix does not
re-shoot 11 web pages.

## The deck file: the one contract

This is the heart of the design. Everything else reads or
writes it.

The grammar is seven rules:

1. `#` at the top is the deck title. It also fills the
   heading of the contents slide.
2. `## slide: <title>` starts a section with that title.
3. `### <headline>` starts a news item inside the current
   section.
4. An image line directly under a `###` is that item's
   picture, optionally followed by a manifest comment.
5. `-` bullets are body lines. A bullet holding only a URL
   renders as a small blue clickable link.
6. `<!-- locked -->`, `<!-- notoc -->` and
   `<!-- profile -->` mark what a human has claimed, what
   stays off the contents, and which slide is the author
   page. Deletion moves an item to a sidecar file. See
   [Protecting what a human wrote](#protecting-what-a-human-wrote).

Note what rule 2 does **not** say. A section is not a
slide. It is a group of news items that share a title, and
the generator turns it into as many slides as the text
needs. Writing seven items under one `## slide:` heading is
normal and correct.

Two section names are reserved and generate their own
content: `## slide: toc` and `## slide: benchmarks`.

Inside a bullet, `**bold**` and `==highlight==` work. The
highlight is the pale yellow used throughout the existing
decks.

Example:

```markdown
# AI News - Sept 18, 2026

## slide: toc

## slide: benchmarks

## slide: AI News

### Meta ships Muse agent with its own cloud computer
![](images/meta-muse.jpg)
<!-- shot: https://about.fb.com/news/ -->
- Handles email, calendars, Instagram and websites
- **Keeps working after you close it**
- Tiers: Free, $20/mo, $100/mo

### Claude formalized Fermat's Last Theorem in 11 days
![](images/fermat-lean.jpg)
<!-- shot: https://www.anthropic.com/research/... -->
- Converted the 1995 proof into ==13 Mln lines of Lean==
- https://www.anthropic.com/research/...
```

### Protecting what a human wrote

The deck file is edited by two parties: the author, and an
assistant running the `slides-update` skill. The author
makes three kinds of edit, and two markers record the two
that need defending.

| Human did | How it is recorded | Effect |
|---|---|---|
| added a topic | it is simply there | nothing may remove it |
| updated a topic | `<!-- locked -->` | no character of it may change |
| deleted a topic | moved to a sidecar file | no slide, and the topic may never return |

A `locked` marker on its own line under a `###` covers that
news item. Directly under a `## slide:` heading it covers
the whole section, which is how the author page and the
thank you page are protected from the moment a deck is
created.

Deletion is the interesting one, because a lock cannot
defend something that is not there. If a story is cut by
removing its lines, the next pass reads the deck, finds the
story missing, finds it again in the news digest, and
helpfully puts it back. The record of the decision has to
live somewhere.

It lives in a sidecar file named after the deck:

    2026-09-18-AI-News.md            the deck
    2026-09-18-AI-News-deleted.md    topics thrown out

The item's `###` heading is the record. There are two ways
to get it into that file, because the two suit different
moments, and both end in the same state:

- **By hand.** Cut the item out of the deck and paste it in.
  One motion, best for a single item while already editing.
- **Mark and sweep.** Put `<!-- deleted -->` under the item
  and run `move_deleted.py`, which cuts every marked item
  across, stamps it with the date, and drops the marker. Best
  when several go at once, because the author never leaves
  the deck file or loses their place.

The second exists because marking is the cheaper gesture
while reading, and a person dropping five stories should not
have to do five cut-and-pastes. The script is the only thing
allowed to remove lines from a deck, and it never destroys
them: text moves verbatim, the previous deck is kept as
`<deck>.bak`, and it refuses to touch a deck that does not
parse. `--dry-run` prints the plan and writes nothing.

The marker also works on a `## slide:` heading, moving that
section and all its items together.

Matching is on the topic, not the string. Exact comparison
looked adequate and is not: a regenerated deck almost never
reproduces a headline word for word, so "The music industry
stops fighting" would sail past a ban on "The music
industry stops fighting and starts licensing". So
`same_topic` treats headlines as one topic when they are
equal after folding case and space, when one contains the
other, or when they share 70% of their significant words.
Two different stories about the same company stay distinct.

Enforcement is not a promise. `check_preserved.py` compares
the deck before and after an editing pass and exits
non-zero on `CHANGED`, `REMOVED`, `RECREATED`, or
`RESURRECTED`. Items are matched by heading text rather
than position, so inserting a story in the middle is not
mistaken for a rewrite. The tool implements its own
splitting of the file rather than importing `deck_parser`,
because a checker that shares code with the thing it checks
shares its bugs. That deliberate duplication now includes
`same_topic`, so a test asserts the two copies agree on a
table of real headlines and cannot drift apart.

The build is the second net. Someone editing by hand at
11:55 will not run the checker, so `s3_make_pptx.py` reads
the sidecar too, reports how many topics are in it, and
warns loudly, naming both headlines, if one is back on a
slide. It warns rather than refuses, because the author may
have deliberately changed their mind, and a build minutes
before a talk should not be blocked by a heuristic.

One limit, stated plainly rather than papered over: an edit
the author makes without adding a marker is
indistinguishable from text the assistant wrote, so nothing
can protect it. The skill is required to say so in its
report instead of implying a guarantee it does not have.

### Image sources

The comment under an image is the download manifest, and it
has three forms:

| Comment | Meaning |
|---|---|
| `<!-- src: URL -->` | download that picture |
| `<!-- shot: URL -->` | screenshot that web page |
| none | a local original, checked but never fetched |

The manifest lives next to the picture it describes, so
there is no second list to fall out of date with the deck.

Two properties this buys, both worth the small amount of
convention. The deck is greppable, diffable, and
reviewable in git, so a week's changes show up as a normal
text diff. And content is fully separated from
presentation, which is what allows the layout to be
someone else's problem.

## Automatic layout

`deck_layout.py` turns sections into placed rectangles. It
runs in four steps.

**1. Estimate.** For each news item, predict its rendered
height from its character count and its column width. A
glyph of Calibri at size S is about `0.50 * S` points wide
for body text and `0.56 * S` for a bold headline, so the
number of wrapped lines follows from the box width, and the
height follows from a line spacing of `1.24 * S`. Those
three constants were measured from the existing decks and
live at the top of the module.

**2. Pack.** Walk the section's items and keep adding them
to the current slide while they still fit the content band
and the slide holds fewer than three. Otherwise start a new
slide. Continuation slides repeat the section title, which
is why the old decks have several slides all called "AI
Updates" and the new ones do too.

**3. Choose one font size per slide.** Try 12, 11, 10, 9,
then 8 pt, and take the first that fits. Every item on a
slide shares the chosen size so the page looks deliberate
rather than patched. An item too tall to share a slide gets
one to itself, stepped down the same ladder.

**4. Distribute the slack.** Whatever vertical space is
left over is shared equally between the items on the
slide, so pages are never top-heavy with a gap at the
bottom.

The geometry, in inches, measured from the existing decks:

```python
SLIDE_W, SLIDE_H = 10.00, 5.625
MARGIN      = 0.09
TITLE_H     = 0.36      # title bar, 20 pt bold
BAND_TOP    = 0.46      # content starts here
BAND_BOT    = 5.53
CONTENT_W   = 9.82
IMG_W       = 3.60      # picture column
TEXT_W_IMG  = 6.10      # text beside a picture
TEXT_W_FULL = 9.82      # text with no picture
MAX_BLOCKS  = 3
FONT_LADDER = [12, 11, 10, 9, 8]
```

Each news item gets text on the left and its picture fitted
into a 3.60 inch column on the right, centered in its box
with the aspect ratio preserved. An item with no picture
gets the full slide width for its text, which is what
happens automatically when an image file is missing.

Typography, also measured from the existing decks:

| Element | Font | Size | Weight |
|---|---|---|---|
| Slide title | Calibri | 20 pt | bold |
| Item headline | Calibri | size + 1 | bold |
| Body bullet | Calibri | 12 to 8 pt | normal |
| Link line | Calibri | 9 pt | normal, blue |

Every content box carries the pale yellow fill `FFF2CC` and
a 0.75 pt red border `FF0000`, both lifted from the
hand-made decks. The slide title is left unfilled and
unbordered, as it is there.

A bordered box has to fit its text, because a border makes
sloppiness visible in a way an invisible box never did. So
the text rectangle is the estimated height of its own text,
not the whole row: the border hugs the words. The picture
still gets the full row, which keeps it as large as the
slide can afford. Each box also carries `spAutoFit`, so
PowerPoint settles the last fraction of an inch against the
real text and corrects any small error in the estimate.

The font ladder is what makes overflow a solved problem
rather than a warning the human has to act on. A build
still reports any slide whose text misses the band even at
8 pt, because silence about clipped text would be worse,
but that is a safety net and not part of the normal
workflow. On the Sept 18 deck it never fired: all 11 slides
landed at 12 pt.

## The three pages that are not news

Three pages recur every week and none of them is a news
item. They are worth writing down because they are the
pages a naive design leaves as manual work.

**Benchmarks: generated from data, not a screenshot.**
`leaderboard.py` already fetches the Arena leaderboards and
writes two XLSX files. It now also writes
`resources_raw/leaderboard.json`, holding the top 25 of
each board with each model's name, score, vendor, and page
URL. `## slide: benchmarks` turns that JSON into two plain
columns, 25 models each, written as `marker score name`
with the marker coloured by vendor and a legend across the
top.

The page also states what the numbers are current to.
Arena publishes a `voteCutoffISOString` beside the
entries, and that, not the day we downloaded the file, is
what the board reflects: on Sept 12 the boards were still
counted through Sept 11. `leaderboard.py` stores it per
board, so if the two ever diverge the page says so instead
of printing one date over both.

Parsed data beats a screenshot here for three reasons worth
the extra code. The text stays crisp at any projector
resolution. The numbers are real text, so they can be
edited or searched. And nothing has to be re-cropped by
hand each week. `leaderboard.py` owns the vendor
classification rules, so the renderer never repeats them.

This page was a real PPTX table first, and that was wrong.
PowerPoint enforces a minimum row height that it applies
regardless of what the file asks for, so 25 rows always
spilled off the bottom of the slide no matter what height
python-pptx wrote. Ordinary paragraphs have no such floor:
line spacing can be set exactly, the font size is computed
from the space available, and the vendor swatch becomes a
coloured character instead of a filled cell. The result
fits by arithmetic rather than by hope, and the column
order puts the score before the name so the ranking reads
down the page.

**Intelligence Index: a screenshot, because it is a live
page.** The Artificial Analysis index is a web page with an
interactive chart, not a data file worth reverse
engineering for one picture a week. So it is an ordinary
news item whose image carries a `shot:` comment, and
headless Chrome captures it. Nothing about it is a special
case in the renderer.

**Jobs and layoffs: two screenshots, same mechanism.**
`layoffs.fyi` and `trueup.io/layoffs` are both `shot:`
entries. The headline numbers go in the bullets, where they
are greppable and easy to update, and the tracker page
itself provides the illustration.

The division of labour generalizes. When a source publishes
numbers we can parse, render them as a table and keep the
data. When it publishes a page, photograph the page. The
`shot:` mechanism means the second case costs one line in
the deck file.

## Slide inventory

The skeleton of a typical week. Slide numbers are
approximate because the generator decides them.

| Slide | Content | How it is produced |
|---|---|---|
| 1 | Table of contents | generated from every `###` heading |
| 2 | LM Arena benchmarks | generated from `leaderboard.json` |
| 3 | Intelligence Index | `shot:` screenshot |
| 4-9 | Top news | deck file, auto-paginated |
| 10 | YouTube channel | `shot:` of the channel page |
| 11-23 | More news | deck file, auto-paginated |
| 24 | Jobs and layoffs | two `shot:` screenshots |
| 25 | About the speaker | deck file, local photo |
| 26 | Thank you | deck file |

The important observation: apart from the contents page and
the benchmark page, every slide is the same thing. A title,
then one to three items, each an headline with bullets and
a picture. The author page, the YouTube page, and the thank
you page are not special cases in code. They are ordinary
news items whose text happens not to change much. That is
what keeps the renderer small and the layout uniform.

The sections whose text is stable week to week can be kept
in a template file and copied forward, so the author page
does not get retyped 52 times a year.

## Images

Two directories, because a fetched file and a
presentation-ready file are different things.

`images_raw/` holds exactly what the network returned,
whether downloaded or screenshotted. Never edited. This is
the "do the expensive work once" boundary: shooting 11 web
pages takes about 40 seconds and can fail because a page
moved, so once a file lands it stays.

`images/` holds normalized copies, and is disposable. It
can be deleted and rebuilt from `images_raw/` in about a
second.

The normalization standard:

| Property | Value | Why |
|---|---|---|
| Format | JPEG, sRGB, quality 85 | PPTX embeds it directly |
| Background | white, alpha flattened | transparent PNGs render black |
| Size | fits the slide's picture box at 150 ppi | see below |
| Aspect ratio | preserved, never padded | see below |
| Trim | 2% fuzz border trim | strips screenshot whitespace |
| Marker | comment `seminar-150ppi-q85` | idempotent, and records the settings |

### Why the deck is small by construction

A PPTX is its pictures. On a representative deck they are
86% of the bytes, and everything else together is under 40
KB. So deck size is an image question, and it is settled
here rather than afterwards.

The old manual fix was to open the finished file in
PowerPoint and run Compress Pictures at "email (96 ppi)",
taking 20 to 30 MB down to about 2 MB. That is a repair,
performed on a file that should never have been that large,
and it is one forgotten click away from shipping 30 MB.

The generator knows exactly how big a picture can ever
appear, because it decides the layout: at most 3.60 by 5.07
inches. At 150 ppi that is 540 by 760 pixels, and any pixel
beyond it cannot be displayed. So `s2_clean_images.py`
imports the box from `deck_layout` and scales to fit it. The
two numbers cannot drift apart, and a test asserts they
match.

"Delete cropped areas" has no equivalent because nothing is
ever cropped in the PPTX. Trimming happens in ImageMagick
before the picture is placed, so there is no hidden margin
to discard.

Measured on the Sept 18 deck, 11 pictures:

| Setting | Deck |
|---|---|
| unbounded, longest side 1600 px | 1336 KB |
| `--ppi 220`, PowerPoint's "print" | 591 KB |
| `--ppi 150`, the default | 338 KB |
| `--ppi 96`, PowerPoint's "email" | 190 KB |

150 ppi is the default because it is four times smaller than
storing 1600 px and still sharper than the 96 ppi the decks
shipped at for years. The resolution is a flag, not a
rewrite, and it is recorded in each file's marker so
changing it rebuilds everything without `--force`.

The scaling gate: if a deck ever passes 8 MB, the build says
so and names the command to shrink it.

Aspect ratio is the one place the earlier design was wrong.
`s2_clean_images.py` used to pad every picture onto a fixed
800x600 white canvas. That is right for Markdown, which
cannot resize a picture, and wrong for slides: the white
bars get baked into the file and then the renderer adds
more of its own. The renderer fits pictures into their box
at placement time, so the job here is only to make them
clean and a sensible size.

Because `images_raw/` is permanent, `s2` needs no backup of
its own. The old version copied the whole directory to
`~/backups` before editing pictures in place, and that
machinery is gone. Removing a component is a win.

## The weekly run

```bash
cd work

# 1. build the ranked pool of events (existing skill)
#    -> data/ai_news_<from>_to_<to>.md

# 2. human writes 2026-09-18-AI-News.md

# 3. refresh the benchmark numbers
python3 leaderboard.py

# 4. fetch and normalize pictures
python3 s1_fetch_images.py 2026-09-18-AI-News.md
python3 s2_clean_images.py

# 5. build the deck
python3 s3_make_pptx.py 2026-09-18-AI-News.md

# 6. read it once, then copy to ../2026/ and commit
```

### Rebuilding minutes before the seminar

Fixing a typo at 11:55 is the case that has to work
without thinking, so the rebuild is one command with no
arguments:

```bash
python3 s3_make_pptx.py
```

With no file named it takes the newest `*-AI-News.md` in
the directory and says which one it picked. It needs no
network, reads only local files, and finishes in about a
third of a second. Three details make it safe to run under
pressure:

- **It writes only `<deck>-generated.pptx`.** The output
  name is never the same as a hand-made deck's name.
- **It refuses to overwrite anything it did not write.**
  Every generated file carries a marker in its core
  properties, and a target that is neither named
  `-generated.pptx` nor carries the marker is left alone
  unless `--force` is given. Five years of hand-made decks
  sit in this repository; none of them can be clobbered by
  a mistyped command.
- **Paths resolve against the Markdown file, not the shell.**
  Running it from the repository root works exactly the
  same as running it from `work`. Otherwise every picture
  would quietly go missing and the deck would still build,
  which is the worst possible outcome minutes before a
  talk.

A grammar mistake stops the build with one line naming the
offending line number, and writes nothing. The previous
PPTX stays on disk, so a bad edit never leaves the author
with no deck at all.

Re-running `s1` and `s2` is only needed when a picture was
added or changed. Text-only edits need step 5 alone.

Step 5 takes under half a second, and steps 4 and 5 are
free to repeat because they skip finished work. Editing one
bullet and rebuilding is instant, which is the point: the
deck stops being a thing you are afraid to touch on
Thursday night.

Lifecycle, the part that is easy to forget. When a news
item is cut from the deck file, three things follow. Its
contents entry disappears and the slides reflow, both
automatically. Its picture becomes an orphan in `images/`,
which is not automatic: `s2` reports unreferenced files
rather than deleting them, because deleting a picture that
turns out to be needed next week costs more than a line of
log.

## Verification

Small project, so the pyramid is short. It still has every
layer, because the failure modes differ at each one.

- **Unit and module.** `test_deck.py` covers the grammar
  and the layout arithmetic: 34 tests, no dependencies, run
  with `python3 test_deck.py`. The layout tests are the
  ones that matter. Since nobody positions anything by
  hand, a packing bug would push text off a slide silently,
  so the tests assert directly that no rectangle leaves the
  content band and none leaves the slide.
- **Integration.** Build the real deck file and reopen the
  PPTX: assert the slide count, that no shape crosses a
  slide edge, and that no text box overlaps a picture.
- **Conformance.** The project's Python rules are
  mechanical, so check them mechanically: files under 800
  lines, functions under 35 lines, lines under 65
  characters, a docstring on every function and class.

The check that matters most cannot be automated. Whether a
slide reads well from the back of a room is a judgment, and
the honest answer is that the human looks at the deck once
before presenting. What automation buys is that the
judgment is now about the words instead of about whether a
text box is 0.2 inches too low.

## Failure handling

What repairs itself:

- A failed download or screenshot is logged and skipped,
  and the item renders with text at full slide width. One
  dead URL does not stop a build.
- A partially written file is discarded, because both `s1`
  and `s2` write to a temporary name and rename only on
  success.
- `images/` is disposable and regenerates from
  `images_raw/`.
- Re-runs are idempotent. Every stage detects finished work
  and skips it.
- A missing `leaderboard.json` warns and leaves the
  benchmark page empty rather than failing the build.

What is never repaired automatically, and why:

- **The deck file.** If a heading is malformed the parser
  stops and names the line. Guessing what the author meant
  would produce a wrong slide that looks right.
- **`images_raw/` and `resources_raw/`.** Originals.
  Nothing deletes them, including the orphan report.
- **Committed decks in `2026/`.** The published record. No
  script writes there; the copy is a human action.
- **Any PPTX this pipeline did not produce.** The renderer
  checks the target before writing and stops rather than
  overwrite a hand-made deck. See
  [Rebuilding minutes before the seminar](#rebuilding-minutes-before-the-seminar).

Retry manners: `s1` tries a download twice with a short
backoff, then gives up and logs. No infinite retry, because
a missing picture is a cosmetic problem and the build
should still finish.

## Known limits and scaling gates

Three simplifications, each with the measurement that would
change the decision.

**Almost every slide has the same shape.** One to three
items, text left, picture right. The one exception is the
author page, which carries a `profile` flag and gets the
mirrored shape from the older decks: a large portrait on
the left, details beside it, the name set larger. This is
the scaling gate below being spent once, deliberately: the
deck file names a *kind* of slide and `deck_layout` owns
what that kind looks like, so there are still no
coordinates anywhere in the Markdown. The hand-made collages in the
older decks are more varied and sometimes more striking.
This is the deliberate trade for never positioning anything
again. Gate: if a recurring kind of slide genuinely needs a
different shape, add a named layout to `deck_layout.py`
chosen by section name, not coordinates in the deck file.

**PPTX only, no PDF or HTML output.** The renderer targets
one format. Gate: if slides need to go on the web, add a
second renderer reading the same deck file. The contract
already supports it, which is the reason for having a
contract.

**Everything is files in one directory, no database.** At
52 decks a year with about 12 to 40 pictures each, five
years is a few thousand files and a few GB. A filesystem
finds that boring. Gate: if a single week's `images/`
exceeds about 500 files, or a build takes over 20 seconds,
revisit.

Dependency discipline: four libraries, all mature, all
already installed. `python-pptx` 1.0.2, `Pillow`,
`openpyxl`, and ImageMagick 7 via `brew`, plus Google
Chrome for screenshots. Pin them in a `requirements.txt`
with exact versions and commit it. The freshest package is
the least reviewed package, and nothing here needs a new
feature.

## The shape of the thing

The whole design is eight decisions.

1. **One Markdown file per week is the source of truth.**
   Everything else is generated from it and can be thrown
   away and rebuilt.
2. **The author writes content and never layout.** The
   deck file has no coordinates, no widths, no slide
   breaks, and no font sizes, and no way to add them.
3. **A section is not a slide.** The generator paginates a
   section into as many slides as its text needs, repeating
   the title, so nobody counts items.
4. **The font ladder guarantees fit.** One shared size per
   slide, stepped down until the text fits, so overflow is
   solved rather than reported.
5. **Every news item is a headline, bullets, and one
   picture.** That uniformity is what makes automatic
   layout possible at all.
6. **Parsed numbers beat pictures of numbers.** The
   benchmark page is parsed text from JSON. Pages with
   nothing parseable get screenshotted instead, and
   `shot:` makes that one line.
7. **Raw pictures are permanent, clean pictures are
   disposable.** Fetching is the expensive step, so it
   happens once, and `s2` needs no backup of its own.
8. **The author owns the file; the assistant is a guest.**
   Two markers record a human update and a human deletion,
   a sidecar file keeps a rejected topic from coming back,
   matched by topic rather than by exact wording, and
   a checker enforces both so the guarantee is verified
   rather than promised.

## Appendix A: Requirements

Functional:

| ID | Requirement |
|---|---|
| F1 | Produce a 16:9 landscape PPTX, 10 x 5.625 in |
| F2 | Content authored in one plain-text file per week |
| F3 | Generate the table of contents from news headings |
| F4 | Lay out every slide automatically, with no positioning by the author |
| F5 | Paginate a section across as many slides as its text needs |
| F6 | Each news item is a headline, bullets, and one illustrative picture |
| F7 | Download pictures, or screenshot pages that are not pictures |
| F8 | Normalize pictures to one slide-ready standard |
| F9 | Render the benchmark page from leaderboard data as text, fitting the slide |
| F10 | Render URLs as visible clickable link lines |
| F11 | Deleting a topic records it in a sidecar file, and it is never recreated |
| F12 | Deletion works either by hand or by marking items and running one script |
| F13 | Output must remain editable in PowerPoint |

Non-functional, with measurable targets:

| ID | Requirement | Measured |
|---|---|---|
| N1 | Render from an unchanged deck file under 20 s | 0.34 s |
| N2 | Rebuild all images under 20 s | 1.6 s |
| N7 | A finished deck stays under 2 MB | 338 KB |
| N3 | Every stage idempotent and re-runnable | yes |
| N4 | No text silently clipped | 0 overflowing slides |
| N5 | Python rules: files under 800 lines, functions under 35 lines, lines under 65 chars | 9 of 9 files pass |
| N6 | No network access needed to build from cached images | yes |

## Appendix B: Technology choices

| Layer | Choice | Why |
|---|---|---|
| Deck format | Markdown | Greppable, diffable, already the project's habit |
| PPTX writing | python-pptx 1.0.2 | The only maintained option; installed |
| Layout | our own, 280 lines | Pure arithmetic over measured constants, so it can be tested without opening a PPTX |
| Image processing | ImageMagick 7 CLI | Installed, scriptable, no Python image dependency for conversion |
| Image measuring | Pillow | Needed only to read width and height when fitting a box |
| Screenshots | headless Google Chrome | Already on the machine; no Playwright or Selenium to install and pin |
| Benchmark data | `leaderboard.py` + openpyxl | Already written and working |
| News gathering | `/ai-news-digest` skill | Already written and working |
| Language | Python 3.13 | The project's language; time goes to ImageMagick, Chrome, and file I/O, so a faster language would speed up the glue, and the glue was never the bottleneck |

## Appendix C: Risks

| Risk | Mitigation |
|---|---|
| Automatic layout looks worse than hand-made | Output stays editable in PowerPoint; a fix in `deck_layout.py` improves every future deck instead of one slide |
| Height estimation drifts from what PowerPoint actually renders | Constants measured from real decks, asserted by layout tests, and the font ladder leaves headroom |
| A screenshot captures a cookie banner or a half-loaded page | `images_raw/` keeps it, so re-shooting is one `--force` away, and the raw file can be replaced by hand |
| Picture URL rots before the deck is built | `images_raw/` is permanent, so anything fetched once survives the source going away |
| `leaderboard.py` breaks when the source page changes | It exits non-zero when no models parse; a missing JSON leaves the page empty with a warning rather than failing the build |
| Deck grammar grows into a language | Seven rules, and any eighth needs a written reason here |
| A human edit is silently reverted by the next pass | The `locked` marker and the deleted sidecar, enforced by `check_preserved.py` and warned about by the build, rather than by good intentions |
| A thrown-out topic returns under a reworded headline | Topic matching folds case and space, accepts one headline containing the other, and compares significant words |

## Appendix D: Status of the code

| File | State |
|---|---|
| `deck_parser.py` | new, the grammar and topic matching |
| `deck_layout.py` | new, all geometry and pagination |
| `s3_make_pptx.py` | new, the only module touching python-pptx |
| `test_deck.py` | new, parser and layout tests |
| `s1_fetch_images.py` | new, replaces `s1_download_images.py`, reads the deck file instead of a hardcoded list, adds `shot:` |
| `s2_clean_images.py` | rewritten: raw to clean, aspect preserved, sized to the slide box at a chosen ppi, no backup machinery, orphan report |
| `leaderboard.py` | extended with the JSON contract; `save_to_excel` split to satisfy the 35-line rule |
| `s1_download_images.py` | superseded, safe to delete |
| `s3_make_pdf.py` | superseded, safe to delete |

The `slides-update` skill drives all of it:

| File | Purpose |
|---|---|
| `.claude/skills/slides-update/SKILL.md` | the weekly procedure and the ownership rules |
| `tools/deck_init.py` | resolve the date, create the deck from the skeleton |
| `tools/check_preserved.py` | enforce `locked` and the deleted sidecar |
| `reference/deleted_skeleton.md` | the sidecar's starting text |
| `reference/deck_skeleton.md` | the starting file, with the stable sections already locked |
