# Architecture Design Document: Weekly AI News Deck

Status: Built and in use · Date: 2026-09-14

## Executive summary

Every Friday a seminar deck of about 13 slides has to exist:
contents, benchmarks, 6 to 10 news topics with pictures and
links, a few standing pages. Built by hand, most of the
evening goes on mechanics: typing headlines twice, dragging
boxes, cropping screenshots, re-checking numbers.

**The Google Slides deck is the single source of truth.**
Scripts write content into the live deck, and the author
edits the same deck by hand whenever they like. There is no
intermediate file to keep in sync. An earlier design kept a
Markdown file as the truth and generated PPTX from it; it
was built, then cancelled in favour of working in Slides
directly.

The hard problem is sharing one document between a person
and a program without either destroying the other's work.
It is solved with three rules the author can learn in a
minute: **a script only changes shapes it created, a
background fill freezes a box, and nothing after the
separator slide is ever written.** Every rule was measured
against the real Slides API, and in the browser, before code
depended on it.

Measured on the Sept 18 deck: 11-slide skeleton created in
29 s; 7 news topics with screenshots added in 68 s; each
refresh step 5 to 12 s; a second run of any step changes
nothing. 150 tests, all offline.

## Contents

- [The principles, at a glance](#the-principles-at-a-glance)
- [The problem](#the-problem)
- [Non-goals](#non-goals)
- [The architecture](#the-architecture)
- [Who owns a shape](#who-owns-a-shape)
- [The deck skeleton](#the-deck-skeleton)
- [Modules and dependencies](#modules-and-dependencies)
- [The AI layer](#the-ai-layer)
- [Lifecycle: what each edit does](#lifecycle-what-each-edit-does)
- [Automatic layout](#automatic-layout)
- [The pages that are not news](#the-pages-that-are-not-news)
- [Pictures](#pictures)
- [Writing safely to a live document](#writing-safely-to-a-live-document)
- [Access and credentials](#access-and-credentials)
- [Verification](#verification)
- [Failure handling](#failure-handling)
- [Known limits and scaling gates](#known-limits-and-scaling-gates)
- [The shape of the thing](#the-shape-of-the-thing)
- [Appendix A: Requirements](#appendix-a-requirements)
- [Appendix B: Technology choices](#appendix-b-technology-choices)
- [Appendix C: Risks](#appendix-c-risks)
- [Appendix D: Measurements behind the design](#appendix-d-measurements-behind-the-design)
- [Appendix E: Code from the cancelled Markdown design](#appendix-e-code-from-the-cancelled-markdown-design)

## The principles, at a glance

1. **Simplicity.** One document, a few short scripts, no
   server, no database, no sync file.
2. **Modularity.** Read model, safe writer, renderer and
   picture host are separate modules; each step is a thin
   script on top.
3. **AI does the judgement, scripts do the writing.** Skills
   find and phrase content; deterministic scripts place it.
4. **One system of record.** The live deck. Everything else
   (screenshots, JSON, the leaderboard cache) is disposable.
5. **The author owns the deck; scripts are guests.**
   Ownership is decided by rules Google enforces, not by
   good behaviour.
6. **Measure before you depend.** Every API behaviour the
   rules rely on was probed first.
7. **Idempotent and order-free.** Any step, any time, any
   number of times.
8. **Boring dependencies.** Official Google client, Pillow,
   ImageMagick, headless Chrome; nothing new to install.

## The problem

The archive in `2026/` holds 41 decks from this year. They
are good decks, and each costs an evening of mechanics:

- The same headline is typed on the news slide and again in
  the contents on slide 1. When a story is cut late, slide 1
  still lists it.
- Every box and picture is positioned by hand, and none of
  that work carries to next week.
- Screenshots arrive at any size. Some are 4000 px wide.
- Standing pages (benchmarks, YouTube counts, layoff
  numbers) go stale unless someone remembers them.

A generator alone does not fix this, because the author
still wants the last word: rewording a bullet, dropping a
story at 11 pm, dragging a slide. So the constraint that
shapes everything is **the author and the automation must
work on the same deck, at any time, in any order, and
neither may undo the other.**

Fixed constraints: one person, a Mac, a personal Google
account; slides 16:9 at 10 x 5.625 in; no access to the rest
of the author's Drive.

## Non-goals

- It does not choose the news on its own. A skill proposes
  topics; the author accepts them by filling the box.
- It does not reproduce the hand-tuned collages of older
  decks. Every news slide has one shape.
- It does not export PPTX, PDF or video, and does not publish.
- It does not see or touch any Drive file it did not create.
- It does not merge concurrent edits. If the author types
  while a script writes, the script re-reads and plans
  again; it never tries to combine the two.

## The architecture

```
  skills (Claude)                 scripts (Python)            Google
  ---------------                 ----------------            ------
  gslides-add-news   --news.json-->  g2_add_news  ---+
  gslides-update-*   --json------->  g3 .. g7     ---+--> Slides API --> live deck
                                     g1_new_deck  ---+                    ^
                                     g8_preflight (read only)             |
                                                                     author edits
                     screenshots --> sources/pictures --> gslides/host --> picture copy
```

Two kinds of worker, one document:

- **Skills** need judgement: read newsletters, pick stories,
  read numbers off a web page, write bullets. Their output
  is a small JSON file.
- **Scripts** are deterministic. They read the deck, decide
  what they may change, and send one batch per slide.

There is no local copy of the deck. Each run starts by
reading the live presentation, so a change the author made a
second ago is already part of the plan.

## Who owns a shape

The rules are the contract between author and automation.

**1. The object id decides ownership.** Scripts give every
slide and shape an id starting with `s-` (slides and page
furniture) or `t-` (topics). Anything else belongs to the
author. This holds because Slides enforces unique ids across
a presentation: a box the author copies or cut-and-pastes
gets a new id (`g3fb561739ea_0_0` in the browser), so it
stops being the script's the moment it lands.

An earlier idea stamped a topic key into alt text. Probing
showed duplication copies alt text verbatim, so three boxes
claimed one topic. Object ids cannot be copied, which is why
they won.

**2. A fill freezes a box.** Scripts draw text boxes with a
red border and no background. Any fill, any colour, means
the author accepted it, and no script changes that box or
its picture again. The check reads
`shapeBackgroundFill.propertyState`; an unfilled box reports
`NOT_RENDERED` but still carries a white colour, so comparing
colours would be wrong.

**3. The separator is a wall.** Nothing after the slide `Not
in the presentation` is ever written. Parked topics live
there.

A script may change a shape only when all three hold: id
starts with `s-` or `t-`, no fill, before the separator. A
picture `t-...-p` follows its text box `t-...-b`.

Ids are built from the topic's own words, so they survive
reordering and editing:

    t-meta-ships-muse-agent-with-its-a345-b
    t = ours, slug = first 32 characters of the headline,
    a345 = hash of the full headline, b = text box (p picture)

Fixed pages have fixed ids without a hash (`s-bench`,
`t-trueup`), and every generated id ends in one, so the two
families cannot collide. Slides requires 5 to 50 characters
of letters, digits, `_` and `-`; `gslides/ids.py` enforces it.

## The deck skeleton

`g1_new_deck.py` creates this for a date:

| # | Slide id | Content at creation | Kept current by |
|---|---|---|---|
| 1 | `s-toc` | title, contents placeholder, epigraph placeholder | g3 |
| 2 | `s-bench` | legend, cutoff note, two columns | g4 |
| 3 | `s-aa-index` | text box, picture | g5 |
| 4 | `s-news-1` | "AI News", one placeholder box | g2 |
| 5 | `s-youtube` | promo text, channel picture | g6 |
| 6 | `s-news-2` | "AI News", one placeholder box | g2 |
| 7 | `s-layoffs` | Layoffs.fyi and TrueUp topics | g7 |
| 8 | `s-about` | portrait and bio | fixed |
| 9 | `s-thanks` | links | fixed |
| 10 | `s-separator` | divider note; topic ledger in speaker notes | never |
| 11 | `s-parked` | empty, for rejected topics | never |

The wording of pages 3, 5, 7, 8 and 9 lives in
`config/skeleton.json`, the one place that text is kept.

The file is named `YYYY-MM-DD-AI-News` and tagged with
`seminarDate` in Drive `appProperties`. Steps find the deck by
that tag, so renaming or moving it breaks nothing. Step 1
refuses to create a second deck for a date.

## Modules and dependencies

Three packages, each with one job and a `README.md` stating
its modules, public API and dependencies:

| Package | Job | Depends on |
|---|---|---|
| `layout/` | content model, inline markup, topic matching, geometry, text measurement, benchmark page sizes; pure | nothing |
| `sources/` | outside content on this machine: leaderboards, downloads, screenshots, ImageMagick | `layout` |
| `gslides/` | the live deck: read model, ownership rules, rendering, safe writes, picture hosting | `layout`, `sources` |

The `g1` to `g8` commands sit above all three as thin
composition roots, one per step. Around them: `config/`
(settings and the standing slides' wording), `assets/`
(starter pictures), `data/` (generated leaderboard files),
`tests/` and `probes/`.

Inside `gslides/`:

| Module | Owns |
|---|---|
| `deck.py` | Read model: Deck, Slide, Shape; `owned`, `frozen`, `editable`, `main`, `parked`, `room_below`; date and name rules; deck lookup |
| `write.py` | Safe writing: revision check, `guard`, `refresh_text`, `refresh_picture`, `refresh_topic`, `replace_line`, `replace_owned`, `run_plan` |
| `render.py` | Content to requests: text styling, boxes, pictures, contents columns, benchmark columns; Slides padding compensation |
| `api.py` | Individual Slides request dictionaries |
| `ids.py` | Id scheme, fixed ids, ownership prefix |
| `host.py` | Short-lived public picture URLs on Drive |
| `client.py` | Sign-in, `config/gslides.json`, project paths |
| `step.py` | Arguments, deck lookup and picture hosting shared by steps |

Dependencies never point back up, and `sources` holding the
local picture work while `gslides/host.py` holds the Drive
upload is what keeps it that way. Only `client`, `read_deck`,
`write.apply` and `host` talk to Google; everything that
decides *what* to write is a pure function from a Deck to a
list of requests, which is what makes it testable offline.

Code rules, checked on every file: under 800 lines,
functions under 35 lines, lines under 65 characters, a
docstring on every module, function and class.

## The AI layer

| Agent | Knowledge base | Memory | Tools | Never decides alone |
|---|---|---|---|---|
| `gslides-add-news` | `ai-news-digest` over the author's newsletters; web pages it cites | the deck itself and the topic ledger (`g2_add_news.py --list`) | `g1`, `g2`, `g3` | whether a topic is in the talk: the author fills or parks it |
| `gslides-update-toc` | the current slides | none | `g3` | the epigraph once written |
| `gslides-update-benchmarks` | LM Arena via `sources/leaderboard.py` | none | `g4` | overriding a filled column |
| `gslides-update-aa-index` | the Artificial Analysis page | none | `g5` | overriding a filled box |
| `gslides-update-youtube` | the channel page | none | `g6` | wording beyond the counts line |
| `gslides-update-layoffs` | layoffs.fyi, TrueUp | none | `g7` | overriding a filled topic |

Agents never edit the deck directly. They reach it only
through the step scripts, so every write passes the same
ownership guard. The agent's memory of "what was already
added" is not in the agent: it is the ledger in the deck,
visible to the author and erasable by deleting lines from
the separator's speaker notes.

## Lifecycle: what each edit does

| The author... | Scripts then... |
|---|---|
| fills a box | never change it or its picture |
| types in an unfilled script box | may overwrite it on the next refresh; filling protects it |
| adds a box or slide | never touch it |
| copies or cut-pastes a script box | treat the pasted box as the author's |
| moves or resizes a script box | keep position and width; height follows new text, growing only into free space |
| reorders slides | follow the new order: contents list it; new news goes before layoffs wherever it is |
| parks a topic past the separator | never add it again, even reworded |
| deletes a topic outright | never add it again: the ledger remembers it |
| fills one contents column | leave the whole contents alone |
| writes the epigraph | keep it |
| deletes or parks a fixed slide | that step reports it and writes nothing |
| copies the whole deck | cannot see the copy (see Access) |

**How "never add again" works.** Before adding, `g2`
compares each candidate headline with the first line of
every text box in the deck, parked section included, and
with every headline in the ledger. Matching is on the topic,
not the wording (`same_topic`: equal after normalising, one
containing the other, or 70% of significant words shared).
Probed with reworded headlines for a parked topic and for a
deleted slide; both were skipped.

**Where new topics go.** First onto `s-news-1` and
`s-news-2` while each still holds only its placeholder and
nobody has put anything else on it; then as new slides just
before `s-layoffs`, at its position at that moment.

## Automatic layout

`layout/deck_layout.py` turns a list of topics into placed
rectangles; the author never positions anything.

1. **Estimate** each topic's height from measured text
   widths (`layout/text_metrics.py`: Arial metrics scaled by 0.9148
   to Calibri, fitted on 12 observed line breaks).
2. **Pack** topics onto a slide while they fit and fewer than
   three share it.
3. **One font size per slide**, trying 12, 11, 10, 9, 8 pt.
4. **Distribute** leftover height as equal gaps, capped at
   0.40 in; the rest goes to pictures.

Geometry (inches): slide 10 x 5.625, margin 0.09, content
band 0.46 to 5.53, text column 6.10 beside a 3.60 picture
column. Titles Calibri 20 bold; headlines red bold; bullets
Calibri 12 as a real bulleted list with a hanging indent;
code Consolas 9 blue; links 9 pt blue. A Slides bullet takes
the colour and size of the first character on its line, so
news bullets are black and contents bullets blue. Decks
written before 2026-09-16 have a typed red dot instead.

**Slides-specific corrections**, all measured:

- Slides pads text 0.10 in left and right and 0.05 in top
  and bottom, and the API cannot change it. Every box is
  grown by the difference and shifted out by half of it, so
  text lands where the layout planned.
- A benchmark column gets 0.06 in of slack, because a URL of
  slashes and dots measures a little short.
- When a script rewrites an existing box, its height is set
  to its new text within the free space below it, instead of
  shrinking the font inside the old height.

## The pages that are not news

**Benchmarks: parsed data, not a screenshot.**
`sources/leaderboard.py` writes the top 25 of the Arena English and
Coding boards to `data/leaderboard.json`, with each
model's score, vendor and the board's vote cutoff. The page
is two text columns, `marker score name`, vendor-coloured,
at 9 pt, centred, with a legend and a "Votes counted through"
note. Text beats a picture of numbers: crisp on a projector,
searchable, nothing to crop. It is plain paragraphs, not a
table, because an earlier PPTX table could not be made short
enough. `g4` rebuilds only when the cutoff changes.

**Intelligence Index: a screenshot**, because it is a live
chart. `g5` retakes it; bullets change only when a skill
supplies them.

**YouTube: one line.** `g6` reads subscriber and video counts
from the channel page and replaces only the line containing
`subscribers`, so the rest of the promo wording is the
author's.

**Layoffs: two independent topics.** `g7` refreshes the
Layoffs.fyi and TrueUp boxes separately, so freezing one
does not block the other.

## Pictures

Slides accepts a picture only as a URL it fetches itself.
`sources/pictures.py` does it in three stages, in a temp directory:

1. **Fetch**: download (`src`) or screenshot with headless
   Chrome at 1600 x 1000 (`shot`).
2. **Clean** with ImageMagick: flatten on white, sRGB, trim
   2% border, fit the picture box at 150 ppi, JPEG q85.
3. **Host**: upload to the deck's folder, share by link, give
   Slides the `uc?export=download` URL, then delete the upload.

Slides copies the picture into the presentation at insert
time. Probed: the picture stayed after the upload was
unshared and deleted. So nothing is pushed to GitHub and no
public file outlives the run.

Before screenshotting, `sources/pictures.challenged()` fetches the page
and looks for a bot wall. TrueUp answers with a Cloudflare
challenge (HTTP 403), so its old picture is kept instead of a
"security verification" screen.

## Writing safely to a live document

Four layers, outermost first:

1. **Planning helpers** return nothing for a shape that is
   not editable.
2. **`guard`** drops any request aimed at an existing shape
   or slide the rules protect, even if a planner let it
   through, and logs what it dropped.
3. **Revision check.** Every batch carries the revision id
   the plan was made from. If the author typed in between,
   Slides refuses with HTTP 400 ("does not match the latest
   revision"). `run_plan` re-reads the deck and plans once
   more, then stops and says so.
4. **Geometry.** Rewrites keep a box's position and width;
   pictures use `replaceImage`, which keeps the frame.

Scripts hold no state between runs. Everything they need is
re-read from the deck, so an interrupted run leaves either
whole batches or none, and the next run finishes the job.

## Access and credentials

OAuth for the author's own Google account, project
`jarvis-lev`, desktop client. Scopes: `drive.file`,
`presentations`, `documents`. `drive.file` means the app can
see only files it created: probed, it listed exactly its own
files and nothing else in the Drive. Consequences, all
measured:

- rename and move keep access (same file id)
- a copy the author makes is invisible to the scripts
- creating inside an existing folder works, given its id

`credentials/client_secret.json` and `credentials/token.json`
are gitignored. The folder id and source URLs are in
`config/gslides.json`, which holds nothing secret.

## Verification

**Code.** 150 tests, no network:

- `tests/test_gdeck.py`, `tests/test_gwrite.py`, `tests/test_gsteps.py` run on
  two saved real API responses in `fixtures/`: a fresh
  skeleton, and a lived-in deck where a human had moved
  layoffs, added a box, filled boxes, parked and deleted
  topics. Human edits are simulated by changing that JSON as
  the API reports them.
- `tests/test_ids.py` covers the id scheme; `tests/test_layout.py` and
  `tests/test_deck.py` cover layout and topic matching.
- A conformance check enforces the code rules on every file.

**Behaviour.** Each step was run twice against a scratch
deck, the second run confirmed a no-op, frozen shapes were
read back unchanged, and every page was exported to PDF and
looked at (`probes/render_deck.py`).

**Before presenting.** `g8_preflight.py` is read-only and
lists unaccepted script boxes, leftover placeholders, a
contents list that no longer matches the slides, boxes that
probably overflow, topics without a picture, and benchmarks
more than 7 days old. Exit 0 means "Ready to present."

## Failure handling

| Failure | Behaviour |
|---|---|
| no deck for the date | step exits, names the `g1` command |
| deck already exists | `g1` prints its link, creates nothing |
| author types during a write | re-read, re-plan once, then stop cleanly |
| fixed slide deleted or parked | step logs it and writes nothing |
| page unreachable or behind a bot wall | old picture kept, named in the log |
| `sources/leaderboard.py` fails | cached JSON used, warning printed |
| YouTube counts not found | skill passes them with `--json` |
| upload left behind by a crash | `Host.clean()` runs in `finally`; names start `_tmp-seminar-` |
| token missing | step names `g_auth.py` |

Nothing is auto-repaired in the deck itself. A script that
cannot tell whether a shape is its own leaves it alone.

## Known limits and scaling gates

- **One shape of news slide.** Text left, picture right, up
  to three topics. Gate: a recurring slide kind that needs
  another shape gets a named layout in `layout/deck_layout.py`, not
  per-slide positions.
- **Overflow is estimated, not measured.** The API cannot
  report rendered text height. Gate: if preflight misses
  real overflow more than once a month, render a thumbnail
  and measure.
- **Sequential, one batch per slide.** Gate: if a full run
  passes 3 minutes, host pictures in parallel.
- **Copies are invisible.** A consequence of `drive.file`,
  accepted for privacy. Gate: none; it is the point.
- **Author's unfilled edits can be overwritten** by a
  refresh step. Documented, and preflight lists every
  unaccepted box.

## The shape of the thing

The whole design is seven decisions.

1. **The live deck is the only source of truth.** No file to
   sync, nothing to regenerate, and the author edits the real
   thing.
2. **Ownership is structural.** A script owns what carries
   its id prefix. Google will not let a copy keep that id, so
   ownership cannot leak.
3. **A fill is consent.** One click tells every script to
   stop, with no markers, no file and no tool.
4. **The separator is a wall, and the ledger is a memory.**
   Parked and deleted topics never come back, however they
   are reworded.
5. **AI proposes, scripts place, the author decides.**
   Judgement lives in skills, writes live in deterministic
   code, acceptance lives in the deck.
6. **Every write is checked three times.** Planner, guard,
   revision id. A clash means re-read, never merge.
7. **Nothing was assumed.** Every API behaviour the rules
   depend on was probed before code relied on it.

## Appendix A: Requirements

| ID | Requirement |
|---|---|
| F1 | Create a dated 16:9 deck with the standing pages, in a chosen Drive folder |
| F2 | Add news topics (headline, bullets, link, picture) with automatic layout |
| F3 | Generate the contents list from the slides in their current order |
| F4 | Refresh benchmarks, Intelligence Index, YouTube counts and layoffs pages |
| F5 | Never change a box the author filled, or its picture |
| F6 | Never change anything the author created or pasted |
| F7 | Never write past the separator; never re-add a parked or deleted topic |
| F8 | Every step runnable at any time, in any order, repeatedly |
| F9 | A read-only pre-seminar check |
| S1 | No access to any Drive file the app did not create |
| S2 | Credentials never committed; pictures public only while being inserted |
| N1 | A second run of any step changes nothing (measured: yes) |
| N2 | Skeleton under 60 s (29 s); a refresh step under 30 s (5 to 12 s) |
| N3 | No text clipped at 12 pt on news slides (0 on the Sept 18 deck) |
| N4 | Code rules: files < 800 lines, functions < 35, lines < 65, docstrings (all pass) |

## Appendix B: Technology choices

| Layer | Choice | Why |
|---|---|---|
| Document | Google Slides | The author works there; editing, sharing and presenting are free |
| API client | `google-api-python-client`, `google-auth-oauthlib` | Official; installed |
| Scope | `drive.file` | Least access that works; probed |
| Layout | own `layout/deck_layout.py` | Arithmetic over measured constants, testable without Google |
| Pictures | headless Chrome, ImageMagick 7, Pillow | Already on the machine; no Playwright to pin |
| Picture hosting | temporary Drive upload | No GitHub push; Slides keeps its own copy |
| Benchmark data | `sources/leaderboard.py` | Already written; owns vendor rules |
| News gathering | `ai-news-digest` skill | Already written |
| Language | Python 3.13 | Time goes to network, Chrome and ImageMagick; faster glue would not help |

## Appendix C: Risks

| Risk | Mitigation |
|---|---|
| A script overwrites the author's work | id ownership + fill freeze + separator, enforced in planner and `guard`; tested on a lived-in fixture |
| Author and script write at once | revision id on every batch; re-read and re-plan |
| A rejected topic returns reworded | topic matching against deck, parked slides and ledger |
| Google changes an API behaviour the rules rely on | probes in `probes/` re-run in a minute |
| Screenshot shows a cookie or bot page | bot-wall check; picture replaceable by hand |
| Token expires or is revoked | `g_auth.py` re-runs sign-in |
| Layout estimate drifts from real rendering | measured padding and line pitch; preflight overflow check |

## Appendix D: Measurements behind the design

| Probe | Result |
|---|---|
| Page size | 9144000 x 5143500 EMU = 10 x 5.625 in, same as the layout |
| Object id | at least 5 characters, unique across the presentation |
| Duplicate a box (API and browser) | new id; alt text copied verbatim |
| Fill by hand in the browser | reads as filled |
| Unfilled box | `propertyState: NOT_RENDERED`, colour still white |
| Default text padding | 0.10 in sides, 0.05 in top and bottom, not settable |
| Line pitch, 12 pt Calibri | 14.24 pt |
| Stored box size | fixed 3,000,000 EMU square plus scale transform |
| Stale `requiredRevisionId` | HTTP 400, "does not match the latest revision" |
| Delete and recreate one id in a batch | works |
| `appProperties` search | finds the deck by date |
| Speaker notes | writable and readable (topic ledger) |
| Drive-hosted picture | inserted; survives unsharing and deleting the upload |
| `drive.file` visibility | only app-created files; rename and move keep access |
| TrueUp | Cloudflare challenge to automated clients |

The probe scripts are in `probes/`.

## Appendix E: Code from the cancelled Markdown design

The Markdown-to-PPTX workflow was removed on 2026-09-14:
its deck files, `s3_make_pptx.py`, `s4_push_slides.py`,
`pptx_text.py`, `move_deleted.py` and the `slides-update`
skill. git history has them.

Four modules from it are still used, for these parts only:

| Module | Used for |
|---|---|
| `layout/deck_parser.py` | `Block`, `Section`, inline markup, `same_topic` |
| `layout/deck_layout.py`, `layout/text_metrics.py` | all layout and measurement |
| `sources/fetch_images.py` | download and Chrome screenshot helpers |
| `sources/clean_images.py` | the ImageMagick conversion |

Each still carries a Markdown-reading command line of its
own (and `layout/deck_parser.py` the Markdown grammar and deleted
sidecar functions, with their tests). None of that is called
by the Slides workflow.
