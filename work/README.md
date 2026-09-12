# Building the weekly deck

This directory turns one Markdown file into the weekly
seminar PPTX.

The Markdown file is the real deck. The PPTX is built from
it and can be deleted and rebuilt at any time. So edits go
in the `.md` file, not in PowerPoint, with one exception
covered at the end.

For the design and the reasoning, see `ADD.md`.

## The three files for one week

    2026-09-18-AI-News.md              you edit this
    2026-09-18-AI-News-generated.pptx  built from it
    2026-09-18-AI-News-deleted.md      topics you threw out

The date is the Friday of the seminar. The `-generated`
part of the name matters: the scripts only ever write files
with that suffix, so the hand-made decks in `../2026/` can
never be overwritten by mistake.

## 1. The easy way: run the skill

In Claude Code:

    /slides-update

or with a date:

    /slides-update 2026-10-02

With no date it uses the coming Friday. On a Friday it uses
today, not next week. It creates the `.md` file if it does
not exist, adds the week's news, and builds the PPTX.

## 2. Doing it by hand

Four commands, run from this directory:

    python3 leaderboard.py
    python3 s1_fetch_images.py 2026-09-18-AI-News.md
    python3 s2_clean_images.py
    python3 s3_make_pptx.py 2026-09-18-AI-News.md

What each one does:

| Script | Job |
|---|---|
| `leaderboard.py` | Downloads the Arena top 25 into XLSX and JSON, with the date the votes were counted through. Fills the Benchmarks slide |
| `s1_fetch_images.py` | Downloads or screenshots every picture the `.md` file asks for, into `images_raw/` |
| `s2_clean_images.py` | Converts them into slide-ready JPEGs in `images/` |
| `s3_make_pptx.py` | Builds the PPTX |

All four are safe to run again. They skip work that is
already done, so nothing is lost by repeating them.

## 3. Editing the Markdown

The whole file is built from four things.

**A section heading** starts a group of news:

    ## slide: AI News

**A news item** inside it:

    ### Nvidia buys Hugging Face for $13 Billion
    ![](images/nvidia-hf.jpg)
    <!-- shot: https://huggingface.co/blog/nvidia -->
    - Hugging Face hosts roughly 3 Mln models
    - **Nvidia sells the compute those models run on**
    - Cheap open models are closing the gap
    - https://huggingface.co/blog/nvidia

Each `-` line becomes a bulleted line with a red dot, so
write one short factual statement per bullet rather than a
paragraph. Five to seven of them per story works well.

Inside a bullet, `**bold**` makes bold text and
`==highlight==` paints the pale yellow used in the older
decks. A bullet that is nothing but a URL becomes a small
blue clickable link, sized to match the body text.

**The epigraph** is one short line under the deck title,
drawn in red at the top right of the contents page:

    # AI News - Sept 18, 2026

    > Agents got their own computers. Oversight is still catching up.

Keep it under about 70 characters. The build warns if it
wraps past two lines and crowds the page. The skill writes
one for you from the week's stories, and will not overwrite
one you changed yourself.

**Two special headings** fill themselves in:

    ## slide: toc           the table of contents
    ## slide: benchmarks    the Arena top 25

The contents page is generated from every `###` headline in
the file, so it can never disagree with the slides. There is
nothing to type there and nothing to keep in step.

### Do not write any layout

There is no way to set a position, a width, a font size, or
a slide break, and this is on purpose. Put as many `###`
items under one `## slide:` heading as the week needs, and
the generator splits them across as many slides as the text
requires, repeating the title. Seven stories under one
heading is normal.

Text boxes come out with the pale yellow fill and thin red
border of the older decks, sized to hug their text with no
slack left over. Pictures get the same red border. Slide
titles are left plain. A slide with room to spare uses
larger text, up to 16 pt. All of it is automatic.

If a slide comes out looking wrong, the fix belongs in
`deck_layout.py`, where it improves every future deck too.

## 4. Protecting your own edits

The assistant may edit this file again next week. One
marker tells it to keep its hands off, on its own line.

**You rewrote something and want it kept exactly:**

    ### A story I reworded myself
    <!-- locked -->
    - my wording, not the assistant's

There is one more marker, `<!-- profile -->`, already on
the About the Speaker section. It gives that slide the
shape the older decks used: a large portrait on the left
with the details beside it, and the name set larger. It is
the only slide that needs it.

Put the marker under a `###` to cover that one item, or
directly under a `## slide:` heading to cover the whole
section. The About the Speaker, YouTube, and Thank You
sections come locked already.

**You threw a topic out and do not want it back:** cut it
out of the deck and paste it into the deleted file. See the
next section.

### Deleting a topic

Do not just cut a story out and throw it away. If you do,
the assistant reads the deck next week, does not find the
story, finds it again in the news, and puts it back.

The story has to end up in `2026-09-18-AI-News-deleted.md`,
which is the record of what you threw out. There are two
ways to get it there. Both end in the same place, so use
whichever suits the moment.

**Method 1: move it yourself.** Good for one item while you
are already editing.

1. Cut the whole `###` item out of
   `2026-09-18-AI-News.md`.
2. Paste it into `2026-09-18-AI-News-deleted.md`, keeping
   its `###` heading. That heading is the record.
3. Optionally add a line saying why. Anything that is not a
   `###` heading is treated as a note and ignored.

**Method 2: mark them and run the script.** Easier when you
are dropping several, because you never leave the deck file
or lose your place.

Put the marker on its own line under each item you want
gone:

    ### A story I do not want
    <!-- deleted -->
    - its bullets stay with it

Then sweep them all out in one go:

    python3 move_deleted.py --dry-run
    python3 move_deleted.py

The dry run lists what would move and changes nothing. The
real run cuts every marked item out of the deck, appends it
to the deleted file with the date it moved, and drops the
marker line, since in the deleted file the heading alone is
the record. Text moves verbatim: bullets, images, links, and
formatting all survive.

The marker also works on a `## slide:` heading, which moves
that whole section and every item in it.

Your previous deck is kept as `2026-09-18-AI-News.md.bak`,
and the script refuses to touch a deck that does not parse.

Rebuild afterwards with `python3 s3_make_pptx.py`.

Nothing in the deleted file goes on a slide, and the
assistant will not write any of those topics back into the
deck. To bring something back, cut it out of the deleted
file and paste it into the deck again.

Matching is on the topic, not the exact wording. A story
re-added under a shortened or reworded heading is still
caught, so the assistant cannot slip one past by rephrasing
it. Two genuinely different stories about the same company
are not confused.

Both the build and the checker tell you if a thrown-out
topic is back:

    WARNING: 'The music industry stops fighting'
      matches 'The music industry stops fighting and starts
      licensing' in 2026-09-18-AI-News-deleted.md

### What is and is not protected

An item you edited **without** adding `<!-- locked -->`
looks exactly like one the assistant wrote, so nothing can
tell them apart. Mark what you want to keep.

Anything you add is safe without a marker: the assistant is
not allowed to remove an item that already exists.

To check that a round of edits preserved your work:

    python3 ../.claude/skills/slides-update/tools/check_preserved.py \
        before.md 2026-09-18-AI-News.md

It exits 0 when every protected item survived, and
otherwise lists what changed. Committing the `.md` file to
git before a session is the wider safety net.

## 5. Adding a picture by hand

Most pictures come from the `.md` file itself. Two forms:

    <!-- src: https://example.com/photo.jpg -->   download it
    <!-- shot: https://example.com/article -->    screenshot the page

`shot:` is the one to reach for. It photographs the source
page with Chrome, which always gives a relevant picture and
saves hunting for one.

For your own picture, a photo or a crop you made yourself:

1. Put the file in `images_raw/`. Any format works: PNG,
   JPEG, HEIC, WebP.
2. Name it after what it shows, with no spaces, for example
   `images_raw/lev-photo.png`.
3. Run `python3 s2_clean_images.py`. It writes
   `images/lev-photo.jpg`.
4. Reference it in the `.md` file **with no comment under
   it**, which is what marks it as yours:

        ### Lev Selector, Ph.D.
        ![](images/lev-photo.jpg)
        - 40+ years of software engineering

A picture with no `src:` or `shot:` comment is never
downloaded or overwritten. It is only checked for existence.

Put originals in `images_raw/`, not in `images/`. Files that
sit in `images/` with no source get reported as orphans
every run, because `images/` is meant to be disposable.

### Replacing a bad screenshot

If a screenshot catches a cookie banner or a half-loaded
page, replace the file in `images_raw/` with a picture you
took yourself, keeping the same name, then run
`s2_clean_images.py` again. Or re-shoot it:

    python3 s1_fetch_images.py 2026-09-18-AI-News.md --force

## 6. Final touches just before the seminar

This is the case that has to work without thinking. Edit
the `.md` file, then run one command:

    python3 s3_make_pptx.py

With no file named it takes the newest `*-AI-News.md` in
this directory and prints which one it picked. It needs no
network and takes under a second.

Only run `s1` and `s2` again if you added or changed a
picture. A text-only fix needs the one command above.

Two things to know:

- **Close the deck in PowerPoint first.** PowerPoint keeps
  its own copy in memory and can write it back over the
  freshly built file when you save.
- If you made the change in PowerPoint instead of the `.md`
  file, it will be lost on the next build. Put it in the
  `.md` file as well.

## 7. Keeping the file small

You used to open the finished PPTX, pick a picture, and run
**Compress Pictures** with "email (96 ppi)" and "delete
cropped areas", which took a deck from 20 or 30 MB down to
about 2 MB. That step is gone. The deck never gets large in
the first place.

A PPTX is almost entirely its pictures, and
`s2_clean_images.py` already scales every picture to the
largest box the layout can give it, which is 3.60 by 5.07
inches. At the default 150 ppi that is 540 by 760 pixels.
Anything beyond that is bytes the slide cannot show. There
are no cropped areas to delete either, because pictures are
trimmed before they go in, not cropped afterwards.

A typical deck of 11 pictures comes out around **340 KB**.

If you want it smaller or sharper, change the resolution
and rebuild:

    python3 s2_clean_images.py --ppi 96     # smallest
    python3 s2_clean_images.py              # 150, default
    python3 s2_clean_images.py --ppi 220    # sharpest
    python3 s3_make_pptx.py

Measured on the Sept 18 deck:

| Setting | Deck size |
|---|---|
| `--ppi 96`, PowerPoint's "email" | 190 KB |
| `--ppi 150`, the default | 338 KB |
| `--ppi 220`, PowerPoint's "print" | 591 KB |

Changing `--ppi` or `--quality` rebuilds every picture
automatically, because the setting is recorded inside each
file. You do not need `--force`.

The pictures in `images_raw/` stay full size. Only the
copies in `images/` are shrunk, so raising the resolution
later costs one command and no re-downloading.

## 8. Reading the output

A build ends with a summary worth a glance:

    Sections: 11, news items: 13
    Slides written: 11
    Topics in the deleted file: 1
    Missing images: 1
    Overflowing slides: 0
    Wrote ...-generated.pptx (338 KB, 86% pictures)

- **Missing images** names each one. That item still renders,
  using the full slide width for its text.
- **Overflowing slides** means one item has too much text to
  fit even at 8 pt. Split it into two `###` items.
- **Topics in the deleted file** counts what you threw out.
  If one of them is back in the deck, a loud `WARNING` names
  it just above this summary.
- **The size line** says how much of the deck is pictures.
  If it ever goes past 8 MB the build tells you how to
  shrink it.

## 9. Finishing the week

Copy the deck into the year directory and commit:

    cp 2026-09-18-AI-News-generated.pptx ../2026/
    cd .. && git add . && git commit -m "slides for 2026-09-18"

Nothing here writes to `../2026/` on its own.

## When something goes wrong

| Message | What to do |
|---|---|
| `ERROR in ...: line 47` | A grammar mistake on that line. Nothing was written and the old PPTX is untouched. Fix the line |
| Every image reported missing | Run `s1_fetch_images.py` then `s2_clean_images.py`. If you ran the build from another directory, that is fine now: paths follow the `.md` file |
| Benchmarks slide is empty | Run `python3 leaderboard.py` |
| `was not produced by this script` | The output name belongs to a hand-made deck. Use the `-generated.pptx` name, which is the default |
| `Orphan (no raw source)` | A file in `images/` with nothing behind it in `images_raw/`. Harmless. Delete it when you are sure |
| ImageMagick not found | `brew install imagemagick` |
| Screenshots all fail | Google Chrome is missing from `/Applications` |

## The files here

| File | What it is |
|---|---|
| `ADD.md` | The design document |
| `YYYY-MM-DD-AI-News.md` | One week's deck, the source of truth |
| `YYYY-MM-DD-AI-News-generated.pptx` | Built output, disposable |
| `YYYY-MM-DD-AI-News-deleted.md` | Topics you threw out, never recreated |
| `deck_parser.py` | Reads the Markdown grammar |
| `deck_layout.py` | Decides every slide, size, and rectangle |
| `s1_fetch_images.py` | Downloads and screenshots pictures |
| `s2_clean_images.py` | Normalizes them |
| `s3_make_pptx.py` | Builds the PPTX |
| `pptx_text.py` | The palette, fonts, and text boxes |
| `bench_page.py` | The benchmarks slide |
| `move_deleted.py` | Sweeps items marked `<!-- deleted -->` into the deleted file |
| `leaderboard.py` | Arena benchmarks into XLSX and JSON |
| `test_deck.py` | Grammar tests. `python3 test_deck.py` |
| `test_layout.py` | Layout tests. `python3 test_layout.py` |
| `test_runner.py` | Runs the asserts in a test file |
| `text_metrics.py` | Measures how wide text will be |
| `images_raw/` | Originals as fetched. Permanent |
| `images/` | Slide-ready JPEGs. Disposable, rebuilt by `s2` |
| `resources_raw/` | `leaderboard.json` and other hand-kept data |

`s1_download_images.py` and `s3_make_pdf.py` are left over
from an earlier version and can be deleted.
