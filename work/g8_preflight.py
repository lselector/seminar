#!/usr/bin/env python3
"""
Check a live deck before presenting it. Changes nothing.

Lists what still needs a human look:

    unreviewed    script boxes in the talk you have not
                  filled yet, so a script may still change
                  them and you have not accepted them
    placeholders  hints like [news goes here] still showing
    contents      a TOC that no longer matches the slides
    overflow      boxes whose text is estimated to spill
                  past their border
    pictures      news topics that have no picture
    benchmarks    a vote cutoff more than a week before
                  the seminar

Pages that are meant to have no border (About the Speaker,
Thank You, the separator) and slide titles are not asked
to be reviewed.

Exit status is 0 when there is nothing to report, else 1.

Usage:
    python3 g8_preflight.py
    python3 g8_preflight.py 2026-09-18

Created: 2026-09-14
Last updated: 2026-09-25
"""

import datetime as dt
import re
import sys

from layout import deck_layout as L
import g3_update_toc as TOC
from gslides import deck as G
from gslides import render as R
from gslides import ids as T
from gslides.step import start

NO_REVIEW = {T.PAGE_ABOUT, T.PAGE_THANKS, T.PAGE_SEPARATOR}
# Fixed decoration rather than content to accept.
NOT_CONTENT = {T.shape_id(T.PAGE_TOC, "epi"),
               T.shape_id(T.PAGE_BENCH, "leg"),
               T.shape_id(T.PAGE_BENCH, "date")}
NO_PICTURE = {T.TOPIC_THANKS, T.TOPIC_NEWS_1, T.TOPIC_NEWS_2}
CUTOFF_BOX = T.shape_id(T.PAGE_BENCH, "date")
STALE_DAYS = 7
OVERFLOW_SLACK = 0.10

# Measured on 2026-09-14: 12 pt Calibri lines are 14.24 pt
# apart in Slides, less than the 1.24 deck_layout plans for.
SLIDES_LINE_RATIO = 1.19
WIDTH_TOLERANCE = 1.03

RE_CUTOFF = re.compile(r"through (\w+) (\d+), (\d{4})")


# --------------------------------------------------------------
def unreviewed(deck):
    """Script boxes in the talk that have no fill yet."""
    found = []
    for slide in deck.main():
        if slide.id in NO_REVIEW:
            continue
        for shape in slide.text_shapes():
            if not T.owned(shape.id) or G.is_title(shape) \
                    or shape.id in NOT_CONTENT:
                continue
            if deck.editable(shape) and \
                    not G.is_placeholder(shape.text):
                found.append(f"{slide.index + 1:2d} "
                             f"{shape.first_line()[:60]}")
    return found


# --------------------------------------------------------------
def placeholders(deck):
    """Hints that were never replaced."""
    return [f"{s.index + 1:2d} {shape.text.strip()}"
            for s in deck.main() for shape in s.text_shapes()
            if T.owned(shape.id) and shape.text.strip()
            and G.is_placeholder(shape.text)]


# --------------------------------------------------------------
def stale_contents(deck):
    """A note when the TOC does not list what is there."""
    wanted = TOC.collect(deck)
    shown = TOC.shown(deck)
    if shown == wanted:
        return []
    return [f"lists {len(shown)}, slides have {len(wanted)}: "
            f"run g3_update_toc.py"]


# --------------------------------------------------------------
def estimated_height(shape):
    """Roughly how tall a box's text renders."""
    size = shape.size_pt or L.BASE_SIZE
    width = (shape.rect.w - 2 * R.PAD_X) * WIDTH_TOLERANCE
    lines = sum(max(1, L.wrapped_lines(p, width, size))
                for p in shape.text.split("\n"))
    return (lines * L.points_to_inches(size * SLIDES_LINE_RATIO)
            + 2 * R.PAD_Y)


# --------------------------------------------------------------
def overflowing(deck):
    """Boxes whose text probably spills past the border."""
    found = []
    for slide in deck.main():
        for shape in slide.text_shapes():
            if not shape.text.strip() or \
                    shape.id.endswith(f"-{T.TITLE}"):
                continue
            over = estimated_height(shape) - shape.rect.h
            if over > OVERFLOW_SLACK:
                found.append(f"{slide.index + 1:2d} "
                             f"{shape.first_line()[:50]} "
                             f"(about {over:.1f} in too tall)")
    return found


# --------------------------------------------------------------
def missing_pictures(deck):
    """News topics in the talk without a picture."""
    found = []
    for slide in deck.main():
        for shape in slide.text_shapes():
            key = T.topic_of(shape.id)
            if not key or key in NO_PICTURE:
                continue
            if not deck.shape(T.shape_id(key, T.PICTURE)):
                found.append(f"{slide.index + 1:2d} "
                             f"{shape.first_line()[:60]}")
    return found


# --------------------------------------------------------------
def old_benchmarks(deck, seminar):
    """A note when the vote cutoff is over a week old."""
    shape = deck.shape(CUTOFF_BOX)
    match = RE_CUTOFF.search(shape.text) if shape else None
    if not match:
        return []
    month = G.MONTHS.index(match.group(1)) + 1 \
        if match.group(1) in G.MONTHS else None
    if not month:
        return []
    cutoff = dt.date(int(match.group(3)), month,
                     int(match.group(2)))
    age = (seminar - cutoff).days
    if age <= STALE_DAYS:
        return []
    return [f"votes counted through {cutoff}, {age} days "
            f"before the seminar: run g4_update_bench.py"]


# --------------------------------------------------------------
def print_section(title, lines, hint):
    """Print one group of findings, if there are any."""
    if not lines:
        return 0
    tail = f" - {hint}" if hint else ""
    print(f"\n{title} ({len(lines)}){tail}")
    for line in lines:
        print(f"   {line}")
    return len(lines)


# --------------------------------------------------------------
def main():
    """Report everything that still needs attention."""
    step = start("Check a deck before the seminar")
    deck = G.read_deck(step.slides, step.deck_id)
    seminar = G.seminar_date(step.args.date)
    total = sum([
        print_section("Not reviewed yet", unreviewed(deck),
                      "fill each box once you accept it"),
        print_section("Placeholders", placeholders(deck),
                      "replace or delete them"),
        print_section("Contents", stale_contents(deck), ""),
        print_section("Probably overflowing",
                      overflowing(deck), "shorten or enlarge"),
        print_section("No picture", missing_pictures(deck),
                      "add one by hand or re-run the step"),
        print_section("Benchmarks", old_benchmarks(deck,
                                                   seminar), ""),
    ])
    print("\nReady to present." if not total else
          f"\n{total} thing(s) to look at.")
    sys.exit(0 if not total else 1)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
