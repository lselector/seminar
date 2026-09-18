#!/usr/bin/env python3
"""
Step 4: refresh the Benchmarks slide of a live deck.

Runs sources/leaderboard.py for fresh LM Arena numbers and
the date the site says its votes run through (the vote
cutoff), then updates the Benchmarks slide in one of two
ways.

The script's own slide (s-bench): if the cutoff is newer
than the one shown, the legend, cutoff note and the two
columns are rebuilt. Boxes you filled stay as they were.

A slide you built by hand, with two tables whose header row
reads Code | Model | Score: each table is filled in place
from its board, matched by the caption above it ("English",
"Coding"). A changed row gets the new model name and link,
the vendor colour in its Code cell, and the score. The box
holding a date ("Data for Sept 02") gets the cutoff date.
Nothing else on that slide is touched: fonts, row heights,
the legend and your notes all stay.

Usage:
    python3 g4_update_bench.py                 this week's Friday
    python3 g4_update_bench.py 2026-09-18
    python3 g4_update_bench.py --force         rebuild anyway
    python3 g4_update_bench.py --no-fetch      use cached JSON

Created: 2026-09-14
Last updated: 2026-09-18
"""

import datetime as dt
import json
import os
import subprocess
import sys

from layout import bench_page as B
from gslides import api as S
from gslides import bench as K
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import ids as T
from gslides.client import DATA_DIR, ROOT, log
from gslides.step import page_or_stop, report_frozen, start

LEADERBOARD = os.path.join(DATA_DIR, "leaderboard.json")
DATE_BOX = T.shape_id(T.PAGE_BENCH, "date")
LABEL = "benchmarks"

# The hand-made tables use a lighter grey for other vendors.
TABLE_COLORS = {**B.VENDOR_COLORS, "other": (0xD9, 0xD9, 0xD9)}


# --------------------------------------------------------------
def fetch_leaderboard():
    """Run sources/leaderboard.py so the JSON is current."""
    log("Fetching LM Arena leaderboards")
    done = subprocess.run([sys.executable, "-m",
                           "sources.leaderboard"],
                          cwd=ROOT, capture_output=True,
                          text=True)
    if done.returncode != 0:
        log("WARNING: the leaderboard fetch failed; using "
            "the cached JSON")
        log("  " + (done.stderr or done.stdout)[-300:])


# --------------------------------------------------------------
def load_boards():
    """The leaderboard data, or None when there is none."""
    if not os.path.isfile(LEADERBOARD):
        return None
    with open(LEADERBOARD, encoding="utf-8") as handle:
        return json.load(handle)


# --------------------------------------------------------------
def cutoff_day(boards):
    """The latest vote cutoff the site gives, as a date."""
    days = [b["cutoff"] for b in boards if b.get("cutoff")]
    return dt.date.fromisoformat(max(days)) if days else None


# --------------------------------------------------------------
def pair_tables(page, boards):
    """Match each board to the table under its caption."""
    tables = K.board_tables(page)
    pairs = []
    for index, board in enumerate(boards[:len(tables)]):
        word = board.get("label", "").lower()
        captions = [s for s in page.text_shapes()
                    if word and s.first_line().lower()
                    .startswith(word)]
        table = tables[index]
        if captions:
            x = captions[0].rect.x
            table = min(tables, key=lambda t: abs(t.rect.x - x))
        pairs.append((table, board))
    return pairs


# --------------------------------------------------------------
def find_date_box(page):
    """The first text box on the page that holds a date."""
    for shape in page.text_shapes():
        if W.DATE_WORDS.search(shape.text):
            return shape
    return None


# --------------------------------------------------------------
def row_requests(table, row, entry):
    """Requests that make one table row show one entry."""
    old = (table.cells[row] + ["", "", ""])[:3]
    name, score = entry.get("name", ""), str(entry.get("score"))
    reqs = []
    if old[1] != name:
        colour = TABLE_COLORS.get(entry.get("vendor"),
                                  TABLE_COLORS["other"])
        reqs += S.cell_colour(table.id, row, 0, colour)
        reqs += S.swap_cell_text(table.id, row, 1, old[1], name)
        if entry.get("url"):
            reqs += S.cell_link(table.id, row, 1, entry["url"])
    return reqs + S.swap_cell_text(table.id, row, 2, old[2],
                                   score)


# --------------------------------------------------------------
def table_requests(table, board):
    """Requests that fill one table from one board."""
    entries = board.get("entries", [])
    reqs = []
    for row in range(1, len(table.cells)):
        if row - 1 < len(entries):
            reqs += row_requests(table, row, entries[row - 1])
    return reqs


# --------------------------------------------------------------
def hand_requests(deck, boards):
    """Fill the hand-made tables and date; [] if current."""
    page = K.table_page(deck)
    if not page:
        return []
    reqs = []
    for table, board in pair_tables(page, boards):
        reqs += table_requests(table, board)
    day = cutoff_day(boards)
    if day:
        reqs += W.replace_date(find_date_box(page), day)
    return reqs


# --------------------------------------------------------------
def allowed_ids(deck):
    """The hand-made tables and date box this step may change."""
    page = K.table_page(deck)
    if not page or deck.slide(T.PAGE_BENCH):
        return ()
    ids = [s.id for s in K.board_tables(page)]
    box = find_date_box(page)
    return tuple(ids + ([box.id] if box else []))


# --------------------------------------------------------------
def script_requests(deck, data, force):
    """Rebuild the script's page unless it is current."""
    if not page_or_stop(deck, T.PAGE_BENCH, LABEL):
        return []
    note = B.cutoff_note(data["boards"][:2])
    shown = deck.shape(DATE_BOX)
    if shown and shown.text.strip() == note and not force:
        log(f"{LABEL}: already current ({note}).")
        return []
    sid = T.PAGE_BENCH
    wanted = R.render_bench(sid, sid, data, B, fill=None)
    report_frozen(deck, W.created(wanted), LABEL)
    return W.replace_owned(deck, sid, wanted)


# --------------------------------------------------------------
def plan_for(data, force):
    """Build the planning function for run_plan."""

    # --------------------------------------
    def plan(deck):
        """The script's page, or else the hand-made one."""
        boards = data["boards"][:2]
        if deck.slide(T.PAGE_BENCH):
            return [(LABEL, script_requests(deck, data, force))]
        if not K.table_page(deck):
            log(f"{LABEL}: no Benchmarks slide in the talk")
            return []
        reqs = hand_requests(deck, boards)
        if not reqs:
            log(f"{LABEL}: already current "
                f"({cutoff_day(boards)}).")
        return [(LABEL, reqs)]

    return plan


# --------------------------------------------------------------
def main():
    """Refresh the benchmarks page of one deck."""
    no_fetch = "--no-fetch" in sys.argv
    argv = [a for a in sys.argv[1:] if a != "--no-fetch"]
    step = start("Update the benchmarks slide", argv)
    if not no_fetch:
        fetch_leaderboard()
    data = load_boards()
    if not data or not data.get("boards"):
        log("ERROR: no leaderboard data. Run "
            "python3 -m sources.leaderboard")
        sys.exit(1)
    allow = allowed_ids(G.read_deck(step.slides, step.deck_id))
    sent = W.run_plan(step.slides, step.deck_id,
                      plan_for(data, step.args.force), LABEL,
                      allow=allow)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
