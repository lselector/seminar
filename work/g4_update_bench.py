#!/usr/bin/env python3
"""
Step 4: refresh the Benchmarks slide of a live deck.

Runs sources/leaderboard.py for fresh LM Arena numbers,
compares the vote cutoff with the one already shown on
s-bench, and if it
is newer rebuilds the page's script shapes: legend, cutoff
note and the two columns.

Every box you filled stays exactly as it was. Shapes you
added yourself are never touched. If the page is gone or
parked past the separator, nothing happens.

Usage:
    python3 g4_update_bench.py                 next Friday
    python3 g4_update_bench.py 2026-09-18
    python3 g4_update_bench.py --force         rebuild anyway
    python3 g4_update_bench.py --no-fetch      use cached JSON

Created: 2026-09-14
Last updated: 2026-09-14
"""

import json
import os
import subprocess
import sys

from layout import bench_page as B
from gslides import render as R
from gslides import write as W
from gslides import ids as T
from gslides.client import DATA_DIR, ROOT, log
from gslides.step import page_or_stop, report_frozen, start

LEADERBOARD = os.path.join(DATA_DIR, "leaderboard.json")
DATE_BOX = T.shape_id(T.PAGE_BENCH, "date")
LABEL = "benchmarks"


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
def plan_for(data, force):
    """Build the planning function for run_plan."""

    # --------------------------------------
    def plan(deck):
        """Rebuild the page unless it is already current."""
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
        return [(LABEL, W.replace_owned(deck, sid, wanted))]

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
    sent = W.run_plan(step.slides, step.deck_id,
                      plan_for(data, step.args.force), LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
