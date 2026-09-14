#!/usr/bin/env python3
"""
Step 5: refresh the Artificial Analysis Intelligence Index slide.

Takes a fresh screenshot of the Intelligence Index page and
swaps it into the picture frame on s-aa-index. If a skill
has written new bullets, they replace the words in the text
box too:

    {"headline": "Artificial Analysis Intelligence Index",
     "bullets": ["Claude Fable 5 leads at 73", "..."]}

The page link is kept as the last bullet. Both the box and
its picture stay as they are once you fill the box.

Usage:
    python3 g5_update_aa_index.py                 picture only
    python3 g5_update_aa_index.py --json aa.json  words too

Created: 2026-09-14
Last updated: 2026-09-14
"""

from gslides import write as W
from gslides import ids as T
from gslides.client import log, settings
from gslides.step import (
    block_from, load_json, page_or_stop, report_frozen,
    start, with_pictures,
)

LABEL = "intelligence index"
HEADLINE = "Artificial Analysis Intelligence Index"
BOX = T.shape_id(T.TOPIC_AA, T.BOX)
PICTURE = T.shape_id(T.TOPIC_AA, T.PICTURE)


# --------------------------------------------------------------
def plan_for(block):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URL."""

        # ----------------------------------
        def plan(deck):
            """Refresh the words, the picture, or both."""
            if not page_or_stop(deck, T.PAGE_AA, LABEL):
                return []
            report_frozen(deck, [BOX], LABEL)
            url = urls.get(PICTURE)
            if block:
                reqs = W.refresh_topic(deck, T.TOPIC_AA,
                                       block, url)
            else:
                reqs = W.refresh_picture(deck, PICTURE, url)
            return [(LABEL, reqs)]

        return plan

    return make


# --------------------------------------------------------------
def main():
    """Refresh the Intelligence Index slide of one deck."""
    step = start("Update the Intelligence Index slide")
    url = settings()["aa_index_url"]
    block = None
    if step.args.json:
        block = block_from(load_json(step.args.json),
                           HEADLINE, url)
    sent = with_pictures(step, {PICTURE: {"shot": url}},
                         plan_for(block), LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
