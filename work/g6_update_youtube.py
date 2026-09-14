#!/usr/bin/env python3
"""
Step 6: refresh the YouTube channel slide.

Reads the subscriber and video counts from the channel page
and replaces only the line of the promo box that mentions
subscribers, so the rest of its wording stays exactly as you
left it. Then it retakes the channel screenshot.

If the counts cannot be read from the page, a skill can pass
them in:

    {"subscribers": "7.51K", "videos": "337"}

Usage:
    python3 g6_update_youtube.py
    python3 g6_update_youtube.py --json counts.json

Created: 2026-09-14
Last updated: 2026-09-14
"""

import re

from gslides import write as W
from sources import fetch_images as F
from gslides import ids as T
from gslides.client import log, settings
from gslides.step import (
    load_json, page_or_stop, report_frozen, start,
    with_pictures,
)

LABEL = "youtube"
MARKER = "subscribers"
BOX = T.shape_id(T.TOPIC_YOUTUBE, T.BOX)
PICTURE = T.shape_id(T.TOPIC_YOUTUBE, T.PICTURE)

RE_SUBSCRIBERS = re.compile(r"([\d.,]+[KM]?) subscribers")
RE_VIDEOS = re.compile(r"([\d,]+) videos")


# --------------------------------------------------------------
def scrape_counts(url):
    """Subscribers and videos from the channel page."""
    data = F.fetch_bytes(url)
    if not data:
        return None
    text = data.decode("utf-8", "ignore")
    subs = RE_SUBSCRIBERS.search(text)
    videos = RE_VIDEOS.search(text)
    if not (subs and videos):
        return None
    return {"subscribers": subs.group(1),
            "videos": videos.group(1)}


# --------------------------------------------------------------
def counts_line(counts):
    """The promo line, in the wording the decks use."""
    return (f"{counts['subscribers']} subscribers, "
            f"{counts['videos']} videos")


# --------------------------------------------------------------
def plan_for(line):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URL."""

        # ----------------------------------
        def plan(deck):
            """Swap the counts line and the screenshot."""
            if not page_or_stop(deck, T.PAGE_YOUTUBE, LABEL):
                return []
            report_frozen(deck, [BOX], LABEL)
            reqs = []
            if line:
                reqs += W.replace_line(deck, BOX, MARKER, line)
            reqs += W.refresh_picture(deck, PICTURE,
                                      urls.get(PICTURE))
            return [(LABEL, reqs)]

        return plan

    return make


# --------------------------------------------------------------
def main():
    """Refresh the YouTube slide of one deck."""
    step = start("Update the YouTube slide")
    url = settings()["youtube_url"]
    counts = (load_json(step.args.json) if step.args.json
              else scrape_counts(url))
    line = counts_line(counts) if counts else None
    if line:
        log(f"{LABEL}: {line}")
    else:
        log(f"{LABEL}: counts not found; pass --json")
    sent = with_pictures(step, {PICTURE: {"shot": url}},
                         plan_for(line), LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
