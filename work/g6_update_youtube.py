#!/usr/bin/env python3
"""
Step 6: refresh the YouTube channel slide.

Reads the subscriber and video counts from the channel page
and replaces only the line that mentions subscribers, so the
rest of the wording stays exactly as you left it. Then it
retakes the channel screenshot.

The counts line is the one thing on this slide the step keeps
current by request, so it is rewritten even in a box you
filled, and even in a promo box you made yourself. Everything
else on the slide, your pictures included, is left alone.

If the promo box is gone altogether, the step draws it again
from config/skeleton.json with the current counts in it.

If the counts cannot be read from the page, a skill can pass
them in:

    {"subscribers": "7.51K", "videos": "337"}

Usage:
    python3 g6_update_youtube.py
    python3 g6_update_youtube.py --json counts.json

Created: 2026-09-14
Last updated: 2026-09-17
"""

import re
from dataclasses import dataclass

from layout import deck_layout as L
from layout import skeleton as K
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from sources import fetch_images as F
from gslides import ids as T
from gslides.client import ROOT, log, settings
from gslides.step import load_json, start, with_pictures

LABEL = "youtube"
MARKER = "subscribers"
PAGE_KEY = "youtube"
BOX = T.shape_id(T.TOPIC_YOUTUBE, T.BOX)
PICTURE = T.shape_id(T.TOPIC_YOUTUBE, T.PICTURE)

RE_SUBSCRIBERS = re.compile(r"([\d.,]+[KM]?) subscribers")
RE_VIDEOS = re.compile(r"([\d,]+) videos")


@dataclass
class Target:
    """What this step changes on the slide it found."""

    page: object
    box: object = None
    allow: tuple = ()


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
def find_page(deck):
    """The channel slide in the talk, or None."""
    page = deck.slide(T.PAGE_YOUTUBE)
    if page and page.index < deck.separator():
        return page
    want = K.page(PAGE_KEY)["title"].lower()
    for slide in deck.main():
        for shape in slide.text_shapes():
            if G.is_title(shape) and \
                    shape.first_line().lower() == want:
                return slide
    return None


# --------------------------------------------------------------
def find_box(page):
    """The box holding the counts line: the script's or yours."""
    found = [s for s in page.text_shapes()
             if MARKER in s.text and not G.is_title(s)]
    for shape in found:
        if shape.id == BOX:
            return shape
    return found[0] if found else None


# --------------------------------------------------------------
def locate(deck):
    """The slide, its counts box, and what to allow."""
    page = find_page(deck)
    if not page:
        return None
    box = find_box(page)
    return Target(page=page, box=box,
                  allow=(box.id,) if box else ())


# --------------------------------------------------------------
def promo_section(line):
    """The skeleton promo section, holding the new counts."""
    section = K.section(K.page(PAGE_KEY), ROOT)
    for block in section.blocks:
        block.bullets = [line if MARKER in b else b
                         for b in block.bullets]
    return section


# --------------------------------------------------------------
def draw_box(slide_id, line):
    """Draw the promo box again, as step 1 first drew it."""
    pages = L.paginate([promo_section(line)])
    if not pages or not pages[0].items:
        return []
    page, item = pages[0], pages[0].items[0]
    reqs = R.make_box(BOX, slide_id, item.text,
                      not page.plain, None)
    return reqs + R.write_body(BOX, R.block_body(item))


# --------------------------------------------------------------
def counts_reqs(deck, target, line):
    """Set the counts line, drawing the box if it is gone."""
    if target.box:
        reqs = W.replace_line(deck, target.box.id, MARKER,
                              line, target.allow)
        if reqs and not deck.editable(target.box):
            log(f"{LABEL}: the box is filled; only the counts "
                f"line is updated")
        return reqs
    if deck.shape(BOX):
        log(f"{LABEL}: the box has no '{MARKER}' line; "
            f"nothing written")
        return []
    log(f"{LABEL}: the promo box is gone; drawing it again")
    return draw_box(target.page.id, line)


# --------------------------------------------------------------
def plan_for(line, found):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URL."""

        # ----------------------------------
        def plan(deck):
            """Swap the counts line and the screenshot."""
            target = locate(deck)
            if not target or target.page.id != found.page.id:
                log(f"{LABEL}: the slide changed. Nothing done.")
                return []
            reqs = counts_reqs(deck, target, line) if line \
                else []
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
    found = locate(G.read_deck(step.slides, step.deck_id))
    if not found:
        log(f"{LABEL}: no channel slide in the talk")
        return
    sent = with_pictures(step, {PICTURE: {"shot": url}},
                         plan_for(line, found), LABEL,
                         allow=found.allow)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
