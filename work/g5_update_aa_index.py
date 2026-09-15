#!/usr/bin/env python3
"""
Step 5: refresh the Artificial Analysis Intelligence Index slide.

Screenshots the "Cost per Intelligence Index Task" chart on
the Intelligence Index page (only the chart, not the whole
page) and swaps it into the slide's picture frame. A date
box on the slide, a box holding only a date such as
"Sept 10", is set to today's date.

The slide is found by its id (s-aa-index) or, if you rebuilt
it by hand, by its title "Artificial Analysis Intelligence
Index". On a hand-made slide this step changes exactly two
things you made, even if they are filled, and nothing else:
the date box, and the picture whose alt text description
contains "auto: aa-index-chart". Pictures without that mark,
such as ones you add, are never replaced.

If a skill has written new bullets, they replace the words
in the script's text box too:

    {"headline": "Artificial Analysis Intelligence Index",
     "bullets": ["Claude Fable 5 leads at 73", "..."]}

The page link is kept as the last bullet. The box and its
picture stay as they are once you fill the box.

Usage:
    python3 g5_update_aa_index.py                 chart, date
    python3 g5_update_aa_index.py 2026-09-18
    python3 g5_update_aa_index.py --json aa.json  words too

Created: 2026-09-14
Last updated: 2026-09-14
"""

import datetime as dt
import re
from dataclasses import dataclass

from gslides import deck as G
from gslides import write as W
from gslides import ids as T
from gslides.client import log, settings
from gslides.step import (
    block_from, load_json, report_frozen, start, with_pictures,
)

LABEL = "intelligence index"
HEADLINE = "Artificial Analysis Intelligence Index"
CHART = "Cost per Intelligence Index Task"
# Alt text description that marks a hand-placed picture as
# the chart. Unmarked pictures are never replaced.
CHART_MARK = "auto: aa-index-chart"
BOX = T.shape_id(T.TOPIC_AA, T.BOX)
PICTURE = T.shape_id(T.TOPIC_AA, T.PICTURE)
DATE = re.compile(
    r"^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
    r"[a-z]*\.? \d{1,2}(, \d{4})?$", re.IGNORECASE)


@dataclass
class Target:
    """What this step changes on the slide it found."""

    page: object
    picture: str = ""
    date: object = None
    allow: tuple = ()


# --------------------------------------------------------------
def find_page(deck):
    """The Intelligence Index slide in the talk, or None."""
    page = deck.slide(T.PAGE_AA)
    if page and page.index < deck.separator():
        return page
    want = HEADLINE.lower()
    for slide in deck.main():
        for shape in slide.text_shapes():
            if G.is_title(shape) and \
                    shape.first_line().lower() == want:
                return slide
    return None


# --------------------------------------------------------------
def find_picture(page):
    """The script's picture, or the one marked as the chart."""
    if any(s.id == PICTURE for s in page.shapes):
        return PICTURE
    return G.marked_picture(page, CHART_MARK)


# --------------------------------------------------------------
def find_date(page):
    """A box whose whole text is a date, or None."""
    for shape in page.text_shapes():
        if DATE.match(shape.text.strip()):
            return shape
    return None


# --------------------------------------------------------------
def locate(deck):
    """The slide, its picture and date box, and what to allow."""
    page = find_page(deck)
    if not page:
        return None
    target = Target(page=page, picture=find_picture(page),
                    date=find_date(page))
    mine = [target.picture] + \
        ([target.date.id] if target.date else [])
    target.allow = tuple(i for i in mine
                         if i and not T.owned(i))
    return target


# --------------------------------------------------------------
def plan_for(block, found, today):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URL."""

        # ----------------------------------
        def plan(deck):
            """Refresh the words, the picture and the date."""
            target = locate(deck)
            if not target or target.page.id != found.page.id:
                log(f"{LABEL}: the slide changed. Nothing done.")
                return []
            url = urls.get(target.picture)
            if block and deck.shape(BOX):
                report_frozen(deck, [BOX], LABEL)
                reqs = W.refresh_topic(deck, T.TOPIC_AA,
                                       block, url)
            else:
                reqs = W.refresh_picture(deck, target.picture,
                                         url, target.allow)
            reqs += W.replace_date(target.date, today)
            return [(LABEL, reqs)]

        return plan

    return make


# --------------------------------------------------------------
def main():
    """Refresh the Intelligence Index slide of one deck."""
    step = start("Update the Intelligence Index slide")
    url = settings()["aa_index_url"]
    found = locate(G.read_deck(step.slides, step.deck_id))
    if not found:
        log(f"{LABEL}: no Intelligence Index slide in the talk")
        return
    if not found.picture:
        log(f"{LABEL}: no picture has the alt text "
            f"'{CHART_MARK}'; pictures left alone")
    block = None
    if step.args.json:
        block = block_from(load_json(step.args.json),
                           HEADLINE, url)
    specs = {found.picture: {"shot": url, "element": CHART}} \
        if found.picture else {}
    sent = with_pictures(step, specs,
                         plan_for(block, found, dt.date.today()),
                         LABEL, allow=found.allow)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
