#!/usr/bin/env python3
"""
Step 7: refresh the Jobs and Layoffs slide.

Two layouts are handled.

The script's own topics (t-layoffs-fyi, t-trueup): the words
come from a skill as JSON, either part may be missing, and
each topic is refreshed on its own so filling one box does
not block the other:

    {"layoffs_fyi": {"bullets": ["Tech layoffs by year",
                                 "128.5K in 2026, ..."]},
     "trueup": {"bullets": ["In 2026: 187,160 people ..."]}}

A slide you built by hand: the script reads the numbers
itself and changes only four things you made.

    TrueUp box      "In 2026: 187,160 people laid off (737 per
                    day)" lines get TrueUp's current totals
    layoffs.fyi box "128.5K in 2026 (as of Sept 10, 2026)"
                    lines get layoffs.fyi's totals, rounded
                    the way each line already is, and the
                    date becomes today
    two pictures    each gets a fresh screenshot of its chart:
                    TrueUp's "Tech Employees Impacted by
                    Layoffs" and layoffs.fyi's "Recent Tech
                    Layoffs". Only the picture whose alt
                    text description contains "auto:
                    trueup-chart" or "auto:
                    layoffs-fyi-chart" is replaced; any
                    other picture, such as one you add, is
                    never touched.

A year a source no longer reports keeps its line as it is.
TrueUp sits behind a Cloudflare check, so its page is read
in a real Chrome window parked off-screen.

Usage:
    python3 g7_update_layoffs.py                 pictures, and
                                                 numbers on a
                                                 hand-made slide
    python3 g7_update_layoffs.py --json lay.json words too

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
    block_from, load_json, page_or_stop, report_frozen,
    start, with_pictures,
)
from sources import layoffs as SRC

LABEL = "layoffs"
HEADLINE = "Jobs and Layoffs"

# JSON part -> topic key, default headline, settings URL key.
TOPICS = [
    ("layoffs_fyi", T.TOPIC_LAYOFFS_FYI, "Layoffs.fyi tracker",
     "layoffs_fyi_url"),
    ("trueup", T.TOPIC_TRUEUP, "TrueUp tech layoff tracker",
     "trueup_url"),
]

# The number and its K are replaced together; K+ also mends
# a doubled "128.9KK".
FYI_LINE = re.compile(r"(\d+(?:\.(\d+))?K+) in (\d{4})")
TRUEUP_LINE = re.compile(
    r"In (\d{4}): ([\d,]+) people\s+laid off\s+\(([\d,]+) "
    r"per day\)")
TRUEUP_CHART = "Tech Employees Impacted by Layoffs"
FYI_CHART = "Recent Tech Layoffs"
# Alt text descriptions that mark hand-placed pictures as
# the two charts. Unmarked pictures are never replaced.
TRUEUP_MARK = "auto: trueup-chart"
FYI_MARK = "auto: layoffs-fyi-chart"
# The layoffs.fyi chart is 400px tall at any width; this
# width gives the 2.8 : 1 shape of its frame on the slide.
FYI_CHART_WIDTH = 1130
# It fetches its data after the page loads (Chart.js 4) and
# grows its bars in an animation headless Chrome may never
# finish. Once the data is in, draw it at once, unanimated.
FYI_CHART_READY = """(() => {
  const charts = Object.values((window.Chart || {}).instances
                               || {});
  const full = charts.filter(c =>
    c.data.datasets.some(d => d.data.length));
  full.forEach(c => { c.options.animation = false;
                      c.update('none'); });
  return full.length > 0;
})()"""


@dataclass
class Hand:
    """The four things step 7 changes on a hand-made slide."""

    page: object
    fyi_box: object = None
    trueup_box: object = None
    fyi_picture: str = ""
    trueup_picture: str = ""

    # --------------------------------------
    def allow(self):
        """Ids of the hand-made shapes this step may change."""
        ids = [self.fyi_picture, self.trueup_picture]
        ids += [b.id for b in (self.fyi_box, self.trueup_box)
                if b]
        return tuple(i for i in ids if i and not T.owned(i))


# --------------------------------------------------------------
def blocks_from(data, config):
    """Topic key -> Block for each part the JSON has."""
    found = {}
    for part, key, headline, url_key in TOPICS:
        if part in data:
            found[key] = block_from(data[part], headline,
                                    config[url_key])
    return found


# --------------------------------------------------------------
def topic_requests(deck, key, block, urls):
    """Words and picture for one topic, as far as allowed."""
    picture = T.shape_id(key, T.PICTURE)
    url = urls.get(picture)
    if block:
        return W.refresh_topic(deck, key, block, url)
    return W.refresh_picture(deck, picture, url)


# --------------------------------------------------------------
def find_page(deck):
    """The layoffs slide in the talk, by id or by title."""
    return G.titled_page(deck, T.PAGE_LAYOFFS, HEADLINE)


# --------------------------------------------------------------
def locate_hand(deck):
    """The hand-made layoffs slide's parts, or None."""
    page = find_page(deck)
    if not page or any(deck.shape(T.shape_id(key, T.BOX))
                       for _, key, _, _ in TOPICS):
        return None
    hand = Hand(page=page)
    for shape in page.text_shapes():
        if TRUEUP_LINE.search(shape.text):
            hand.trueup_box = hand.trueup_box or shape
        elif FYI_LINE.search(shape.text):
            hand.fyi_box = hand.fyi_box or shape
    if not (hand.fyi_box or hand.trueup_box):
        return None
    hand.fyi_picture = G.marked_picture(page, FYI_MARK)
    hand.trueup_picture = G.marked_picture(page, TRUEUP_MARK)
    return hand


# --------------------------------------------------------------
def thousands(count, decimals):
    """128873 -> 128.9K with one decimal, 129K with none."""
    return f"{count / 1000:.{decimals}f}K"


# --------------------------------------------------------------
def fyi_edits(text, totals, today):
    """Edits that bring the layoffs.fyi lines up to date."""
    edits = []
    for match in FYI_LINE.finditer(text):
        count = totals.get(int(match.group(3)))
        if count is None:
            continue
        decimals = len(match.group(2) or "")
        edits.append((match.start(1), match.end(1),
                      thousands(count, decimals)))
    date = W.date_edit(text, today) if edits else None
    return edits + ([date] if date else [])


# --------------------------------------------------------------
def trueup_edits(text, totals):
    """Edits that bring the TrueUp lines up to date."""
    edits = []
    for match in TRUEUP_LINE.finditer(text):
        found = totals.get(int(match.group(1)))
        if not found:
            continue
        people, per_day = found
        edits.append((match.start(2), match.end(2), people))
        edits.append((match.start(3), match.end(3), per_day))
    return edits


# --------------------------------------------------------------
def hand_plan_for(fyi, trueup, today):
    """Build the planning function for a hand-made slide."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URLs."""

        # ----------------------------------
        def plan(deck):
            """Numbers, date and pictures, in two batches."""
            hand = locate_hand(deck)
            if not hand:
                log(f"{LABEL}: the slide changed. Nothing done.")
                return []
            allow = hand.allow()
            batches = []
            for box, pic, edits in (
                    (hand.trueup_box, hand.trueup_picture,
                     lambda t: trueup_edits(t, trueup)),
                    (hand.fyi_box, hand.fyi_picture,
                     lambda t: fyi_edits(t, fyi, today))):
                reqs = W.refresh_picture(deck, pic,
                                         urls.get(pic), allow)
                if box:
                    reqs += W.edit_spans(box, edits(box.text))
                batches.append((box.id if box else pic, reqs))
            return batches

        return plan

    return make


# --------------------------------------------------------------
def plan_for(blocks):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URLs."""

        # ----------------------------------
        def plan(deck):
            """Refresh both topics independently."""
            if not page_or_stop(deck, T.PAGE_LAYOFFS, LABEL):
                return []
            batches = []
            for _, key, _, _ in TOPICS:
                report_frozen(deck, [T.shape_id(key, T.BOX)],
                              LABEL)
                batches.append((key, topic_requests(
                    deck, key, blocks.get(key), urls)))
            return batches

        return plan

    return make


# --------------------------------------------------------------
def run_hand(step, hand, config):
    """Read both sources and update a hand-made slide."""
    for mark, pic in ((TRUEUP_MARK, hand.trueup_picture),
                      (FYI_MARK, hand.fyi_picture)):
        if not pic:
            log(f"{LABEL}: no picture has the alt text "
                f"'{mark}'; that chart is left alone")
    fyi = SRC.fyi_totals() if hand.fyi_box else {}
    log(f"  layoffs.fyi: {fyi or 'not read'}")
    trueup = SRC.trueup_totals() if hand.trueup_box else {}
    log(f"  TrueUp: {trueup or 'not read'}")
    specs = {}
    if hand.trueup_picture:
        specs[hand.trueup_picture] = {
            "shot": config["trueup_url"],
            "element": TRUEUP_CHART, "visible": True}
    if hand.fyi_picture:
        specs[hand.fyi_picture] = {
            "shot": config["layoffs_fyi_chart_url"],
            "element": FYI_CHART, "width": FYI_CHART_WIDTH,
            "ready": FYI_CHART_READY}
    plan = hand_plan_for(fyi, trueup, dt.date.today())
    return with_pictures(step, specs, plan, LABEL,
                         allow=hand.allow())


# --------------------------------------------------------------
def main():
    """Refresh the layoffs slide of one deck."""
    step = start("Update the layoffs slide")
    config = settings()
    hand = locate_hand(G.read_deck(step.slides, step.deck_id))
    if hand:
        if step.args.json:
            log(f"{LABEL}: hand-made slide; --json is not used")
        sent = run_hand(step, hand, config)
    else:
        data = load_json(step.args.json) if step.args.json \
            else {}
        specs = {T.shape_id(key, T.PICTURE):
                 {"shot": config[url_key]}
                 for _, key, _, url_key in TOPICS}
        sent = with_pictures(step, specs,
                             plan_for(blocks_from(data, config)),
                             LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
