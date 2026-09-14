#!/usr/bin/env python3
"""
Step 7: refresh the Jobs and Layoffs slide.

The numbers come from a skill that reads layoffs.fyi and
TrueUp and writes them as JSON; either part may be missing:

    {"layoffs_fyi": {"bullets": ["Tech layoffs by year",
                                 "128.5K in 2026, ..."]},
     "trueup": {"bullets": ["In 2026: 187,160 people ..."]}}

Each topic is refreshed on its own, so filling one box does
not block the other. Both screenshots are retaken; TrueUp
sits behind a Cloudflare check, and when the check is there
its old picture is kept rather than replaced by the "security
verification" page.

Usage:
    python3 g7_update_layoffs.py                 pictures only
    python3 g7_update_layoffs.py --json lay.json words too

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

LABEL = "layoffs"

# JSON part -> topic key, default headline, settings URL key.
TOPICS = [
    ("layoffs_fyi", T.TOPIC_LAYOFFS_FYI, "Layoffs.fyi tracker",
     "layoffs_fyi_url"),
    ("trueup", T.TOPIC_TRUEUP, "TrueUp tech layoff tracker",
     "trueup_url"),
]


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
def main():
    """Refresh the layoffs slide of one deck."""
    step = start("Update the layoffs slide")
    config = settings()
    data = load_json(step.args.json) if step.args.json else {}
    blocks = blocks_from(data, config)
    specs = {T.shape_id(key, T.PICTURE):
             {"shot": config[url_key]}
             for _, key, _, url_key in TOPICS}
    sent = with_pictures(step, specs, plan_for(blocks), LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
