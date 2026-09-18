#!/usr/bin/env python3
"""
The shared start of every update step (g2 to g8).

Each step takes the same arguments and has to find the same
deck before it can do anything, so that lives here:

    [YYYY-MM-DD]     seminar date; default the coming
                     Friday, which on a Friday is today
                     until 3 pm US Eastern, then next
                     week's (deck.next_friday)
    --deck ID        use this presentation id directly
    --json FILE      content prepared by a skill
    --force          refresh even if nothing looks new

Usage:
    from gslides.step import start
    step = start("update benchmarks")
    step.drive, step.slides, step.deck_id, step.args

Created: 2026-09-14
Last updated: 2026-09-18
"""

import argparse
import json
import sys
from dataclasses import dataclass

from gslides import deck as G
from gslides import write as W
from layout.deck_parser import Block
from gslides.client import log, services, settings
from gslides.host import Host
from sources.pictures import prepare, remove_workdir, workdir


@dataclass
class Step:
    """Everything a step needs once the deck is found."""

    name: str
    args: object
    drive: object
    slides: object
    deck_id: str
    folder_id: str


# --------------------------------------------------------------
def parse(argv, name):
    """Read the common command line."""
    parser = argparse.ArgumentParser(description=name)
    parser.add_argument("date", nargs="?", default=None)
    parser.add_argument("--deck", default=None)
    parser.add_argument("--json", default=None)
    parser.add_argument("--force", action="store_true")
    return parser.parse_args(argv)


# --------------------------------------------------------------
def locate(drive, args):
    """The id of the deck to work on, or exit with help."""
    if args.deck:
        return args.deck
    date = G.seminar_date(args.date)
    found = G.find_decks(drive, date)
    if not found:
        log(f"ERROR: no deck for {date}. Create it first:")
        log(f"  python3 g1_new_deck.py {date}")
        sys.exit(1)
    if len(found) > 1:
        log(f"WARNING: {len(found)} decks tagged {date}; "
            f"using the newest, {found[0]['name']}")
    log(f"Deck: {found[0]['name']}")
    return found[0]["id"]


# --------------------------------------------------------------
def start(name, argv=None):
    """Parse arguments, sign in and find the deck."""
    args = parse(sys.argv[1:] if argv is None else argv, name)
    drive, slides = services()
    deck_id = locate(drive, args)
    return Step(name=name, args=args, drive=drive,
                slides=slides, deck_id=deck_id,
                folder_id=settings()["folder_id"])


# --------------------------------------------------------------
def load_json(path):
    """Read the content file a skill prepared."""
    if not path:
        log("ERROR: this step needs --json FILE")
        sys.exit(1)
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


# --------------------------------------------------------------
def page_or_stop(deck, slide_id, label):
    """The slide a step updates, if it can still be used."""
    page = deck.slide(slide_id)
    if page is None:
        log(f"{label}: slide {slide_id} is gone. Nothing done.")
        return None
    if page.index >= deck.separator():
        log(f"{label}: slide {slide_id} is parked. Nothing "
            f"done.")
        return None
    return page


# --------------------------------------------------------------
def report_frozen(deck, ids, label):
    """Say which of these shapes a human has frozen."""
    shapes = [deck.shape(i) for i in ids if deck.shape(i)]
    frozen = [s.id for s in shapes if not deck.editable(s)]
    if frozen:
        log(f"{label}: left alone (filled by you): "
            f"{', '.join(frozen)}")
    return frozen


# --------------------------------------------------------------
def host_pictures(step, specs, allow=()):
    """Prepare and upload the pictures that may change.

    specs maps a picture id to how to get it, such as
    {"t-aa-index-p": {"shot": url}}. A picture whose box you
    filled is skipped before any screenshot is taken.
    """
    deck = G.read_deck(step.slides, step.deck_id)
    host = Host(step.drive, step.folder_id)
    folder = workdir()
    urls = {}
    for image_id, spec in specs.items():
        shape = deck.shape(image_id)
        if not shape or not (deck.editable(shape)
                             or image_id in allow):
            log(f"  picture {image_id}: left alone")
            continue
        log(f"  picture {image_id}: fetching")
        path = prepare(spec, shape.rect.w, folder)
        if not path:
            log(f"  picture {image_id}: not available, the "
                f"old picture stays")
            continue
        urls[image_id] = host.put(path)
    return urls, host, folder


# --------------------------------------------------------------
def with_pictures(step, specs, make_plan, label, allow=()):
    """Host pictures, run the plan, and always clean up."""
    urls, host, folder = host_pictures(step, specs, allow)
    try:
        return W.run_plan(step.slides, step.deck_id,
                          make_plan(urls), label, allow=allow)
    finally:
        host.clean()
        remove_workdir(folder)


# --------------------------------------------------------------
def block_from(data, headline, source_url=None):
    """A Block from a skill's JSON, ending with its source."""
    bullets = [b for b in data.get("bullets", []) if b.strip()]
    if source_url and source_url not in bullets:
        bullets.append(source_url)
    return Block(headline=data.get("headline", headline),
                 bullets=bullets)
