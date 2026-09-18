#!/usr/bin/env python3
"""
Step 1: create the Google Slides deck for one seminar.

Builds the eleven-page skeleton described in ADD.md in the
folder named in config/gslides.json:

     1 s-toc        title, TOC and epigraph placeholders
     2 (copied)     benchmarks, last week's page 2 *
     3 (copied)     Intelligence Index, last week's page *
     4 s-news-1     "AI News" with one placeholder box
     5 (copied)     the channel promo, last week's page *
     6 s-news-2     "AI News" with one placeholder box
     7 (copied)     jobs and layoffs, last week's page *
     8 s-about      About the Speaker
     9 s-thanks     Thank You!
    10 s-separator  everything after this is left out
    11 s-parked     where rejected topics go

* Pages 2, 3, 5 and 7 are the author's designs. Page 2 has
two Code | Model | Score tables, captions, legend, the Elo
note and the model sizes; the Slides API cannot draw a table
that compact (it cannot set cell padding). Page 3 has the
author's notes, a date box and the chart marked
auto: aa-index-chart. Page 5 has the author's promo box and
channel screenshots. Page 7 has the author's layoffs.fyi and
TrueUp boxes and two marked charts. So the new deck starts
as a Drive copy of the latest earlier deck with every slide
but those four deleted; the other pages are drawn around
them, and steps 4 to 7 put this week's numbers, dates and
charts in. With no earlier deck, they are the script's own
pages: s-bench, s-aa-index, s-youtube and s-layoffs.

The wording of pages 3, 5, 7, 8 and 9 comes from
config/skeleton.json, the one place that text lives.
Their first pictures are the ones in work/assets;
the update steps refresh them later.

Script boxes get a red border and no fill. Fill a box by hand
and every script leaves it alone from then on.

The deck is tagged with its date in Drive appProperties, so
the other steps find it even after it is renamed or moved.
If a deck for that date already exists, this stops and
prints its link: it never replaces a deck.

Usage:
    python3 g1_new_deck.py              this week's Friday *
    python3 g1_new_deck.py 2026-09-25   a given date

* Today on a Friday until 3 pm US Eastern, when the
  seminar ends; from then on, the Friday after.

Created: 2026-09-14
Last updated: 2026-09-18
"""

import json
import os
import sys

from layout import bench_page as B
from layout import deck_layout as L
from layout import skeleton as K
from gslides import bench as BK
from gslides import deck as G
from gslides import render as R
from gslides import api as S
from gslides import ids as T
from gslides.client import (
    DATA_DIR, ROOT, log, services, settings,
)
from gslides.host import Host

LEADERBOARD = os.path.join(DATA_DIR, "leaderboard.json")

# config/skeleton.json entry -> fixed page id and topic keys.
FIXED_PAGES = {
    "aa_index": (T.PAGE_AA, [T.TOPIC_AA]),
    "youtube": (T.PAGE_YOUTUBE, [T.TOPIC_YOUTUBE]),
    "layoffs": (T.PAGE_LAYOFFS,
                [T.TOPIC_LAYOFFS_FYI, T.TOPIC_TRUEUP]),
    "about": (T.PAGE_ABOUT, [T.TOPIC_ABOUT]),
    "thanks": (T.PAGE_THANKS, [T.TOPIC_THANKS]),
}
NEWS_PAGES = [(T.PAGE_NEWS_1, T.TOPIC_NEWS_1),
              (T.PAGE_NEWS_2, T.TOPIC_NEWS_2)]

# Pages 3 to 9, in order, after contents and benchmarks.
MIDDLE = ["aa_index", NEWS_PAGES[0], "youtube",
          NEWS_PAGES[1], "layoffs", "about", "thanks"]

# The author's own pages carried over from last week's deck,
# found by fixed id or by title, and the drawn page each one
# goes in front of. Page 2 is found by its tables instead.
TITLED = [("aa_index", T.PAGE_AA, T.PAGE_NEWS_1),
          ("youtube", T.PAGE_YOUTUBE, T.PAGE_NEWS_2),
          ("layoffs", T.PAGE_LAYOFFS, T.PAGE_ABOUT)]

NEWS_TITLE = "AI News"
SEPARATOR_TITLE = "Not in the presentation"
SEPARATOR_TEXT = (
    "Everything after this slide is left out of the talk.\n"
    "Paste topics you decided not to present on the next "
    "slide. Scripts never write past this point, and a "
    "topic parked here is never added again."
)
PARKED_TITLE = "Parked topics"
HINT_SIZE = 14


# --------------------------------------------------------------
def placeholder(box_id, slide_id, rect, hint, size=HINT_SIZE,
                colour=R.GREY, italic=True):
    """A bordered, unfilled box holding a bracketed hint."""
    body = R.plain_body(hint, size, colour=colour,
                        italic=italic)
    reqs = R.make_box(box_id, slide_id, R.fit_rect(rect, body),
                      border=True)
    return reqs + R.write_body(box_id, body)


# --------------------------------------------------------------
def toc_page(title):
    """Page 1: the title, a TOC box and an epigraph box."""
    sid = T.PAGE_TOC
    limit = R.EPIGRAPH_X - L.MARGIN - 0.10
    reqs = R.render_title(sid, sid, title, limit)
    epi = R.epigraph_rect(G.HINT_EPIGRAPH)
    reqs += placeholder(T.shape_id(sid, "epi"), sid, epi,
                        G.HINT_EPIGRAPH, R.EPIGRAPH_SIZE - 4,
                        colour=R.RED)
    top = max(L.BAND_TOP, epi.y + epi.h + 0.10)
    toc = L.Rect(L.MARGIN, top, 3.2, 0.5)
    reqs += placeholder(T.shape_id(sid, "c0"), sid, toc,
                        G.HINT_TOC)
    return sid, reqs


# --------------------------------------------------------------
def bench_page():
    """Page 2: benchmarks from the cached leaderboard."""
    data = None
    if os.path.isfile(LEADERBOARD):
        with open(LEADERBOARD, encoding="utf-8") as handle:
            data = json.load(handle)
    else:
        log("  no leaderboard.json yet; benchmarks start empty")
    sid = T.PAGE_BENCH
    return sid, R.render_bench(sid, sid, data, B, fill=None)


# --------------------------------------------------------------
def news_page(sid, key):
    """Pages 4 and 6: a title and one empty news box."""
    reqs = R.render_title(sid, sid, NEWS_TITLE)
    rect = L.Rect(L.MARGIN, L.BAND_TOP + 0.30,
                  L.CONTENT_W - L.IMG_W - L.COL_GAP, 0.50)
    reqs += placeholder(T.shape_id(key, T.BOX), sid, rect,
                        G.HINT_NEWS)
    return sid, reqs


# --------------------------------------------------------------
def fixed_page(name, entry, locate):
    """Pages 3, 5, 7, 8, 9: from config/skeleton.json."""
    sid, keys = FIXED_PAGES[name]
    pages = L.paginate([K.section(entry, ROOT)])
    if not pages:
        return sid, R.render_title(sid, sid, entry["title"])
    return sid, R.render_content(sid, sid, pages[0], keys,
                                 fill=None, locate=locate)


# --------------------------------------------------------------
def closing_pages():
    """Pages 10 and 11: the separator and the parked page."""
    sep = T.PAGE_SEPARATOR
    reqs = R.render_title(sep, sep, SEPARATOR_TITLE)
    rect = L.Rect(L.MARGIN, L.BAND_TOP + 0.40,
                  L.CONTENT_W * 0.7, 1.40)
    box = T.shape_id(sep, "note")
    note = R.plain_body(SEPARATOR_TEXT, 18, colour=R.RED,
                        bold=True)
    reqs += R.make_box(box, sep, R.fit_rect(rect, note),
                       border=True)
    reqs += R.write_body(box, note)
    parked = T.PAGE_PARKED
    return [(sep, reqs),
            (parked, R.render_title(parked, parked,
                                    PARKED_TITLE))]


# --------------------------------------------------------------
def build_pages(title, locate, skip=()):
    """Every skeleton page as (slide id, requests).

    skip names pages ("bench", "aa_index", "youtube",
    "layoffs") copied from last week's deck, not drawn.
    """
    static = K.load()
    pages = [toc_page(title)]
    if "bench" not in skip:
        pages.append(bench_page())
    for entry in MIDDLE:
        if isinstance(entry, tuple):
            pages.append(news_page(*entry))
        elif entry not in skip:
            pages.append(fixed_page(entry, static[entry],
                                    locate))
    return pages + closing_pages()


# --------------------------------------------------------------
def create_file(drive, date, folder):
    """Create the empty, date-tagged presentation."""
    made = drive.files().create(body={
        "name": G.deck_name(date),
        "mimeType": G.DECK_MIME,
        "parents": [folder],
        "appProperties": {G.DATE_KEY: date.isoformat()},
    }, fields="id, webViewLink").execute()
    return made["id"], made.get("webViewLink", "")


# --------------------------------------------------------------
def carried_pages(deck):
    """Last week's pages to keep, as {key: slide id}."""
    found = {}
    bench = BK.board_page(deck)
    if bench:
        found["bench"] = bench.id
    for key, slide_id, _ in TITLED:
        title = K.page(key)["title"]
        page = G.titled_page(deck, slide_id, title)
        if page:
            found[key] = page.id
    return found


# --------------------------------------------------------------
def only_pages(deck, keep_ids):
    """Requests that delete every slide but the kept ones."""
    reqs = []
    for slide in deck.slides:
        if slide.id not in keep_ids:
            reqs += S.delete_object(slide.id)
    return reqs


# --------------------------------------------------------------
def copy_previous(drive, slides, date, folder):
    """Start from last week's deck, keeping its own pages.

    Returns (deck id, link, {key: slide id}), or None when
    there is no earlier deck or it has neither page. The
    copy keeps the author's tables, boxes, charts, alt text
    marks and styles exactly; steps 4 and 7 then put this
    week's numbers in.
    """
    prev = G.find_previous_deck(drive, date)
    if not prev:
        return None
    made = drive.files().copy(fileId=prev["id"], body={
        "name": G.deck_name(date),
        "parents": [folder],
        "appProperties": {G.DATE_KEY: date.isoformat()},
    }, fields="id, webViewLink").execute()
    deck = G.read_deck(slides, made["id"])
    kept = carried_pages(deck)
    if not kept:
        drive.files().delete(fileId=made["id"]).execute()
        log(f"  {prev['name']} has no page to carry over")
        return None
    slides.presentations().batchUpdate(
        presentationId=made["id"],
        body={"requests": only_pages(deck, kept.values())},
    ).execute()
    log(f"  copied {', '.join(kept)} from {prev['name']}")
    return made["id"], made.get("webViewLink", ""), kept


# --------------------------------------------------------------
def send_pages(slides, deck_id, pages):
    """One batch per drawn page, in order."""
    for index, (sid, reqs) in enumerate(pages):
        head = S.new_slide(sid)
        head[0]["createSlide"]["insertionIndex"] = index
        slides.presentations().batchUpdate(
            presentationId=deck_id,
            body={"requests": head + reqs}).execute()
        log(f"  drew {sid}")


# --------------------------------------------------------------
def final_order(page_ids, kept):
    """Drawn pages with the carried ones back in place."""
    order = list(page_ids)
    if "bench" in kept:
        order.insert(1, kept["bench"])
    for key, _, before in TITLED:
        if key in kept:
            order.insert(order.index(before), kept[key])
    return order


# --------------------------------------------------------------
def arrange(slides, deck_id, order):
    """Drop the blank start slide; put every slide in place.

    Moving each slide to its index, front to back, leaves
    the deck in exactly this order.
    """
    deck = G.read_deck(slides, deck_id)
    reqs = []
    for slide in deck.slides:
        if slide.id not in order:
            reqs += S.delete_object(slide.id)
    for index, slide_id in enumerate(order):
        reqs += S.move_slide(slide_id, index)
    slides.presentations().batchUpdate(
        presentationId=deck_id,
        body={"requests": reqs}).execute()
    for index, slide_id in enumerate(order):
        log(f"  {index + 1:2d} {slide_id}")


# --------------------------------------------------------------
def start_deck(drive, slides, date, folder):
    """The new file: last week's carried pages, else empty."""
    copied = copy_previous(drive, slides, date, folder)
    if copied:
        return copied
    deck_id, link = create_file(drive, date, folder)
    return deck_id, link, {}


# --------------------------------------------------------------
def refuse_if_exists(drive, date):
    """Stop when a deck for this date is already there."""
    found = G.find_decks(drive, date)
    if not found:
        return
    log(f"A deck for {date} already exists. Nothing changed.")
    for item in found:
        log(f"  {item['name']}  {item.get('webViewLink', '')}")
    sys.exit(0)


# --------------------------------------------------------------
def main():
    """Create one seminar deck from the skeleton."""
    date = G.seminar_date(sys.argv[1] if len(sys.argv) > 1
                          else None)
    drive, slides = services()
    folder = settings()["folder_id"]
    refuse_if_exists(drive, date)

    deck_id, link, kept = start_deck(drive, slides, date,
                                     folder)
    log(f"Created {G.deck_name(date)}")
    host = Host(drive, folder)
    try:
        pages = build_pages(G.deck_title(date), host.put,
                            skip=kept)
        send_pages(slides, deck_id, pages)
        order = final_order([sid for sid, _ in pages], kept)
        arrange(slides, deck_id, order)
    finally:
        host.clean()
    log("=" * 50)
    log(f"Deck ready: {len(order)} slides")
    if kept:
        log("Copied pages still show last week's numbers: run "
            "g4, g5, g6 and g7")
    log(link)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
