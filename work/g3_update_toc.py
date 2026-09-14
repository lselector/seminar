#!/usr/bin/env python3
"""
Step 3: rebuild the table of contents on the first slide.

Walks the talk from the first slide to the separator, in the
order the slides are in right now, and collects the first
line of every text box: the script's and yours alike. Titles
and empty placeholders are skipped, and so are the slides
the archive decks leave out of the contents: Benchmarks, the
YouTube promo, About the Speaker and Thank You.

The list is written as up to three columns on s-toc. If you
filled any TOC column, the whole TOC is yours and is left
alone, so the columns never disagree with each other.

The epigraph is written only when a skill supplies one and
the epigraph box still holds its placeholder unfilled:

    {"epigraph": "Agents got their own computers."}

Usage:
    python3 g3_update_toc.py                  next Friday
    python3 g3_update_toc.py 2026-09-18
    python3 g3_update_toc.py --json epi.json  with an epigraph

Created: 2026-09-14
Last updated: 2026-09-14
"""

from layout import deck_layout as L
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import api as S
from gslides import ids as T
from gslides.client import log
from gslides.step import load_json, page_or_stop, start

LABEL = "table of contents"
MAX_COLUMNS = 3
EPIGRAPH_BOX = T.shape_id(T.PAGE_TOC, "epi")

# Pages the archive decks keep out of the contents.
NOT_IN_TOC = {T.PAGE_TOC, T.PAGE_BENCH, T.PAGE_YOUTUBE,
              T.PAGE_ABOUT, T.PAGE_THANKS, T.PAGE_SEPARATOR,
              T.PAGE_PARKED}


# --------------------------------------------------------------
def column_ids():
    """Every id a TOC column may have."""
    return [T.shape_id(T.PAGE_TOC, f"c{i}")
            for i in range(MAX_COLUMNS)]


# --------------------------------------------------------------
def headline(shape):
    """The line a box contributes to the contents."""
    line = shape.first_line()
    if line.startswith(R.BULLET.strip()):
        line = line[len(R.BULLET.strip()):].strip()
    return line


# --------------------------------------------------------------
def collect(deck):
    """Headlines of the talk, in the current slide order."""
    names = []
    for slide in deck.main():
        if slide.id in NOT_IN_TOC:
            continue
        for shape in slide.text_shapes():
            if G.is_title(shape) or \
                    G.is_placeholder(shape.text):
                continue
            line = headline(shape)
            if line and line not in names:
                names.append(line)
    return names


# --------------------------------------------------------------
def columns_top(deck):
    """Where the columns start: under the epigraph."""
    epi = deck.shape(EPIGRAPH_BOX)
    if not epi:
        return L.BAND_TOP
    return max(L.BAND_TOP, epi.rect.y + epi.rect.h + 0.10)


# --------------------------------------------------------------
def toc_requests(deck, names):
    """Rebuild the columns, or nothing if any is frozen."""
    existing = [deck.shape(i) for i in column_ids()
                if deck.shape(i)]
    frozen = [s.id for s in existing if not deck.editable(s)]
    if frozen:
        log(f"{LABEL}: left alone, you filled "
            f"{', '.join(frozen)}")
        return []
    sid = T.PAGE_TOC
    wanted = R.render_toc_columns(sid, sid, names,
                                  columns_top(deck), fill=None)
    reqs = W.replace_owned(deck, sid, wanted, sweep=False)
    new_ids = set(W.created(wanted))
    for shape in existing:
        if shape.id not in new_ids:
            reqs = S.delete_object(shape.id) + reqs
    return reqs


# --------------------------------------------------------------
def epigraph_requests(deck, text):
    """Write the epigraph into a still-empty epigraph box."""
    box = deck.shape(EPIGRAPH_BOX)
    if not text or not box:
        return []
    if not G.is_placeholder(box.text):
        log(f"{LABEL}: epigraph already written, kept")
        return []
    body = R.plain_body(text, R.EPIGRAPH_SIZE, colour=R.RED,
                        bold=True, italic=True)
    return W.refresh_text(deck, EPIGRAPH_BOX, body)


# --------------------------------------------------------------
def plan_for(epigraph):
    """Build the planning function for run_plan."""

    # --------------------------------------
    def plan(deck):
        """Contents and, if given, the epigraph."""
        if not page_or_stop(deck, T.PAGE_TOC, LABEL):
            return []
        names = collect(deck)
        log(f"{LABEL}: {len(names)} headlines")
        for name in names:
            log(f"    {name}")
        return [("epigraph", epigraph_requests(deck, epigraph)),
                (LABEL, toc_requests(deck, names))]

    return plan


# --------------------------------------------------------------
def main():
    """Rebuild the contents of one deck."""
    step = start("Update the table of contents")
    extra = load_json(step.args.json) if step.args.json else {}
    sent = W.run_plan(step.slides, step.deck_id,
                      plan_for(extra.get("epigraph")), LABEL)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
