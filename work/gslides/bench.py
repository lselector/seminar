#!/usr/bin/env python3
"""
Find the LM Arena benchmarks page in a deck.

The benchmarks page holds two tables whose header row reads
Code | Model | Score, one per board (English, Coding). The
author designed that page by hand, and each new deck starts
from a copy of the last one, so the page is found by its
tables, not by an id. A deck made before that carries the
script's own column page, s-bench, instead.

Step 1 uses this to pick page 2 out of last week's deck;
step 4 uses it to find the tables it fills.

Usage:
    from gslides import bench as K
    page = K.board_page(deck)       # tables, else s-bench
    tables = K.board_tables(page)

Created: 2026-09-18
Last updated: 2026-09-18
"""

from gslides import deck as G
from gslides import ids as T

HEADER = ["code", "model", "score"]


# --------------------------------------------------------------
def is_board_table(shape):
    """A table whose header row is Code | Model | Score."""
    if shape.kind != G.KIND_TABLE or not shape.cells:
        return False
    head = [c.strip().lower() for c in shape.cells[0][:3]]
    return head == HEADER


# --------------------------------------------------------------
def board_tables(page):
    """The board tables on one slide, left to right."""
    return sorted((s for s in page.shapes if is_board_table(s)),
                  key=lambda s: s.rect.x)


# --------------------------------------------------------------
def table_page(deck):
    """A slide in the talk holding board tables, or None."""
    for slide in deck.main():
        if board_tables(slide):
            return slide
    return None


# --------------------------------------------------------------
def board_page(deck):
    """The benchmarks page: the table page, else s-bench."""
    page = table_page(deck)
    if page:
        return page
    bench = deck.slide(T.PAGE_BENCH)
    if bench and bench.index < deck.separator():
        return bench
    return None
