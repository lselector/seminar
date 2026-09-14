#!/usr/bin/env python3
"""
Tests for gslides/deck.py, the ownership rules of a live deck.

Runs on fixtures/skeleton_deck.json, a real presentations.get
response saved from a deck g1_new_deck.py built. Human edits
are simulated by changing that JSON the way the API would
report them.

Usage:
    python3 -m tests.test_gdeck

Created: 2026-09-14
Last updated: 2026-09-14
"""

import copy
import datetime as dt
import json
import os

from gslides import deck as G
from gslides import ids as T
from tests.runner import run

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURE = os.path.join(HERE, "fixtures", "skeleton_deck.json")

with open(FIXTURE, encoding="utf-8") as _handle:
    RAW = json.load(_handle)


# --------------------------------------------------------------
def raw():
    """A fresh copy of the saved deck JSON."""
    return copy.deepcopy(RAW)


# --------------------------------------------------------------
def element(data, object_id):
    """The JSON of one page element, found by id."""
    for slide in data["slides"]:
        for item in slide.get("pageElements", []):
            if item["objectId"] == object_id:
                return item
    raise KeyError(object_id)


# --------------------------------------------------------------
def fill(data, object_id):
    """Give a box a yellow fill, as the Slides UI would."""
    props = element(data, object_id)["shape"].setdefault(
        "shapeProperties", {})
    props["shapeBackgroundFill"] = {"solidFill": {
        "color": {"rgbColor": {"red": 1, "green": 0.95,
                               "blue": 0.8}}, "alpha": 1}}
    return data


# --------------------------------------------------------------
def add_human_box(data, slide_id, text="my own note"):
    """Add a text box with a Google-assigned id."""
    for slide in data["slides"]:
        if slide["objectId"] == slide_id:
            slide["pageElements"].append({
                "objectId": "SLIDES_API123456_0",
                "size": {"width": {"magnitude": 914400,
                                   "unit": "EMU"},
                         "height": {"magnitude": 914400,
                                    "unit": "EMU"}},
                "transform": {"scaleX": 1, "scaleY": 1,
                              "translateX": 0, "translateY": 0,
                              "unit": "EMU"},
                "shape": {"shapeType": "TEXT_BOX", "text": {
                    "textElements": [{"textRun": {
                        "content": text + "\n"}}]}},
            })
    return data


# --------------------------------------------------------------
def move_slide(data, slide_id, index):
    """Move one slide to a new place in the deck."""
    slides = data["slides"]
    found = next(s for s in slides if s["objectId"] == slide_id)
    slides.remove(found)
    slides.insert(index, found)
    return data


# --------------------------------------------------------------
def test_skeleton_has_the_eleven_fixed_pages_in_order():
    """g1 built exactly the pages the plan lists."""
    deck = G.parse_deck(raw())
    assert [s.id for s in deck.slides] == [
        T.PAGE_TOC, T.PAGE_BENCH, T.PAGE_AA, T.PAGE_NEWS_1,
        T.PAGE_YOUTUBE, T.PAGE_NEWS_2, T.PAGE_LAYOFFS,
        T.PAGE_ABOUT, T.PAGE_THANKS, T.PAGE_SEPARATOR,
        T.PAGE_PARKED]


# --------------------------------------------------------------
def test_unfilled_boxes_read_as_unfilled():
    """propertyState NOT_RENDERED means no fill."""
    deck = G.parse_deck(raw())
    box = deck.shape("t-news-1-b")
    assert box.kind == G.KIND_TEXT and not box.filled
    assert box.text == G.HINT_NEWS, box.text


# --------------------------------------------------------------
def test_script_box_without_fill_is_editable():
    """The basic case: ours, unfilled, before the separator."""
    deck = G.parse_deck(raw())
    assert deck.editable(deck.shape("t-aa-index-b"))


# --------------------------------------------------------------
def test_a_fill_freezes_the_box():
    """Any fill means a human accepted it."""
    deck = G.parse_deck(fill(raw(), "t-aa-index-b"))
    box = deck.shape("t-aa-index-b")
    assert box.filled and deck.frozen(box)
    assert not deck.editable(box)


# --------------------------------------------------------------
def test_a_picture_follows_its_text_box():
    """Filling the box freezes its picture too."""
    deck = G.parse_deck(raw())
    assert deck.editable(deck.shape("t-aa-index-p"))
    deck = G.parse_deck(fill(raw(), "t-aa-index-b"))
    assert not deck.editable(deck.shape("t-aa-index-p"))


# --------------------------------------------------------------
def test_freezing_one_layoffs_box_leaves_the_other():
    """Two topics on a page freeze independently."""
    deck = G.parse_deck(fill(raw(), "t-trueup-b"))
    assert not deck.editable(deck.shape("t-trueup-b"))
    assert deck.editable(deck.shape("t-layoffs-fyi-b"))


# --------------------------------------------------------------
def test_human_boxes_are_never_editable():
    """A Google-assigned id belongs to the human."""
    deck = G.parse_deck(add_human_box(raw(), T.PAGE_NEWS_1))
    human = deck.shape("SLIDES_API123456_0")
    assert human.text == "my own note"
    assert not deck.editable(human)


# --------------------------------------------------------------
def test_a_slide_moved_past_the_separator_is_parked():
    """Script shapes on a parked slide are off limits."""
    data = move_slide(raw(), T.PAGE_NEWS_1, 10)
    deck = G.parse_deck(data)
    assert T.PAGE_NEWS_1 in [s.id for s in deck.parked()]
    assert not deck.editable(deck.shape("t-news-1-b"))


# --------------------------------------------------------------
def test_the_separator_itself_is_not_editable():
    """Its own note box sits at the boundary, not before."""
    deck = G.parse_deck(raw())
    assert not deck.editable(deck.shape("s-separator-note"))


# --------------------------------------------------------------
def test_without_a_separator_everything_is_main():
    """A deleted separator does not hide the whole deck."""
    data = raw()
    data["slides"] = [s for s in data["slides"]
                      if s["objectId"] != T.PAGE_SEPARATOR]
    deck = G.parse_deck(data)
    assert len(deck.main()) == len(deck.slides)
    assert deck.parked() == []


# --------------------------------------------------------------
def test_reordering_changes_order_but_not_ownership():
    """Moving a slide within the talk keeps it editable."""
    deck = G.parse_deck(move_slide(raw(), T.PAGE_LAYOFFS, 2))
    assert deck.slides[2].id == T.PAGE_LAYOFFS
    assert deck.editable(deck.shape("t-trueup-b"))


# --------------------------------------------------------------
def test_speaker_notes_id_is_read():
    """Every slide exposes where its notes go."""
    deck = G.parse_deck(raw())
    assert deck.slide(T.PAGE_SEPARATOR).notes_id


# --------------------------------------------------------------
def test_rectangles_come_back_in_inches():
    """A box's position matches what g1 asked for.

    Google stores the box as a 3,000,000 EMU square with a
    scale transform, so this also proves the scale is used.
    """
    from layout import deck_layout as L
    from gslides import render as R
    deck = G.parse_deck(raw())
    title = deck.shape("s-news-1-title")
    want_x = L.MARGIN - R.PAD_X
    want_y = L.TITLE_Y - R.PAD_Y
    want_h = L.TITLE_H + 2 * R.PAD_Y
    assert abs(title.rect.x - want_x) < 0.001, title.rect
    assert abs(title.rect.y - want_y) < 0.001, title.rect
    assert abs(title.rect.h - want_h) < 0.001, title.rect


# --------------------------------------------------------------
def test_placeholder_detection():
    """A bracketed hint or nothing at all is still empty."""
    assert G.is_placeholder(G.HINT_NEWS)
    assert G.is_placeholder("")
    assert G.is_placeholder("  [anything]  ")
    assert not G.is_placeholder("Mistral raises a round")
    assert not G.is_placeholder("[a]\nreal text")


# --------------------------------------------------------------
def test_seminar_date_rules():
    """Next Friday, or today when today is Friday."""
    monday = dt.date(2026, 9, 14)
    friday = dt.date(2026, 9, 18)
    assert G.seminar_date(None, monday) == friday
    assert G.seminar_date(None, friday) == friday
    assert G.seminar_date("2026-10-02") == dt.date(2026, 10, 2)


# --------------------------------------------------------------
def test_deck_title_and_name_match_the_archive():
    """The archive's month spellings are kept."""
    date = dt.date(2026, 9, 18)
    assert G.deck_title(date) == "AI News - Sept 18, 2026"
    assert G.deck_name(date) == "2026-09-18-AI-News"


# --------------------------------------------------------------
def test_find_decks_searches_the_hidden_tag():
    """The lookup uses appProperties, not the file name."""
    seen = {}

    class Files:
        """Stand-in for drive.files()."""

        # ----------------------------------
        def list(self, **kwargs):
            """Record the query."""
            seen.update(kwargs)
            return self

        # ----------------------------------
        def execute(self):
            """Return no files."""
            return {"files": []}

    class Drive:
        """Stand-in for the Drive client."""

        # ----------------------------------
        def files(self):
            """The files collection."""
            return Files()

    G.find_decks(Drive(), dt.date(2026, 9, 18))
    assert "appProperties has" in seen["q"], seen
    assert "2026-09-18" in seen["q"], seen


# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
