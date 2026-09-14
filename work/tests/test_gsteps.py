#!/usr/bin/env python3
"""
Tests for the decisions the update steps make.

Uses fixtures/lived_in_deck.json: a real deck after g1 built
it, g2 to g7 ran on it, and a human edited it. In that deck
the human has:

    moved the layoffs slide up to third place
    put a box of their own on s-news-1
    filled the TrueUp box, one benchmark column and a TOC
    column
    copied the Mistral topic onto the parked slide, then
    removed it from the talk
    deleted the slide holding the Google and music topics

Usage:
    python3 -m tests.test_gsteps

Created: 2026-09-14
Last updated: 2026-09-14
"""

import copy
import datetime as dt
import json
import os

import g2_add_news as NEWS
import g3_update_toc as TOC
import g8_preflight as CHECK
from gslides import deck as G
from gslides import write as W
from gslides import ids as T
from layout.deck_parser import Block
from tests.test_gdeck import raw as skeleton_raw
from tests.runner import run

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURE = os.path.join(HERE, "fixtures", "lived_in_deck.json")

with open(FIXTURE, encoding="utf-8") as _handle:
    LIVED = json.load(_handle)


# --------------------------------------------------------------
def lived():
    """A fresh parse of the lived-in deck."""
    return G.parse_deck(copy.deepcopy(LIVED))


# --------------------------------------------------------------
def skeleton():
    """A fresh parse of the untouched skeleton deck."""
    return G.parse_deck(skeleton_raw())


# --------------------------------------------------------------
def topic(headline):
    """A candidate news topic."""
    return Block(headline=headline, bullets=["a fact"])


# --------------------------------------------------------------
def test_ledger_is_read_from_the_separator_notes():
    """Every topic g2 added is recorded with its key."""
    entries = NEWS.read_ledger(lived())
    keys = [k for k, _ in entries]
    assert len(entries) == 7, entries
    assert all(k.startswith("t-") for k in keys), keys


# --------------------------------------------------------------
def test_a_topic_already_in_the_talk_is_skipped():
    """Same headline as a box in the deck."""
    deck = lived()
    reason = NEWS.why_known(
        topic("Claude formalized Fermat's Last Theorem in 11 "
              "days"), deck, NEWS.read_ledger(deck), [])
    assert reason, reason


# --------------------------------------------------------------
def test_a_reworded_parked_topic_is_skipped():
    """The copy past the separator still blocks it."""
    deck = G.parse_deck(copy.deepcopy(LIVED))
    ledger = [e for e in NEWS.read_ledger(deck)
              if "mistral" not in e[0]]
    reason = NEWS.why_known(
        topic("Mistral raises largest European tech round"),
        deck, ledger, [])
    assert reason and reason.startswith("parked"), reason


# --------------------------------------------------------------
def test_a_reworded_deleted_topic_is_skipped_by_the_ledger():
    """The deleted slide's topics are gone but remembered."""
    deck = lived()
    reason = NEWS.why_known(
        topic("Music industry starts licensing instead of "
              "fighting"), deck, NEWS.read_ledger(deck), [])
    assert reason and "added before" in reason, reason


# --------------------------------------------------------------
def test_a_new_topic_passes_and_duplicates_in_a_batch_do_not():
    """Only the first of two same-topic candidates is new."""
    deck = lived()
    first = topic("Robots take over retail stores")
    second = topic("Robots take over retail stores fast")
    candidates = [(first, None), (second, None)]
    fresh, skipped = NEWS.select(candidates, deck)
    assert len(fresh) == 1 and len(skipped) == 1
    assert "repeats" in skipped[0][1]


# --------------------------------------------------------------
def test_the_slide_title_ai_news_never_blocks_a_topic():
    """Titles are not topics."""
    deck = skeleton()
    fresh, _ = NEWS.select([(topic("Meta launches AI News "
                                   "app"), None)], deck)
    assert len(fresh) == 1


# --------------------------------------------------------------
def test_placeholder_slides_are_free_until_someone_uses_them():
    """A human box on s-news-1 takes it out of play."""
    assert NEWS.free_placeholder(skeleton(), T.PAGE_NEWS_1,
                                 T.TOPIC_NEWS_1)
    deck = lived()
    assert not NEWS.free_placeholder(deck, T.PAGE_NEWS_1,
                                     T.TOPIC_NEWS_1)
    assert not NEWS.free_placeholder(deck, T.PAGE_NEWS_2,
                                     T.TOPIC_NEWS_2)


# --------------------------------------------------------------
def test_new_slides_go_before_layoffs_wherever_it_moved():
    """The anchor follows the slide, not a fixed position."""
    assert NEWS.anchor_index(skeleton()) == 6
    deck = lived()
    assert NEWS.anchor_index(deck) == \
        deck.slide(T.PAGE_LAYOFFS).index


# --------------------------------------------------------------
def test_ledger_lines_are_appended_not_rewritten():
    """New entries go after the existing notes text."""
    deck = lived()
    sep = deck.slide(T.PAGE_SEPARATOR)
    reqs = NEWS.ledger_requests(deck, [topic("New one")],
                                ["t-new-one-abcd"])
    body = reqs[0]["insertText"]
    assert body["objectId"] == sep.notes_id
    assert body["insertionIndex"] == len(sep.notes)
    assert body["text"] == "\nt-new-one-abcd | New one"


# --------------------------------------------------------------
def test_toc_follows_the_current_slide_order():
    """Layoffs moved up, so its topics come first."""
    names = TOC.collect(lived())
    assert names[0] == "An Anthropic resignation turns into " \
        "an extinction debate", names
    assert names.index("Layoffs.fyi tracker") < \
        names.index("Artificial Analysis Intelligence Index")


# --------------------------------------------------------------
def test_toc_includes_human_topics_but_not_parked_ones():
    """Yours count; the parked copy does not."""
    names = TOC.collect(lived())
    assert "My own topic: robots in retail" in names
    assert not any(n.startswith("Mistral") for n in names)


# --------------------------------------------------------------
def test_toc_leaves_a_frozen_contents_alone():
    """One filled column freezes the whole TOC."""
    deck = lived()
    assert TOC.toc_requests(deck, ["anything"]) == []


# --------------------------------------------------------------
def test_replace_line_changes_one_line_only():
    """The counts line is swapped; nothing else is touched."""
    deck = lived()
    box = deck.shape("t-youtube-b")
    reqs = W.replace_line(deck, "t-youtube-b", "subscribers",
                          "9.99K subscribers, 400 videos")
    start = box.text.index("7.51K")
    assert reqs[0]["deleteText"]["textRange"]["startIndex"] \
        == start
    assert reqs[1]["insertText"]["insertionIndex"] == start
    assert W.replace_line(deck, "t-youtube-b", "subscribers",
                          "7.51K subscribers, 337 videos") == []


# --------------------------------------------------------------
def test_a_box_may_grow_only_into_free_space():
    """Room ends where the next shape below begins."""
    deck = lived()
    fyi = deck.shape("t-layoffs-fyi-b")
    trueup = deck.shape("t-trueup-b")
    room = deck.room_below(fyi)
    assert fyi.rect.y + room <= trueup.rect.y, (room, fyi.rect)
    assert room >= fyi.rect.h


# --------------------------------------------------------------
def test_preflight_does_not_ask_to_review_decoration():
    """Epigraph, legend and cutoff note are not content."""
    lines = CHECK.unreviewed(lived())
    text = " ".join(lines)
    assert "Votes counted" not in text
    assert "Agents got" not in text
    assert any("Meta ships" in line for line in lines)


# --------------------------------------------------------------
def test_preflight_reads_the_benchmark_cutoff():
    """Over a week old is reported, a few days is not."""
    deck = lived()
    assert CHECK.old_benchmarks(deck, dt.date(2026, 9, 18)) \
        == []
    assert CHECK.old_benchmarks(deck, dt.date(2026, 10, 2))


# --------------------------------------------------------------
def test_skeleton_json_builds_the_eleven_slides_in_order():
    """g1 needs no Markdown: config/skeleton.json is enough."""
    import g1_new_deck as G1
    pages = G1.build_pages("AI News - Sept 18, 2026",
                           lambda path: "https://x/" + path)
    assert [sid for sid, _ in pages] == [
        T.PAGE_TOC, T.PAGE_BENCH, T.PAGE_AA, T.PAGE_NEWS_1,
        T.PAGE_YOUTUBE, T.PAGE_NEWS_2, T.PAGE_LAYOFFS,
        T.PAGE_ABOUT, T.PAGE_THANKS, T.PAGE_SEPARATOR,
        T.PAGE_PARKED]
    made = [W.target(r) for _, reqs in pages for r in reqs
            if "createShape" in r or "createImage" in r]
    for key in (T.TOPIC_AA, T.TOPIC_YOUTUBE, T.TOPIC_TRUEUP,
                T.TOPIC_ABOUT, T.TOPIC_THANKS):
        assert T.shape_id(key, T.BOX) in made, key
    assert T.shape_id(T.TOPIC_ABOUT, T.PICTURE) in made


# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
