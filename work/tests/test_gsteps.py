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
Last updated: 2026-09-16
"""

import copy
import datetime as dt
import json
import os

import g2_add_news as NEWS
import g3_update_toc as TOC
import g4_update_bench as BENCH
import g5_update_aa_index as AA
import g6_update_youtube as YT
from gslides import bench as BK
import g7_update_layoffs as LAY
from sources import layoffs as LAYSRC
from layout import deck_layout as L
import g8_preflight as CHECK
from gslides import deck as G
from gslides import render as R
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
    assert names[0].startswith("An Anthropic resignation"), names
    assert names.index("Layoffs.fyi tracker") < \
        names.index("Artificial Analysis Intelligence Index")


# --------------------------------------------------------------
def test_toc_items_fit_one_line_without_links():
    """Links and trailing marks go; long headlines are cut."""
    assert TOC.shorten("English - https://lmarena.ai/x") == \
        "English"
    assert TOC.shorten("Tech Layoffs by year (US only):") == \
        "Tech Layoffs by year (US only)"
    long = ("Google's agentic video understanding cuts video "
            "tokens by up to 88% across every benchmark")
    short = TOC.shorten(long)
    assert R.toc_fits(short) and long.startswith(short), short
    assert short.split()[-1] not in TOC.DANGLING, short


# --------------------------------------------------------------
def test_toc_labels_rename_merge_and_leave_out():
    """A label replaces, "-" drops, equal labels list once."""
    deck = lived()
    lines = TOC.headlines(deck)
    labels = {lines[0]: "Anthropic Safety Debate",
              lines[1]: "Anthropic Safety Debate",
              lines[2]: "-"}
    names = TOC.collect(deck, labels)
    assert names[0] == "Anthropic Safety Debate", names
    assert names.count("Anthropic Safety Debate") == 1
    assert TOC.shorten(lines[2]) not in names


# --------------------------------------------------------------
def test_toc_labels_are_kept_in_the_toc_notes():
    """Only new or changed labels are appended; last wins."""
    deck = lived()
    page = deck.slide(T.PAGE_TOC)
    reqs = TOC.label_requests(deck, {"Sept 10": ""})
    body = reqs[0]["insertText"]
    assert body["objectId"] == page.notes_id
    assert body["text"].endswith("Sept 10 => -")
    page.notes = body["text"].lstrip("\n") + "\nSept 10 => X"
    assert TOC.read_labels(deck) == {"Sept 10": "X"}
    assert TOC.label_requests(deck, {"Sept 10": "X"}) == []


# --------------------------------------------------------------
def test_toc_is_a_bold_blue_bulleted_list():
    """Every item is a real list bullet and bold blue."""
    reqs = R.render_toc_columns("s-toc", "s-toc",
                                ["One", "Two", "Three"], 1.0)
    texts = [r["insertText"]["text"] for r in reqs
             if "insertText" in r]
    assert not any(line.startswith(R.BULLET.strip())
                   for t in texts for line in t.split("\n"))
    listed = [r for r in reqs if "createParagraphBullets" in r]
    assert len(listed) == 3
    styles = [r["updateTextStyle"]["style"] for r in reqs
              if "updateTextStyle" in r]
    assert styles and all(s.get("bold") for s in styles)
    assert len(texts) == 2


# --------------------------------------------------------------
def test_news_bullets_are_a_real_list():
    """Fact lines get list bullets; headline and link do not."""
    block = Block("Head", bullets=["One **fact**", "Two",
                                   "https://example.com"])
    item = L.Item(block, 12, L.Rect(0, 0, 4, 2))
    body = R.block_body(item)
    assert "●" not in body.text
    reqs = R.write_body("b1", body)
    listed = [r["createParagraphBullets"]["textRange"]
              for r in reqs if "createParagraphBullets" in r]
    starts = [len("Head\n"), len("Head\nOne fact\n")]
    assert [t["startIndex"] for t in listed] == starts


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
def hand_aa_deck(date_filled=False):
    """A talk whose Intelligence Index slide a human rebuilt."""
    sid = "g_hand"

    # --------------------------------------
    def shape(oid, kind, text="", y=2.0, w=2.0, filled=False,
              alt=""):
        """One element on the hand-made slide."""
        return G.Shape(id=oid, slide=sid, kind=kind, text=text,
                       filled=filled, rect=L.Rect(0.5, y, w, 1),
                       description=alt)

    shapes = [
        shape("g_title", G.KIND_TEXT, AA.HEADLINE + "\n",
              y=0.05),
        shape("g_box", G.KIND_TEXT,
              AA.HEADLINE + "\nhttps://x\n", y=0.7, filled=True),
        shape("g_date", G.KIND_TEXT, "Sept 10\n", y=1.5,
              filled=date_filled),
        shape("g_small", G.KIND_IMAGE, y=1.0, w=0.5),
        shape("g_mine", G.KIND_IMAGE, y=3.0, w=9.5),
        shape("g_chart", G.KIND_IMAGE, y=2.5, w=9.0,
              alt="Chart. " + AA.CHART_MARK),
    ]
    slides = [G.Slide(id="s-toc", index=0),
              G.Slide(id=sid, index=1, shapes=shapes)]
    return G.Deck(id="d", title="t", revision="r", slides=slides)


# --------------------------------------------------------------
def test_aa_finds_a_hand_made_slide_by_its_title():
    """The marked picture and the date box, nothing else.

    g_mine is the largest picture, but it is the author's
    own: no mark, so it is never chosen.
    """
    target = AA.locate(hand_aa_deck())
    assert target.page.id == "g_hand"
    assert target.picture == "g_chart"
    assert target.date.id == "g_date"
    assert set(target.allow) == {"g_chart", "g_date"}


# --------------------------------------------------------------
def test_aa_without_a_marked_picture_touches_no_picture():
    """No mark, no picture: guessing is never allowed."""
    deck = hand_aa_deck()
    deck.shape("g_chart").description = ""
    target = AA.locate(deck)
    assert target.picture == ""
    assert set(target.allow) == {"g_date"}


# --------------------------------------------------------------
def test_aa_updates_a_filled_date_box_on_a_hand_slide():
    """The author's styling fill does not block the date."""
    target = AA.locate(hand_aa_deck(date_filled=True))
    assert target.date.id == "g_date"
    assert "g_date" in target.allow


# --------------------------------------------------------------
def test_aa_script_slide_needs_no_exception():
    """The script's own slide uses its picture, allows none."""
    target = AA.locate(lived())
    assert target.picture == AA.PICTURE
    assert target.allow == ()


# --------------------------------------------------------------
def test_aa_date_is_rewritten_in_deck_style():
    """Sept 10 becomes Sept 14; new text goes in first."""
    assert W.date_label(dt.date(2026, 9, 14)) == "Sept 14"
    assert AA.DATE.match("Sept 10") and AA.DATE.match("May 3")
    assert not AA.DATE.match("Sept 10 was busy")
    box = hand_aa_deck().shape("g_date")
    reqs = W.replace_date(box, dt.date(2026, 9, 14))
    assert reqs[0]["insertText"]["insertionIndex"] == 7
    assert reqs[0]["insertText"]["text"] == "Sept 14"
    assert reqs[1]["deleteText"]["textRange"]["endIndex"] == 7
    assert W.replace_date(box, dt.date(2026, 9, 10)) == []


# --------------------------------------------------------------
def test_replace_date_keeps_words_and_padding():
    """Only the date inside "Data for Sept 02" changes."""
    box = G.Shape(id="g_d", slide="s", kind=G.KIND_TEXT,
                  text="Data for Sept 02")
    reqs = W.replace_date(box, dt.date(2026, 9, 5))
    assert reqs[0]["insertText"]["text"] == "Sept 05"
    assert reqs[0]["insertText"]["insertionIndex"] == 16
    assert reqs[1]["deleteText"]["textRange"]["startIndex"] == 9


# --------------------------------------------------------------
def hand_bench_deck():
    """A talk whose Benchmarks slide is two hand-made tables."""
    sid = "g_bench"

    # --------------------------------------
    def shape(oid, kind, x, y, text="", cells=None):
        """One element on the hand-made slide."""
        return G.Shape(id=oid, slide=sid, kind=kind, text=text,
                       rect=L.Rect(x, y, 3, 1),
                       cells=cells or [])

    head = ["Code", "Model", "Score"]
    shapes = [
        shape("g_date", G.KIND_TEXT, 4, 0.1,
              "Data for Sept 02"),
        shape("g_en", G.KIND_TEXT, 0.1, 0.6,
              "English - https://x"),
        shape("g_co", G.KIND_TEXT, 3.0, 0.6,
              "Coding - https://y"),
        shape("g_t_co", G.KIND_TABLE, 3.4, 1.0, cells=[
            head, ["■", "old-coder", "1500"],
            ["■", "same", "1400"]]),
        shape("g_t_en", G.KIND_TABLE, 0.4, 1.0, cells=[
            head, ["■", "old-english", "1490"]]),
    ]
    slides = [G.Slide(id="s-toc", index=0),
              G.Slide(id=sid, index=1, shapes=shapes)]
    return G.Deck(id="d", title="t", revision="r", slides=slides)


BOARDS = [
    {"label": "English", "cutoff": "2026-09-13", "entries": [
        {"name": "claude-fable-5", "score": 1506,
         "vendor": "anthropic", "url": "https://a/fable"}]},
    {"label": "Coding", "cutoff": "2026-09-12", "entries": [
        {"name": "gpt-6", "score": 1543, "vendor": "openai",
         "url": "https://a/gpt"},
        {"name": "same", "score": 1401, "vendor": "other"}]},
]


# --------------------------------------------------------------
def test_bench_tables_are_matched_by_their_captions():
    """English goes to the table under English, not by order."""
    page = hand_bench_deck().slide("g_bench")
    pairs = BENCH.pair_tables(page, BOARDS)
    assert [(t.id, b["label"]) for t, b in pairs] == \
        [("g_t_en", "English"), ("g_t_co", "Coding")]


# --------------------------------------------------------------
def test_bench_rows_change_only_what_differs():
    """New name: colour, text, link. Same name: score only."""
    deck = hand_bench_deck()
    table = deck.shape("g_t_co")
    reqs = BENCH.table_requests(table, BOARDS[1])
    kinds = [next(iter(r)) for r in reqs]
    assert kinds.count("updateTableCellProperties") == 1
    texts = [r["insertText"]["text"] for r in reqs
             if "insertText" in r]
    assert texts == ["gpt-6", "1543", "1401"], texts
    colour = reqs[0]["updateTableCellProperties"]
    assert colour["tableRange"]["location"] == \
        {"rowIndex": 1, "columnIndex": 0}


# --------------------------------------------------------------
def test_bench_hand_page_uses_the_site_cutoff_date():
    """The latest cutoff lands in the date box, padded."""
    deck = hand_bench_deck()
    reqs = BENCH.hand_requests(deck, BOARDS)
    dates = [r["insertText"]["text"] for r in reqs
             if r.get("insertText", {}).get("objectId")
             == "g_date"]
    assert dates == ["Sept 13"], dates
    assert set(BENCH.allowed_ids(deck)) == \
        {"g_t_en", "g_t_co", "g_date"}


# --------------------------------------------------------------
def test_the_benchmarks_page_is_found_by_its_tables():
    """Your table page wins; else the script's s-bench."""
    assert BK.board_page(hand_bench_deck()).id == "g_bench"
    assert [t.id for t in BK.board_tables(
        hand_bench_deck().slide("g_bench"))] == \
        ["g_t_en", "g_t_co"]
    assert BK.board_page(skeleton()).id == T.PAGE_BENCH


# --------------------------------------------------------------
def test_the_new_deck_copies_the_latest_earlier_deck():
    """Only a deck from before the new date counts."""
    def deck_file(fid, day, made="2026-01-01"):
        """A Drive listing entry with its date tag."""
        return {"id": fid, "createdTime": made,
                "appProperties": {G.DATE_KEY: day}}
    files = [deck_file("a", "2026-09-11"),
             deck_file("b", "2026-09-18", "2026-09-14"),
             deck_file("b2", "2026-09-18", "2026-09-15"),
             deck_file("c", "2026-09-25"),
             deck_file("d", "2026-10-02"),
             {"id": "e", "appProperties": {}},
             {"id": "f"}]
    pick = G.latest_before(files, dt.date(2026, 9, 25))
    assert pick["id"] == "b2"
    assert G.latest_before(files, dt.date(2026, 9, 11)) is None


# --------------------------------------------------------------
def test_the_copy_keeps_the_authors_pages_draws_the_rest():
    """Other slides go; kept ones return to 2, 3, 5, 7."""
    import g1_new_deck as G1
    deck = hand_bench_deck()
    reqs = G1.only_pages(deck, ["g_bench"])
    assert [W.target(r) for r in reqs] == ["s-toc"]
    assert G1.carried_pages(deck) == {"bench": "g_bench"}
    assert G1.carried_pages(lived()) == {
        "bench": T.PAGE_BENCH, "aa_index": T.PAGE_AA,
        "youtube": T.PAGE_YOUTUBE, "layoffs": T.PAGE_LAYOFFS}
    kept = {"bench": "g_bench", "aa_index": "g_aa",
            "youtube": "g_yt", "layoffs": "g_lay"}
    pages = G1.build_pages("AI News - Sept 25, 2026",
                           lambda path: "https://x/" + path,
                           skip=kept)
    ids = [sid for sid, _ in pages]
    assert len(ids) == 7
    for gone in (T.PAGE_BENCH, T.PAGE_AA, T.PAGE_YOUTUBE,
                 T.PAGE_LAYOFFS):
        assert gone not in ids
    order = G1.final_order(ids, kept)
    assert order[1:3] == ["g_bench", "g_aa"]
    assert order[4] == "g_yt" and order[6] == "g_lay"
    assert order.index("g_lay") + 1 == \
        order.index(T.PAGE_ABOUT) and len(order) == 11


FYI_TEXT = ("  Tech Layoffs by year (US only):\n"
            "128.5K in 2026 (as of Sept 10, 2026)\n"
            "124K in 2025 \n153K in 2024\n264K in 2023\n"
            "165K in 2022      https://layoffs.fyi")
TRUEUP_TEXT = ("  The Tech Layoff Tracker\n"
               "In 2026: 187,160 people laid off (737 per day)\n"
               "In 2025: 245,953 people laid off (674 per day)\n"
               "In 2024: 238,461 people laid off (653 per day)\n"
               "https://trueup.io/layoffs")


# --------------------------------------------------------------
def hand_layoffs_deck():
    """A layoffs slide with the author's boxes and pictures."""
    sid = T.PAGE_LAYOFFS

    # --------------------------------------
    def shape(oid, kind, y, h, text="", alt=""):
        """One element on the hand-made slide."""
        return G.Shape(id=oid, slide=sid, kind=kind, text=text,
                       rect=L.Rect(1, y, 5, h), description=alt)

    shapes = [
        shape("g_links", G.KIND_TEXT, 0.05, 0.4,
              "https://layoffs.fyi\nhttps://trueup.io/layoffs"),
        shape("g_tru", G.KIND_TEXT, 0.6, 1.0, TRUEUP_TEXT),
        shape("g_fyi", G.KIND_TEXT, 3.8, 1.4, FYI_TEXT),
        shape("g_mine", G.KIND_IMAGE, 3.9, 1.2),
        shape("g_pic_low", G.KIND_IMAGE, 3.5, 2.0,
              alt=LAY.FYI_MARK),
        shape("g_pic_top", G.KIND_IMAGE, 0.5, 2.8,
              alt=LAY.TRUEUP_MARK),
    ]
    slides = [G.Slide(id="s-toc", index=0),
              G.Slide(id=sid, index=1, shapes=shapes)]
    return G.Deck(id="d", title="t", revision="r", slides=slides)


# --------------------------------------------------------------
def test_layoffs_hand_slide_parts_are_found():
    """Boxes by their lines, pictures only by their mark."""
    hand = LAY.locate_hand(hand_layoffs_deck())
    assert hand.trueup_box.id == "g_tru"
    assert hand.fyi_box.id == "g_fyi"
    assert hand.trueup_picture == "g_pic_top"
    assert hand.fyi_picture == "g_pic_low"
    assert set(hand.allow()) == \
        {"g_tru", "g_fyi", "g_pic_top", "g_pic_low"}
    assert "g_mine" not in hand.allow()
    assert LAY.locate_hand(lived()) is None


# --------------------------------------------------------------
def test_your_own_picture_on_a_content_page_is_never_touched():
    """No step may delete or replace a picture a human added.

    Checked at every level: the planner's sweep, picture
    refresh, and the last-line guard on raw requests.
    """
    deck = lived()
    page = deck.slide(T.PAGE_NEWS_2)
    mine = G.Shape(id="g_my_photo", slide=page.id,
                   kind=G.KIND_IMAGE, rect=L.Rect(1, 1, 2, 2))
    page.shapes.append(mine)
    rebuilt = W.replace_owned(deck, page.id, [])
    assert "g_my_photo" not in {W.target(r) for r in rebuilt}
    assert W.refresh_picture(deck, "g_my_photo", "https://x") \
        == []
    raw = [{"deleteObject": {"objectId": "g_my_photo"}},
           {"replaceImage": {"imageObjectId": "g_my_photo",
                             "url": "https://x"}}]
    kept, dropped = W.guard(deck, raw)
    assert not kept and dropped == ["g_my_photo"]
    assert not NEWS.free_placeholder(deck, T.PAGE_NEWS_2,
                                     T.TOPIC_NEWS_2)


# --------------------------------------------------------------
def test_layoffs_fyi_numbers_keep_their_rounding():
    """One decimal stays one decimal; the date is today."""
    totals = {2026: 128873, 2025: 122606, 2024: 152922,
              2023: 265660, 2022: 165269}
    day = dt.date(2026, 9, 14)
    edits = LAY.fyi_edits(FYI_TEXT, totals, day)
    text = W.apply_edits(FYI_TEXT, edits)
    assert "128.9K in 2026 (as of Sept 14, 2026)" in text, text
    assert "123K in 2025 \n153K in 2024\n266K in 2023\n" \
        "165K in 2022" in text, text
    doubled = text.replace("K in", "KK in")
    assert W.apply_edits(doubled, LAY.fyi_edits(doubled, totals,
                                              day)) == text
    assert LAY.fyi_edits(FYI_TEXT, {}, day) == []


# --------------------------------------------------------------
def test_trueup_sentence_and_lines():
    """Totals read from the sentence land in the right lines."""
    page = ("So far in 2026, there have been 602\xa0layoffs at "
            "tech companies with 187,604 people impacted (727 "
            "people per day). In 2025, there were 783 layoffs "
            "at tech companies w/ 245,953 people impacted (674 "
            "people per day).")
    totals = LAYSRC.parse_trueup(page)
    assert totals == {2026: ("187,604", "727"),
                      2025: ("245,953", "674")}
    edits = LAY.trueup_edits(TRUEUP_TEXT, totals)
    assert [n for _, _, n in edits] == \
        ["187,604", "727", "245,953", "674"]


# --------------------------------------------------------------
def test_replace_spans_runs_backwards_and_skips_equal():
    """Later spans are written first; unchanged spans skipped."""
    text = "a 11 b 22"
    reqs = W.replace_spans("box", text,
                           [(2, 4, "111"), (7, 9, "22")])
    assert len(reqs) == 2
    assert reqs[0]["insertText"]["insertionIndex"] == 4
    reqs = W.replace_spans("box", text,
                           [(2, 4, "1"), (7, 9, "3")])
    assert reqs[0]["insertText"]["insertionIndex"] == 9


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
def youtube_deck():
    """The lived-in deck and its channel slide."""
    deck = lived()
    return deck, deck.slide(T.PAGE_YOUTUBE)


# --------------------------------------------------------------
def test_a_filled_counts_box_still_gets_the_new_counts():
    """The counts line is the one thing a fill cannot stop."""
    deck, _ = youtube_deck()
    box = deck.shape("t-youtube-b")
    box.filled = True
    assert W.replace_line(deck, box.id, "subscribers",
                          "9.99K subscribers, 400 videos") == []
    reqs = W.replace_line(deck, box.id, "subscribers",
                          "9.99K subscribers, 400 videos",
                          (box.id,))
    assert reqs and "deleteText" in reqs[0]


# --------------------------------------------------------------
def test_the_counts_box_you_made_yourself_is_found():
    """A hand-made promo box is the one step 6 writes to."""
    deck, page = youtube_deck()
    page.shapes = [s for s in page.shapes
                   if s.id != "t-youtube-b"]
    mine = G.Shape(id="g_promo", slide=page.id,
                   kind=G.KIND_TEXT, filled=True,
                   text="Weekly videos every Friday\n"
                        "7.51K subscribers, 337 videos",
                   rect=L.Rect(1, 2, 4, 1))
    page.shapes.append(mine)
    found = YT.locate(deck)
    assert found.box.id == "g_promo"
    assert found.allow == ("g_promo",)
    reqs = YT.counts_reqs(deck, found,
                          "7.52K subscribers, 339 videos")
    assert reqs and W.target(reqs[0]) == "g_promo"


# --------------------------------------------------------------
def test_a_deleted_promo_box_is_drawn_again():
    """Step 6 heals the slide from config/skeleton.json."""
    deck, page = youtube_deck()
    page.shapes = [s for s in page.shapes
                   if s.id != "t-youtube-b"]
    found = YT.locate(deck)
    assert found.box is None and found.allow == ()
    line = "7.52K subscribers, 339 videos"
    reqs = YT.counts_reqs(deck, found, line)
    made = [W.target(r) for r in reqs if "createShape" in r]
    assert made == ["t-youtube-b"], made
    written = "".join(r["insertText"]["text"] for r in reqs
                      if "insertText" in r)
    assert line in written
    assert "https://www.youtube.com/@lev-selector" in written


# --------------------------------------------------------------
def text_box(text, h, w=2.0, styles=None):
    """A hand-made text box with a given height."""
    return G.Shape(id="g_box", slide="s", kind=G.KIND_TEXT,
                   text=text, rect=L.Rect(1, 1, w, h),
                   size_emu=(w * 914400, 914400),
                   transform={"scaleX": 1, "scaleY": h,
                              "unit": "EMU"},
                   styles=styles or [(12, False)])


# --------------------------------------------------------------
def heights(reqs):
    """New heights, in inches, that requests set."""
    return [r["updatePageElementTransform"]["transform"]
            ["scaleY"] for r in reqs
            if "updatePageElementTransform" in r]


# --------------------------------------------------------------
def test_an_edit_that_adds_a_line_grows_the_box_by_it():
    """One more 12pt line adds one line height, padding kept."""
    box = text_box("one\ntwo", 0.50)
    line = L.points_to_inches(12 * L.LINE_RATIO)
    grown = heights(W.edit_spans(box, [(7, 7, "\nthree")]))
    assert len(grown) == 1 and abs(grown[0] - (0.50 + line)) \
        < 0.001, grown
    assert heights(W.edit_spans(box, [(0, 3, "ONE")])) == []


# --------------------------------------------------------------
def test_an_edit_that_unwraps_a_line_shrinks_the_box():
    """A long line cut short takes its extra lines away."""
    long = "word " * 30
    box = text_box(long.strip(), 1.80)
    shrunk = heights(W.edit_spans(box, [(0, len(long) - 1,
                                         "short")]))
    assert len(shrunk) == 1 and 0.2 < shrunk[0] < 0.5, shrunk
    small = text_box(long.strip(), 0.50)
    floor = heights(W.edit_spans(small, [(0, len(long) - 1,
                                          "short")]))
    assert floor and floor[0] > 0.15, floor


# --------------------------------------------------------------
def test_a_rewritten_box_is_exactly_as_tall_as_its_text():
    """refresh_text sizes the box from the body it writes."""
    body = R.plain_body("first line\nsecond line", 14)
    box = text_box("old", 3.0, w=4.0)
    want = R.box_height(R.body_paragraphs(body), 4.0)
    assert heights(W.fit_rewrite(box, body)) == \
        [W.S.emu(want) / 914400]
    assert [p[:3] for p in R.body_paragraphs(body)] == \
        [("first line", 14, False), ("second line", 14, False)]


# --------------------------------------------------------------
def test_paragraph_styles_are_read_per_paragraph():
    """Size and bold come from most of each paragraph."""
    shape_json = {"text": {"textElements": [
        {"textRun": {"content": "Title\n", "style": {
            "fontSize": {"magnitude": 14}, "bold": True}}},
        {"textRun": {"content": "In 2026: ", "style": {
            "fontSize": {"magnitude": 12}}}},
        {"textRun": {"content": "1", "style": {
            "fontSize": {"magnitude": 12}, "bold": True}}},
        {"textRun": {"content": " laid off\n", "style": {
            "fontSize": {"magnitude": 12}}}},
    ]}}
    assert G.paragraph_styles(shape_json) == \
        [(14, True), (12, False)]


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
def test_deck_text_covers_the_talk_and_skips_parked():
    """Titles and box text of the talk; nothing parked."""
    import g9_deck_text as TXT
    deck = lived()
    text = TXT.deck_text(deck)
    heads = [l for l in text.splitlines()
             if l.startswith("## Slide")]
    assert len(heads) == len(deck.main()), heads
    assert "## Slide 2: " in text
    assert len(deck.slides) > len(deck.main())
    first = deck.main()[0]
    assert TXT.slide_text(first, 1).startswith("## Slide 1")
    assert TXT.split_out(["2026-09-25", "--out", "x.md"]) == \
        ("x.md", ["2026-09-25"])

# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
