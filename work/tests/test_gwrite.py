#!/usr/bin/env python3
"""
Tests for gslides/write.py: nothing may reach a protected shape.

Uses the same saved skeleton deck as tests/test_gdeck.py, and a
small fake Slides client for the revision-clash retry.

Usage:
    python3 -m tests.test_gwrite

Created: 2026-09-14
Last updated: 2026-09-14
"""

from layout import deck_layout as L
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import api as S
from gslides import ids as T
from tests.test_gdeck import add_human_box, fill, move_slide, raw
from tests.runner import run

RECT = L.Rect(1, 1, 2, 1)


# --------------------------------------------------------------
def body(text="new words"):
    """A one-line body to write."""
    return R.plain_body(text, 12)


# --------------------------------------------------------------
def kinds(requests):
    """The request type of each request, in order."""
    return [next(iter(r)) for r in requests]


# --------------------------------------------------------------
def test_target_finds_the_object_of_each_request_type():
    """Every request type used here names its target."""
    assert W.target(S.textbox("t-x-b", "s-y", RECT)[0]) == \
        "t-x-b"
    assert W.target({"replaceImage": {
        "imageObjectId": "t-x-p"}}) == "t-x-p"
    assert W.target(S.delete_object("t-x-b")[0]) == "t-x-b"


# --------------------------------------------------------------
def test_refresh_text_replaces_words_in_an_editable_box():
    """Delete the old text, then write and style the new."""
    deck = G.parse_deck(raw())
    reqs = W.refresh_text(deck, "t-aa-index-b", body())
    assert kinds(reqs)[:2] == ["deleteText", "insertText"]


# --------------------------------------------------------------
def test_refresh_text_leaves_a_frozen_box_alone():
    """A filled box gets no requests at all."""
    deck = G.parse_deck(fill(raw(), "t-aa-index-b"))
    assert W.refresh_text(deck, "t-aa-index-b", body()) == []


# --------------------------------------------------------------
def test_refresh_text_never_touches_a_human_box():
    """Even when asked by id."""
    deck = G.parse_deck(add_human_box(raw(), T.PAGE_NEWS_1))
    assert W.refresh_text(deck, "SLIDES_API123456_0",
                          body()) == []


# --------------------------------------------------------------
def test_refresh_picture_follows_the_box():
    """A frozen box keeps its old picture."""
    deck = G.parse_deck(raw())
    assert W.refresh_picture(deck, "t-aa-index-p", "https://x")
    deck = G.parse_deck(fill(raw(), "t-aa-index-b"))
    assert W.refresh_picture(deck, "t-aa-index-p",
                             "https://x") == []


# --------------------------------------------------------------
def test_refresh_picture_keeps_the_alt_text_mark():
    """replaceImage wipes alt text; the mark is written back."""
    deck = G.parse_deck(raw())
    pic = deck.shape("t-aa-index-p")
    assert [list(r)[0] for r in W.refresh_picture(
        deck, pic.id, "https://x")] == ["replaceImage"]
    pic.description = "auto: aa-index-chart"
    reqs = W.refresh_picture(deck, pic.id, "https://x")
    assert list(reqs[0]) == ["replaceImage"]
    alt = reqs[1]["updatePageElementAltText"]
    assert alt["objectId"] == pic.id
    assert alt["description"] == "auto: aa-index-chart"


# --------------------------------------------------------------
def test_replace_owned_rebuilds_editable_shapes():
    """Existing editable shapes are deleted, then recreated."""
    deck = G.parse_deck(raw())
    new = R.make_box("s-bench-k0", T.PAGE_BENCH, RECT, True)
    reqs = W.replace_owned(deck, T.PAGE_BENCH, new)
    deleted = [W.target(r) for r in reqs if "deleteObject" in r]
    assert "s-bench-k0" in deleted
    assert "createShape" in kinds(reqs)


# --------------------------------------------------------------
def test_replace_owned_spares_frozen_shapes():
    """A filled column keeps its old self."""
    deck = G.parse_deck(fill(raw(), "s-bench-k1"))
    new = R.make_box("s-bench-k0", T.PAGE_BENCH, RECT, True)
    new += R.make_box("s-bench-k1", T.PAGE_BENCH, RECT, True)
    reqs = W.replace_owned(deck, T.PAGE_BENCH, new)
    touched = {W.target(r) for r in reqs}
    assert "s-bench-k1" not in touched, touched
    assert "s-bench-k0" in touched


# --------------------------------------------------------------
def test_replace_owned_sweeps_stale_script_shapes_only():
    """Leftover script shapes go; human shapes stay."""
    deck = G.parse_deck(add_human_box(raw(), T.PAGE_BENCH))
    new = R.make_box("s-bench-k0", T.PAGE_BENCH, RECT, True)
    reqs = W.replace_owned(deck, T.PAGE_BENCH, new)
    deleted = {W.target(r) for r in reqs if "deleteObject" in r}
    assert "s-bench-k1" in deleted, deleted
    assert "SLIDES_API123456_0" not in deleted, deleted


# --------------------------------------------------------------
def test_guard_blocks_writes_to_human_and_frozen_shapes():
    """The last check drops what the planner let through."""
    data = add_human_box(fill(raw(), "t-trueup-b"),
                         T.PAGE_LAYOFFS)
    deck = G.parse_deck(data)
    reqs = (S.insert_text("SLIDES_API123456_0", "x")
            + S.insert_text("t-trueup-b", "x")
            + S.insert_text("t-layoffs-fyi-b", "x"))
    kept, dropped = W.guard(deck, reqs)
    assert [W.target(r) for r in kept] == ["t-layoffs-fyi-b"]
    assert dropped == ["SLIDES_API123456_0", "t-trueup-b"]


# --------------------------------------------------------------
def test_guard_lets_through_only_explicitly_allowed_shapes():
    """allow opens one human shape, not the others."""
    data = add_human_box(raw(), T.PAGE_LAYOFFS)
    deck = G.parse_deck(data)
    reqs = S.insert_text("SLIDES_API123456_0", "x")
    kept, dropped = W.guard(deck, reqs,
                            allow=("SLIDES_API123456_0",))
    assert len(kept) == 1 and not dropped
    kept, dropped = W.guard(deck, reqs, allow=("other",))
    assert not kept and dropped == ["SLIDES_API123456_0"]


# --------------------------------------------------------------
def test_guard_allows_brand_new_shapes():
    """Creating and filling a new id is always fine."""
    deck = G.parse_deck(raw())
    reqs = R.make_box("t-new-topic-1234-b", T.PAGE_NEWS_2,
                      RECT, True)
    reqs += S.insert_text("t-new-topic-1234-b", "hello")
    kept, dropped = W.guard(deck, reqs)
    assert len(kept) == len(reqs) and not dropped


# --------------------------------------------------------------
def test_guard_protects_human_and_parked_slides():
    """No script may delete a slide it does not own."""
    data = move_slide(raw(), T.PAGE_NEWS_2, 10)
    data["slides"][0]["objectId"] = "p"
    deck = G.parse_deck(data)
    reqs = S.delete_object("p") + S.delete_object(
        T.PAGE_NEWS_2) + S.delete_object(T.PAGE_NEWS_1)
    kept, dropped = W.guard(deck, reqs)
    assert [W.target(r) for r in kept] == [T.PAGE_NEWS_1]
    assert set(dropped) == {"p", T.PAGE_NEWS_2}


class FakeSlides:
    """Refuses the first write as a revision clash."""

    # --------------------------------------
    def __init__(self, clashes=1):
        """Count reads and writes."""
        self.clashes = clashes
        self.reads = 0
        self.writes = []
        self.pending = None

    # --------------------------------------
    def presentations(self):
        """The presentations collection."""
        return self

    # --------------------------------------
    def get(self, presentationId):
        """Queue a read of the saved deck."""
        self.pending = ("get", None)
        return self

    # --------------------------------------
    def batchUpdate(self, presentationId, body):
        """Queue a write."""
        self.pending = ("write", body)
        return self

    # --------------------------------------
    def execute(self):
        """Run whatever was queued."""
        kind, body = self.pending
        if kind == "get":
            self.reads += 1
            return raw()
        if self.clashes:
            self.clashes -= 1
            raise W.Conflict("clash")
        self.writes.append(body)
        return {"writeControl": {"requiredRevisionId": "r2"}}


# --------------------------------------------------------------
def test_run_plan_rereads_and_replans_after_a_clash():
    """A clash costs one fresh read, not a lost edit."""
    fake = FakeSlides(clashes=1)
    plans = []

    # --------------------------------------
    def plan(deck):
        """Rewrite the AA box."""
        plans.append(deck.revision)
        return [("aa", W.refresh_text(deck, "t-aa-index-b",
                                      body()))]

    sent = W.run_plan(fake, "deck", plan, "test")
    assert fake.reads == 2 and len(plans) == 2
    assert sent > 0 and len(fake.writes) == 1
    control = fake.writes[0]["writeControl"]
    assert control["requiredRevisionId"] == RAW_REVISION


# --------------------------------------------------------------
def test_run_plan_gives_up_when_the_deck_keeps_changing():
    """Two clashes in a row stop the script cleanly."""
    fake = FakeSlides(clashes=5)
    sent = W.run_plan(
        fake, "deck",
        lambda d: [("aa", S.insert_text("t-aa-index-b", "x"))],
        "test")
    assert sent == -1 and fake.writes == []


RAW_REVISION = raw()["revisionId"]


# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
