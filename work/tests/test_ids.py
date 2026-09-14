#!/usr/bin/env python3
"""
Tests for gslides/ids.py, the stable object id scheme.

What these protect: an id must stay the same when a topic
moves, must differ when the topic differs, and must always
be something Slides will accept.

Usage:
    python3 -m tests.test_ids

Created: 2026-09-14
Last updated: 2026-09-14
"""

from gslides import ids as T
from layout.deck_layout import Item, Page, Rect
from layout.deck_parser import Block
from tests.runner import run

RECT = Rect(0, 0, 1, 1)


# --------------------------------------------------------------
def block(headline, bullets=None, image=None):
    """Build one news block for a test."""
    return Block(headline=headline,
                 bullets=bullets or ["a bullet"],
                 image=image)


# --------------------------------------------------------------
def page(kind, blocks=(), title=""):
    """Build one paginated page for a test."""
    items = [Item(block=b, size=12, text=RECT) for b in blocks]
    return Page(title=title, kind=kind, items=items)


# --------------------------------------------------------------
def test_slug_keeps_words_and_drops_punctuation():
    """A headline becomes lowercase words joined by dashes."""
    got = T.slug("Meta ships Muse agent!")
    assert got == "meta-ships-muse-agent", got


# --------------------------------------------------------------
def test_slug_cuts_on_a_word_boundary():
    """A long headline is cut, but not mid-word."""
    long = "the quick brown fox jumps over the lazy dog today"
    got = T.slug(long)
    assert len(got) <= T.SLUG_MAX, got
    assert not got.endswith("-"), got
    assert long.replace(" ", "-").startswith(got), got


# --------------------------------------------------------------
def test_slug_survives_markup():
    """Bold and highlight marks do not reach the id."""
    plain = T.slug("Meta ships Muse agent")
    marked = T.slug("**Meta** ships ==Muse== agent")
    assert plain == marked, (plain, marked)


# --------------------------------------------------------------
def test_key_is_the_same_for_the_same_topic():
    """The id does not depend on where the topic sits."""
    one = T.block_key(block("Mistral raises a big round"))
    two = T.block_key(block("Mistral raises a big round"))
    assert one == two, (one, two)


# --------------------------------------------------------------
def test_key_differs_for_different_topics():
    """Two topics never share an id."""
    one = T.block_key(block("Mistral raises a big round"))
    two = T.block_key(block("Suno licenses its music"))
    assert one != two, one


# --------------------------------------------------------------
def test_long_headlines_sharing_a_prefix_stay_apart():
    """The hash separates topics that slug the same."""
    head = "the quick brown fox jumps over the lazy dog "
    one = T.block_key(block(head + "on Monday"))
    two = T.block_key(block(head + "on Tuesday"))
    assert one != two, one
    assert T.valid(one) and T.valid(two)


# --------------------------------------------------------------
def test_headless_block_falls_back_to_its_first_bullet():
    """A bare ### block still gets a stable id."""
    one = T.block_key(block("", ["Gemini cuts video tokens"]))
    two = T.block_key(block("", ["Gemini cuts video tokens"]))
    assert one == two, (one, two)
    assert T.valid(one), one


# --------------------------------------------------------------
def test_every_id_is_acceptable_to_slides():
    """Length and character rules hold for odd input."""
    for headline in ["", "!!!", "a", "x" * 300,
                     "Пример текста", "3 things"]:
        key = T.block_key(block(headline))
        assert T.valid(key), (headline, key)


# --------------------------------------------------------------
def test_fixed_pages_get_fixed_names():
    """The contents and benchmarks pages are named."""
    assert T.page_key(page("toc")) == "s-toc"
    assert T.page_key(page("benchmarks")) == "s-bench"


# --------------------------------------------------------------
def test_duplicate_topics_are_made_unique():
    """Two identical headlines still get separate ids."""
    same = block("The very same headline")
    pages = [page("content", [same, same])]
    ids = T.assign_ids(pages)
    every = ids.every()
    assert len(set(every)) == len(every), every
    assert all(T.valid(k) for k in every), every


# --------------------------------------------------------------
def test_ids_survive_the_deck_being_reordered():
    """Moving a page does not rename its topic."""
    one = block("Mistral raises a big round")
    two = block("Suno licenses its music")
    before = T.assign_ids([page("content", [one]),
                           page("content", [two])])
    after = T.assign_ids([page("content", [two]),
                          page("content", [one])])
    assert before.block(0, 0) == after.block(1, 0)
    assert before.block(1, 0) == after.block(0, 0)


# --------------------------------------------------------------
def test_slide_and_shape_ids_never_collide():
    """A page id is not also a shape id on that page."""
    pages = [page("toc"), page("content", [block("A topic")])]
    ids = T.assign_ids(pages)
    names = []
    for index, one in enumerate(pages):
        slide = ids.page(index)
        names += [slide, T.shape_id(slide, "title")]
        for spot in range(len(one.items)):
            key = ids.block(index, spot)
            names += [T.shape_id(key, "b"),
                      T.shape_id(key, "p")]
    assert len(set(names)) == len(names), sorted(names)
    assert all(T.valid(n) for n in names), sorted(names)


# --------------------------------------------------------------
def test_a_copied_box_cannot_reuse_the_key():
    """Shape ids are unique, so a copy needs a new one."""
    key = T.block_key(block("Meta ships Muse agent"))
    body = T.shape_id(key, "b")
    art = T.shape_id(key, "p")
    assert body != art, body
    assert body != key, body


# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
