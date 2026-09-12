#!/usr/bin/env python3
"""
Tests for the automatic layout and the text measurement.

These are the ones that matter most. Nobody positions
anything by hand, so a mistake in the packing arithmetic
would push text off a slide silently and nobody would
notice until the seminar. They assert directly that no
rectangle leaves the content band or the slide.

The measurement tests pin text_metrics against line breaks
read off real rendered slides, and the "agrees" tests pin
constants that are deliberately defined in two places.

No pytest needed. Each test is a function whose name starts
with "test_".

Usage:
    python3 test_layout.py

Created: 2026-09-12
Last updated: 2026-09-12
"""

import deck_layout as L
from deck_parser import KIND_BENCHMARKS, KIND_TOC, parse_deck
from test_runner import run

SIMPLE = """
# AI News - Sept 18, 2026

## slide: toc

## slide: benchmarks

## slide: AI News

### First story
![](images/one.jpg)
<!-- src: https://example.com/one.png -->
- a plain bullet
- a **bold** and ==marked== bullet
- https://example.com/article

### Second story
![](images/two.jpg)
<!-- shot: https://example.com/page -->
- only one bullet
"""


# --------------------------------------------------------------
def test_image_box_matches_the_layout():
    """Pictures are sized to the box the layout gives."""
    import s2_clean_images as S
    settings = S.Settings(ppi=150)
    width, height = settings.box()
    assert width == int(L.IMG_W * 150)
    assert height == int(L.BAND_H * 150)
    assert settings.geometry() == f"{width}x{height}>"

# --------------------------------------------------------------
def test_settings_stamp_tracks_its_settings():
    """Changing ppi or quality changes the marker."""
    import s2_clean_images as S
    base = S.Settings(ppi=150, quality=85).stamp()
    assert S.Settings(ppi=96, quality=85).stamp() != base
    assert S.Settings(ppi=150, quality=80).stamp() != base
    assert S.Settings(ppi=150, quality=85).stamp() == base

# --------------------------------------------------------------
def test_resize_only_shrinks():
    """The geometry never enlarges a small picture."""
    import s2_clean_images as S
    assert S.Settings().geometry().endswith(">")

# --------------------------------------------------------------
def test_epigraph_box_sits_top_right():
    """The epigraph never collides with the title."""
    import s3_make_pptx as s3
    assert s3.EPIGRAPH_X > L.SLIDE_W / 2
    right = s3.EPIGRAPH_X + s3.epigraph_width()
    assert abs(right - (L.SLIDE_W - L.MARGIN)) < 0.001
    assert s3.epigraph_height("short one") <= L.BAND_TOP + 0.5

# --------------------------------------------------------------
def profile_item(flag_on_section=True):
    """Lay out a one-block profile page and return it."""
    flag = "<!-- profile -->\n"
    deck = parse_deck(
        "# T\n## slide: About\n"
        + (flag if flag_on_section else "")
        + "### Lev Selector, Ph.D.\n"
        + ("" if flag_on_section else flag)
        + "![](images/me.jpg)\n- 40+ years of this\n"
    )
    return L.paginate(deck.sections)[0].items[0]


# --------------------------------------------------------------
def test_profile_puts_the_picture_on_the_left():
    """The portrait leads, the text sits beside it."""
    for on_section in (True, False):
        item = profile_item(on_section)
        assert item.image.x == L.MARGIN
        assert item.image.w == L.PROFILE_IMG_W
        assert item.text.x > item.image.x + item.image.w


# --------------------------------------------------------------
def test_profile_picture_is_larger_and_top_aligned():
    """Bigger than a news thumbnail, level with the text."""
    item = profile_item()
    assert item.image.w == item.image.h
    assert item.image.y == item.text.y == L.BAND_TOP


# --------------------------------------------------------------
def test_profile_name_is_bigger_than_a_headline():
    """The speaker's name reads as a name, not a headline."""
    item = profile_item()
    name = L.headline_size(item.block, item.size)
    assert name > item.size + L.HEADLINE_EXTRA


# --------------------------------------------------------------
def test_news_blocks_keep_the_picture_on_the_right():
    """The profile shape does not leak into news slides."""
    block = make_block("A story", ["one line"], "a.jpg")
    item = L.paginate([make_section([block])])[0].items[0]
    assert item.text.x == L.MARGIN
    assert item.image.x > item.text.x


# --------------------------------------------------------------
def test_month_names_agree():
    """The benchmark date matches the deck title's months."""
    import importlib.util
    import os
    import bench_page
    path = os.path.join(
        "..", ".claude", "skills", "slides-update",
        "tools", "deck_init.py"
    )
    if not os.path.isfile(path):
        return
    spec = importlib.util.spec_from_file_location(
        "deck_init", path
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    assert bench_page.MONTHS == module.MONTHS


# --------------------------------------------------------------
def test_cutoff_line_reads_well():
    """One date when the boards agree, both when not."""
    import bench_page as B
    same = [
        {"label": "English", "cutoff": "2026-09-11"},
        {"label": "Coding", "cutoff": "2026-09-11"},
    ]
    assert B.cutoff_note(same) == (
        "Votes counted through Sept 11, 2026"
    )
    differ = [
        {"label": "English", "cutoff": "2026-09-11"},
        {"label": "Coding", "cutoff": "2026-09-04"},
    ]
    note = B.cutoff_note(differ)
    assert "English: Sept 11, 2026" in note
    assert "Coding: Sept 4, 2026" in note
    assert B.cutoff_note([{"label": "X"}]) == ""


# --------------------------------------------------------------
def test_box_inset_agrees():
    """The layout's inset matches the real box margins."""
    import pptx_text as T
    from pptx.util import Emu, Inches
    both = T.LEFT_MARGIN_EMU * 2
    assert abs(both - Inches(L.TEXT_INSET)) < 200

# --------------------------------------------------------------
def test_bullet_inset_agrees():
    """The layout's bullet indent matches the renderer."""
    import pptx_text as T
    assert abs(L.BULLET_INSET - T.BULLET_INSET) < 1e-9

# --------------------------------------------------------------
def test_measured_wrapping_matches_slides():
    """Line breaks reproduce what the renderer drew."""
    import text_metrics as M
    cases = [
        (3.08, 14, "Artificial Analysis Intelligence Index", 1),
        (3.08, 14, "Meta ships Muse agent with its own cloud "
                   "computer and browser", 2),
        (3.08, 14, "Layoffs.fyi tracker", 1),
        (6.00, 12, "Handles email, calendars, Instagram and "
                   "ordinary websites. Ties into Gmail, "
                   "Spotify, Ticketmaster and OpenTable.", 2),
    ]
    for width, size, text, expect in cases:
        assert M.wrap_lines(text, width, size) == expect, text

# --------------------------------------------------------------
def test_font_ladder_grows_but_packs_at_base():
    """Roomy slides get bigger text, packing is unchanged."""
    assert max(L.FONT_LADDER) > L.BASE_SIZE
    assert L.BASE_SIZE in L.FONT_LADDER
    short = make_block("Short one", ["a brief line"])
    page = L.paginate([make_section([short])])[0]
    assert page.items[0].size == max(L.FONT_LADDER)

# --------------------------------------------------------------
def test_link_size_follows_the_body():
    """A link never dwarfs or vanishes under its text."""
    assert L.link_size(12) == 9
    assert L.link_size(16) == 12
    assert L.link_size(8) == 9

# --------------------------------------------------------------
def make_block(headline, bullets, image=None):
    """Build a Block without going through the parser."""
    from deck_parser import Block
    return Block(
        headline=headline, bullets=list(bullets),
        image=image
    )

# --------------------------------------------------------------
def make_section(blocks, title="News"):
    """Build a content Section from blocks."""
    from deck_parser import Section
    return Section(title=title, blocks=list(blocks))

# --------------------------------------------------------------
def test_generated_sections_make_one_page():
    """The toc and benchmark pages hold no blocks."""
    deck = parse_deck(SIMPLE)
    pages = L.paginate(deck.sections[:2])
    assert len(pages) == 2
    assert pages[0].kind == KIND_TOC
    assert pages[0].items == []
    assert pages[1].kind == KIND_BENCHMARKS

# --------------------------------------------------------------
def test_short_blocks_share_a_page():
    """Small blocks are packed together, up to the cap."""
    blocks = [
        make_block(f"H{i}", ["short bullet"], "img.jpg")
        for i in range(3)
    ]
    pages = L.paginate([make_section(blocks)])
    assert len(pages) == 1
    assert len(pages[0].items) == 3

# --------------------------------------------------------------
def test_block_cap_per_page():
    """No page ever carries more than MAX_BLOCKS."""
    blocks = [
        make_block(f"H{i}", ["short"], "img.jpg")
        for i in range(10)
    ]
    pages = L.paginate([make_section(blocks)])
    for page in pages:
        assert len(page.items) <= L.MAX_BLOCKS
    assert sum(len(p.items) for p in pages) == 10

# --------------------------------------------------------------
def test_long_section_splits_across_pages():
    """A section longer than one slide is paginated."""
    bullets = ["a fairly long bullet line " * 4] * 6
    blocks = [
        make_block(f"Story number {i}", bullets, "img.jpg")
        for i in range(6)
    ]
    pages = L.paginate([make_section(blocks)])
    assert len(pages) > 1
    for page in pages:
        assert page.title == "News"

# --------------------------------------------------------------
def test_pages_stay_inside_the_band():
    """Placed blocks never run past the content band."""
    bullets = ["some text that wraps a couple of times " * 3]
    blocks = [
        make_block(f"H{i}", bullets * 2, "img.jpg")
        for i in range(7)
    ]
    for page in L.paginate([make_section(blocks)]):
        for item in page.items:
            bottom = item.text.y + item.text.h
            assert bottom <= L.BAND_BOT + 0.02, bottom
            assert item.text.x >= L.MARGIN - 0.001

# --------------------------------------------------------------
def test_rects_stay_on_the_slide():
    """Every rectangle fits within the slide itself."""
    bullets = ["text " * 30]
    blocks = [
        make_block("Headline here", bullets, "img.jpg")
        for _ in range(5)
    ]
    for page in L.paginate([make_section(blocks)]):
        for item in page.items:
            for rect in [item.text, item.image]:
                if rect is None:
                    continue
                assert rect.x >= -0.001
                assert rect.y >= -0.001
                assert rect.x + rect.w <= L.SLIDE_W + 0.01
                assert rect.y + rect.h <= L.SLIDE_H + 0.01

# --------------------------------------------------------------
def test_huge_block_gets_a_smaller_font():
    """A block too tall at 12 pt is stepped down."""
    giant = ["a long sentence about one news item " * 6] * 9
    block = make_block("Enormous story", giant, "img.jpg")
    pages = L.paginate([make_section([block])])
    assert len(pages) == 1
    assert pages[0].items[0].size < L.BASE_SIZE

# --------------------------------------------------------------
def test_block_without_image_gets_full_width():
    """Text spans the slide when there is no picture."""
    block = make_block("No picture", ["one bullet"])
    page = L.paginate([make_section([block])])[0]
    assert page.items[0].image is None
    assert abs(page.items[0].text.w - L.TEXT_W_FULL) < 0.001

# --------------------------------------------------------------
def test_image_box_sits_right_of_the_text():
    """The picture box never overlaps the text box."""
    block = make_block("With picture", ["one bullet"], "a.jpg")
    item = L.paginate([make_section([block])])[0].items[0]
    assert item.image is not None
    assert item.image.x >= item.text.x + item.text.w

# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
