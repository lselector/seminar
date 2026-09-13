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
    assert width == int(L.MAX_IMG_W * 150)
    assert height == int(L.BAND_H * 150)
    assert settings.geometry() == f"{width}x{height}>"

    narrow, _ = settings.box(L.IMG_W)
    assert narrow == int(L.IMG_W * 150) < width


# --------------------------------------------------------------
def test_each_picture_is_sized_to_its_own_box():
    """A promo picture is stored wider than a news one."""
    deck = parse_deck(
        "# T\n"
        "## slide: News\n### A story\n"
        "![](images/news.jpg)\n- a line\n"
        "## slide: Channel\n<!-- promo -->\n###\n"
        "![](images/promo.jpg)\n- a line\n"
        "## slide: About\n<!-- profile -->\n### Me\n"
        "![](images/face.jpg)\n- a line\n"
    )
    widths = L.image_widths(deck.sections)
    assert widths["promo"] == L.PROMO_IMG_W
    assert widths["news"] == L.IMG_W
    assert widths["face"] == L.PROFILE_IMG_W
    assert widths["promo"] > widths["news"] > widths["face"]

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
        assert item.image.w == L.PROFILE_IMG_W
        assert item.text.x > item.image.x + item.image.w


# --------------------------------------------------------------
def test_profile_pair_is_centred_both_ways():
    """Portrait and details sit centred as one pair."""
    item = profile_item()
    left = min(item.image.x, item.text.x)
    right = max(
        item.image.x + item.image.w,
        item.text.x + item.text.w,
    )
    assert abs((left + right) / 2 - L.SLIDE_W / 2) < 1e-9

    top = min(item.image.y, item.text.y)
    bottom = max(
        item.image.y + item.image.h,
        item.text.y + item.text.h,
    )
    middle = L.BAND_TOP + L.BAND_H / 2
    assert abs((top + bottom) / 2 - middle) < 1e-9


# --------------------------------------------------------------
def test_profile_picture_is_square_and_does_not_span():
    """It stays a portrait, not a stretched banner."""
    item = profile_item()
    assert item.image.w == item.image.h == L.PROFILE_IMG_W
    pair = (
        L.PROFILE_IMG_W + L.PROFILE_GAP + item.text.w
    )
    assert pair < L.CONTENT_W


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
def closing_page():
    """Lay out a sign-off page and return it."""
    deck = parse_deck(
        "# T\n## slide: Thank You!\n<!-- closing -->\n"
        "### Questions and discussion\n"
        "- Videos of every seminar are on YouTube\n"
        "- https://www.youtube.com/@lev-selector\n"
    )
    return L.paginate(deck.sections)[0]


# --------------------------------------------------------------
def test_closing_title_sits_a_third_down():
    """The sign-off title drops to a third of the page."""
    rect = L.closing_title_rect()
    assert abs(rect.y - L.SLIDE_H / 3) < 1e-9
    assert L.CLOSING_TITLE_SIZE == 40


# --------------------------------------------------------------
def test_closing_boxes_are_centred():
    """Title and links share the slide's centre line."""
    page = closing_page()
    middle = L.SLIDE_W / 2
    title = L.closing_title_rect()
    assert abs(title.x + title.w / 2 - middle) < 1e-9
    body = page.items[0].text
    assert abs(body.x + body.w / 2 - middle) < 1e-9


# --------------------------------------------------------------
def test_closing_body_sits_under_the_title():
    """The links clear the title without overlapping."""
    page = closing_page()
    title = L.closing_title_rect()
    body = page.items[0].text
    assert body.y >= title.y + title.h
    assert body.y + body.h <= L.BAND_BOT + 0.01


# --------------------------------------------------------------
def test_closing_is_one_page_with_no_picture():
    """A sign-off section never paginates or takes art."""
    page = closing_page()
    assert page.closing is True
    assert len(page.items) == 1
    assert page.items[0].image is None


# --------------------------------------------------------------
def test_closing_does_not_affect_news():
    """Ordinary slides keep the top-left title."""
    block = make_block("A story", ["one line"], "a.jpg")
    page = L.paginate([make_section([block])])[0]
    assert page.closing is False
    assert page.items[0].text.y > L.BAND_TOP


# --------------------------------------------------------------
def test_headless_block_reserves_no_headline_room():
    """A block with no heading is shorter by one line."""
    assert L.headline_height("", 6.0, 12) == 0.0
    with_head = L.block_height(
        make_block("A heading", ["one line"]), 12
    )
    without = L.block_height(make_block("", ["one line"]), 12)
    assert without < with_head


# --------------------------------------------------------------
def test_closing_page_needs_no_headline():
    """The sign-off page lays out from bullets alone."""
    deck = parse_deck(
        "# T\n## slide: Thank You!\n<!-- closing -->\n"
        "###\n- Videos are on YouTube\n"
        "- https://www.youtube.com/@lev-selector\n"
    )
    page = L.paginate(deck.sections)[0]
    assert page.closing is True
    assert page.items[0].block.headline == ""
    body = page.items[0].text
    assert abs(body.x + body.w / 2 - L.SLIDE_W / 2) < 1e-9


# --------------------------------------------------------------
def test_special_pages_draw_no_box():
    """Author and sign-off pages carry no fill or border."""
    for src in (
        "# T\n## slide: About\n<!-- profile -->\n"
        "### Me\n![](images/me.jpg)\n- a line\n",
        "# T\n## slide: Thanks\n<!-- closing -->\n"
        "###\n- a line\n",
    ):
        page = L.paginate(parse_deck(src).sections)[0]
        assert page.plain is True


# --------------------------------------------------------------
def test_news_pages_still_draw_a_box():
    """Ordinary slides keep the yellow box."""
    block = make_block("A story", ["one line"], "a.jpg")
    page = L.paginate([make_section([block])])[0]
    assert page.plain is False


# --------------------------------------------------------------
def test_title_box_hugs_its_text():
    """A short title does not reserve the whole width."""
    narrow = L.title_width("AI News")
    wide = L.title_width("Artificial Analysis Intelligence")
    assert narrow < wide < L.CONTENT_W


# --------------------------------------------------------------
def test_title_box_leaves_slack_for_the_real_font():
    """Measured in Arial, rendered in Calibri Bold."""
    import text_metrics as M
    text = "About the Speaker"
    measured = M.text_width(text, L.TITLE_SIZE, bold=True)
    assert L.title_width(text) > measured
    assert L.TITLE_SLACK > 1.0


# --------------------------------------------------------------
def test_title_box_respects_its_limit():
    """The contents title leaves room for the epigraph."""
    limit = 2.0
    assert L.title_width("A very long slide title indeed",
                         limit) == limit
    assert L.title_width("AI News", limit) < limit


# --------------------------------------------------------------
def test_regular_slides_are_capped_at_twelve():
    """News body text is Calibri 12, never larger."""
    short = make_block("Short one", ["a brief line"])
    page = L.paginate([make_section([short])])[0]
    assert page.items[0].size == L.BASE_SIZE == 12


# --------------------------------------------------------------
def test_special_pages_may_go_larger():
    """The author page still gets its bigger text."""
    deck = parse_deck(
        "# T\n## slide: About\n<!-- profile -->\n"
        "### Me\n![](images/me.jpg)\n- a line\n"
    )
    item = L.paginate(deck.sections)[0].items[0]
    assert item.size > L.BASE_SIZE


# --------------------------------------------------------------
def test_code_is_fixed_width_nine_point_blue():
    """Code runs match what the older decks used."""
    import pptx_text as T
    assert T.CODE_SIZE == 9
    assert str(T.COLOR_CODE) == "3C78D8"
    assert T.CODE_FONT == "Consolas"


# --------------------------------------------------------------
def test_headline_is_red():
    """Every box title is bold and red."""
    import pptx_text as T
    assert str(T.COLOR_HEADLINE) == "FF0000"


# --------------------------------------------------------------
def test_markup_does_not_inflate_the_height():
    """Markup characters are not measured as text."""
    plain = L.bullet_height("a b c d", 6.0, 12)
    marked = L.bullet_height("a **b** `c` ==d==", 6.0, 12)
    assert plain == marked


# --------------------------------------------------------------
def test_benchmarks_run_at_nine_point():
    """The busiest page uses smaller text than the rest."""
    import bench_page as B
    assert B.BENCH_SIZE == 9 < L.BASE_SIZE
    assert B.BENCH_CAPTION_SIZE == 10
    assert B.bench_font_size(25, 10.0) == 9


# --------------------------------------------------------------
def test_benchmark_columns_sit_together():
    """The pair is centred, not pinned to the two edges."""
    import bench_page as B
    boards = [
        {"label": "English", "url": "https://a.example",
         "entries": [
             {"score": 1500, "name": "m", "vendor": "other"}
         ]},
        {"label": "Coding", "url": "https://b.example",
         "entries": [
             {"score": 1500, "name": "m", "vendor": "other"}
         ]},
    ]
    size = B.BENCH_SIZE
    widths = [B.column_width(b, size) for b in boards]
    total = sum(widths) + B.BENCH_GAP
    left = (L.SLIDE_W - total) / 2
    assert abs(left + total / 2 - L.SLIDE_W / 2) < 1e-9
    assert left > L.MARGIN


# --------------------------------------------------------------
def test_benchmark_column_hugs_its_rows():
    """No empty yellow below the last model."""
    import bench_page as B
    tall = B.column_height(25, B.BENCH_SIZE)
    assert tall < L.BAND_BOT - B.bench_top()
    assert B.column_height(10, 9) < B.column_height(25, 9)


# --------------------------------------------------------------
def test_date_note_clears_the_columns():
    """The cutoff box is not hidden by anything."""
    import bench_page as B
    assert B.DATE_SIZE == L.BASE_SIZE == 12
    header_bottom = B.HEADER_Y + B.header_height()
    assert B.bench_top() > header_bottom


# --------------------------------------------------------------
def gaps_around(page):
    """White space above, between and below the boxes."""
    boxes = sorted(
        (i.text.y, i.text.h) for i in page.items
    )
    edges = [boxes[0][0] - L.BAND_TOP]
    for (y, h), (nxt, _) in zip(boxes, boxes[1:]):
        edges.append(nxt - (y + h))
    edges.append(L.BAND_BOT - (boxes[-1][0] + boxes[-1][1]))
    return edges


# --------------------------------------------------------------
def test_boxes_never_touch_the_title():
    """There is always white space under the title."""
    for count in (1, 2, 3):
        blocks = [
            make_block(f"H{i}", ["a short line"], "a.jpg")
            for i in range(count)
        ]
        page = L.paginate([make_section(blocks)])[0]
        assert gaps_around(page)[0] >= L.BLOCK_GAP


# --------------------------------------------------------------
def test_dense_slides_share_the_space_evenly():
    """Above, between and below come out the same."""
    bullets = ["a fairly long sentence here " * 3] * 9
    blocks = [
        make_block(f"Headline {i}", bullets, "a.jpg")
        for i in range(2)
    ]
    page = L.paginate([make_section(blocks)])[0]
    edges = gaps_around(page)
    assert max(edges) < L.MAX_BLOCK_GAP, "not a dense slide"
    assert max(edges) - min(edges) < 0.01


# --------------------------------------------------------------
def test_sparse_slides_do_not_float_one_box():
    """A lone short box stays near the top, not centred."""
    block = make_block("Short", ["one line"], "a.jpg")
    page = L.paginate([make_section([block])])[0]
    top, bottom = gaps_around(page)
    assert top <= L.MAX_BLOCK_GAP + 1e-6
    assert bottom > top


# --------------------------------------------------------------
def promo_page():
    """Lay out the channel page and return it."""
    deck = parse_deck(
        "# T\n## slide: Weekly videos\n<!-- promo -->\n"
        "###\n![](images/yt.jpg)\n"
        "- Weekly videos every Friday\n"
        "- https://www.youtube.com/@lev-selector\n"
    )
    return L.paginate(deck.sections)[0]


# --------------------------------------------------------------
def test_promo_text_is_large():
    """The call to action reads at 22 pt, links at 18."""
    item = promo_page().items[0]
    assert item.size == L.PROMO_SIZE == 22
    assert L.link_size(item.size) == L.PROMO_LINK_SIZE == 18


# --------------------------------------------------------------
def test_promo_puts_text_left_picture_right():
    """A wide picture beside a narrower column of text."""
    item = promo_page().items[0]
    assert item.text.x == L.MARGIN
    assert item.text.w == L.PROMO_TEXT_W
    assert item.image.x > item.text.x + item.text.w
    assert item.image.w == L.PROMO_IMG_W > L.IMG_W


# --------------------------------------------------------------
def test_promo_does_not_change_news_slides():
    """Regular slides keep 12 pt and the narrow picture."""
    block = make_block("A story", ["one line"], "a.jpg")
    item = L.paginate([make_section([block])])[0].items[0]
    assert item.size == L.BASE_SIZE
    assert item.image.w == L.IMG_W


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
def test_font_ladder_caps_at_the_body_size():
    """Sizes above BASE_SIZE exist but are not offered
    to a regular slide."""
    assert max(L.FONT_LADDER) > L.BASE_SIZE
    assert L.BASE_SIZE in L.FONT_LADDER
    assert max(L.ladder_below(L.BASE_SIZE)) == L.BASE_SIZE
    assert max(L.ladder_below(max(L.FONT_LADDER))) > 12
    assert L.ladder_below(L.BASE_SIZE) == [12, 11, 10, 9, 8]

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
