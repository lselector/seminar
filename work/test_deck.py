#!/usr/bin/env python3
"""
Tests for the deck parser and the automatic layout.

These two modules hold all of the real logic in the
pipeline: grammar parsing and layout arithmetic. Both are
pure functions, so they are tested on strings and numbers
with no files, no network, and no PowerPoint library.

The layout tests are the important ones. Since the author
never positions anything by hand, a mistake in the packing
arithmetic would silently push text off the slide, and
nobody would notice until the seminar.

No pytest needed. Each test is a function whose name starts
with "test_". Layout and measurement live in test_layout.py.

Usage:
    python3 test_deck.py

Created: 2026-09-12
Last updated: 2026-09-12
"""

import deck_layout as L
from test_runner import run
from deck_parser import (
    DeckError, KIND_BENCHMARKS, KIND_CONTENT, KIND_TOC,
    all_headlines, image_manifest, is_link, parse_deck,
)

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


DELETED = """
# T

## slide: News

### Kept story
![](images/a.jpg)
<!-- src: https://e.com/a.png -->
- still here

### Thrown out story
<!-- deleted -->
![](images/b.jpg)
<!-- src: https://e.com/b.png -->
- the human dropped this

## slide: Whole section gone
<!-- deleted -->

### Something
- also gone
"""

SIDECAR = """
# Topics thrown out of AI News - Sept 18, 2026

Anything that is not a heading is a note and ignored.

### The music industry stops fighting
- was here, not wanted
- dropped because it is not technical enough

### Another rejected topic
"""

MARKED = """# AI News - Sept 18, 2026

## slide: toc

## slide: AI News

### Story that stays
![](images/a.jpg)
<!-- src: https://e.com/a.png -->
- **bold** bullet
- https://e.com/a

### Story that goes
<!-- deleted -->
![](images/b.jpg)
<!-- src: https://e.com/b.png -->
- dropped on purpose

## slide: Whole section that goes
<!-- deleted -->

### Inside the doomed section
- goes with its section

## slide: Jobs and Layoffs

### Layoffs tracker
- stays
"""

# (headline a, headline b, is it one topic?)
TOPIC_CASES = [
    ("The music industry stops fighting",
     "The music industry stops fighting", True),
    ("  the MUSIC industry stops Fighting  ",
     "The music industry stops fighting", True),
    ("The music industry stops fighting",
     "The music industry stops fighting and starts "
     "licensing", True),
    ("Suno licenses music with Warner and BMG",
     "The music industry stops fighting and starts "
     "licensing", False),
    ("Mistral raises the largest round in European tech",
     "Mistral raises the largest round in European tech "
     "history", True),
    ("Meta ships Muse agent", "Claude formalized Fermat",
     False),
    ("", "anything at all", False),
]


# --------------------------------------------------------------
def load_checker():
    """Load check_preserved.py by path, or return None."""
    import importlib.util
    import os
    path = os.path.join(
        "..", ".claude", "skills", "slides-update",
        "tools", "check_preserved.py"
    )
    if not os.path.isfile(path):
        return None
    spec = importlib.util.spec_from_file_location(
        "check_preserved", path
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module



# --------------------------------------------------------------
def test_deck_title():
    """The level-one heading becomes the deck title."""
    deck = parse_deck(SIMPLE)
    assert deck.title == "AI News - Sept 18, 2026"


# --------------------------------------------------------------
def test_section_kinds():
    """Named sections map to generated slide kinds."""
    deck = parse_deck(SIMPLE)
    kinds = [s.kind for s in deck.sections]
    assert kinds == [KIND_TOC, KIND_BENCHMARKS, KIND_CONTENT]
    assert deck.sections[2].title == "AI News"
    assert deck.sections[0].title == ""


# --------------------------------------------------------------
def test_blocks_and_bullets():
    """Blocks collect their headline and their bullets."""
    deck = parse_deck(SIMPLE)
    blocks = deck.sections[2].blocks
    assert len(blocks) == 2
    assert blocks[0].headline == "First story"
    assert len(blocks[0].bullets) == 3
    assert len(blocks[1].bullets) == 1


# --------------------------------------------------------------
def test_image_manifest():
    """Images carry their src or shot source through."""
    deck = parse_deck(SIMPLE)
    manifest = image_manifest(deck)
    assert manifest == [
        ("images/one.jpg",
         "https://example.com/one.png", "src"),
        ("images/two.jpg",
         "https://example.com/page", "shot"),
    ]


# --------------------------------------------------------------
def test_local_image_has_no_source():
    """An image with no comment is a local original."""
    deck = parse_deck(
        "# T\n## slide: X\n### H\n![](images/me.jpg)\n"
    )
    assert image_manifest(deck) == [
        ("images/me.jpg", None, "local")
    ]


# --------------------------------------------------------------
def test_headlines_feed_the_contents():
    """Every headline is available for the toc page."""
    deck = parse_deck(SIMPLE)
    assert all_headlines(deck) == [
        "First story", "Second story"
    ]


# --------------------------------------------------------------
def test_locked_block():
    """A locked flag under ### locks that item only."""
    deck = parse_deck(
        "# T\n## slide: X\n"
        "### Kept\n<!-- locked -->\n- hand written\n"
        "### Open\n- generated\n"
    )
    blocks = deck.sections[0].blocks
    assert blocks[0].locked is True
    assert blocks[1].locked is False
    assert deck.sections[0].locked is False


# --------------------------------------------------------------
def test_locked_section():
    """A locked flag under ## locks the whole section."""
    deck = parse_deck(
        "# T\n## slide: About\n<!-- locked -->\n"
        "### Me\n- a bullet\n"
    )
    assert deck.sections[0].locked is True
    assert deck.sections[0].blocks[0].locked is False


# --------------------------------------------------------------
def test_locked_does_not_disturb_content():
    """The flag is not mistaken for a bullet or image."""
    deck = parse_deck(
        "# T\n## slide: X\n### H\n<!-- locked -->\n"
        "![](images/a.jpg)\n<!-- src: https://e.com/a.png -->\n"
        "- one\n- two\n"
    )
    block = deck.sections[0].blocks[0]
    assert block.bullets == ["one", "two"]
    assert block.image == "images/a.jpg"
    assert block.image_src == "https://e.com/a.png"
    assert block.locked is True


# --------------------------------------------------------------
def test_deleted_block_is_recorded():
    """A tombstone marks the block and locks it."""
    deck = parse_deck(DELETED)
    blocks = deck.sections[0].blocks
    assert blocks[1].deleted is True
    assert blocks[1].locked is True
    assert blocks[0].deleted is False


# --------------------------------------------------------------
def test_deleted_leaves_the_contents():
    """A tombstoned item gets no contents entry."""
    deck = parse_deck(DELETED)
    assert all_headlines(deck) == ["Kept story"]


# --------------------------------------------------------------
def test_deleted_image_is_not_fetched():
    """A tombstoned item's picture is not downloaded."""
    deck = parse_deck(DELETED)
    paths = [item[0] for item in image_manifest(deck)]
    assert paths == ["images/a.jpg"]


# --------------------------------------------------------------
def test_deleted_makes_no_slides():
    """Tombstoned items and sections produce no pages."""
    deck = parse_deck(DELETED)
    pages = L.paginate(deck.sections)
    assert len(pages) == 1
    headlines = [
        item.block.headline for item in pages[0].items
    ]
    assert headlines == ["Kept story"]


# --------------------------------------------------------------
def test_deleted_counts_are_reported():
    """Both kinds of tombstone are counted."""
    from deck_parser import deleted_count
    assert deleted_count(parse_deck(DELETED)) == 2


# --------------------------------------------------------------
def test_deleted_section_keeps_its_blocks_in_the_file():
    """A tombstone hides content without removing it."""
    deck = parse_deck(DELETED)
    assert len(deck.sections) == 2
    assert len(deck.sections[1].blocks) == 1
    assert deck.sections[1].deleted is True


# --------------------------------------------------------------
def test_sidecar_path_is_derived():
    """The deleted file sits beside the deck file."""
    from deck_parser import deleted_sidecar_path
    assert deleted_sidecar_path(
        "work/2026-09-18-AI-News.md"
    ) == "work/2026-09-18-AI-News-deleted.md"


# --------------------------------------------------------------
def test_sidecar_topics_are_read():
    """Only the headings count, notes are ignored."""
    from deck_parser import parse_deleted_topics
    assert parse_deleted_topics(SIDECAR) == [
        "The music industry stops fighting",
        "Another rejected topic",
    ]


# --------------------------------------------------------------
def test_missing_sidecar_is_not_an_error():
    """A deck with no deleted file has no deleted topics."""
    from deck_parser import read_deleted_topics
    assert read_deleted_topics("/tmp/no-such-deck.md") == []


# --------------------------------------------------------------
def test_recreated_topic_is_caught():
    """A topic back in the deck is reported."""
    from deck_parser import (
        parse_deleted_topics, recreated_topics,
    )
    deck = parse_deck(
        "# T\n## slide: News\n"
        "### The music industry stops fighting\n- back again\n"
        "### A fresh story\n- fine\n"
    )
    topics = parse_deleted_topics(SIDECAR)
    found = recreated_topics(deck, topics)
    assert len(found) == 1
    assert found[0][0] == "The music industry stops fighting"


# --------------------------------------------------------------
def test_recreated_match_is_forgiving():
    """Case and spacing changes do not slip past."""
    from deck_parser import recreated_topics
    deck = parse_deck(
        "# T\n## slide: News\n"
        "###   the MUSIC industry   stops fighting\n- x\n"
    )
    found = recreated_topics(
        deck, ["The music industry stops fighting"]
    )
    assert len(found) == 1


# --------------------------------------------------------------
def test_topic_matching_cases():
    """same_topic behaves on a table of real headlines."""
    from deck_parser import same_topic
    for one, other, expected in TOPIC_CASES:
        got = same_topic(one, other)
        assert got is expected, f"{one!r} vs {other!r}"


# --------------------------------------------------------------
def test_checker_agrees_on_topic_matching():
    """The checker's copy of same_topic must not drift."""
    from deck_parser import same_topic
    checker = load_checker()
    if checker is None:
        print("    (skipped: checker not found)")
        return
    for one, other, _ in TOPIC_CASES:
        assert checker.same_topic(one, other) is \
            same_topic(one, other), f"{one!r} vs {other!r}"


# --------------------------------------------------------------
def test_clean_deck_has_no_recreated_topics():
    """A deck with none of the deleted topics passes."""
    from deck_parser import (
        parse_deleted_topics, recreated_topics,
    )
    deck = parse_deck(SIMPLE)
    topics = parse_deleted_topics(SIDECAR)
    assert recreated_topics(deck, topics) == []


# --------------------------------------------------------------
def test_source_split_is_lossless():
    """Splitting and rejoining reproduces the file."""
    import move_deleted as M
    preamble, segments = M.split_source(MARKED)
    rebuilt = "\n".join(
        preamble + [s.text() for s in segments]
    )
    assert rebuilt == MARKED.rstrip("\n")


# --------------------------------------------------------------
def test_plan_moves_picks_flagged_items():
    """Flagged items and whole sections are selected."""
    import move_deleted as M
    _, segments = M.split_source(MARKED)
    keep, move = M.plan_moves(segments)
    moved = [M.heading_label(s) for s in move]
    assert moved == [
        "Story that goes",
        "slide: Whole section that goes",
        "Inside the doomed section",
    ]
    assert "Story that stays" in \
        [M.heading_label(s) for s in keep]


# --------------------------------------------------------------
def test_moved_text_keeps_its_content():
    """Bullets, images, and links survive the move."""
    import move_deleted as M
    _, segments = M.split_source(MARKED)
    _, move = M.plan_moves(segments)
    text = M.render_moved(move, "2026-09-12")
    assert "![](images/b.jpg)" in text
    assert "<!-- src: https://e.com/b.png -->" in text
    assert "- dropped on purpose" in text
    assert "<!-- moved: 2026-09-12 -->" in text


# --------------------------------------------------------------
def test_moved_text_drops_the_flag():
    """The marker does not travel to the sidecar."""
    import move_deleted as M
    _, segments = M.split_source(MARKED)
    _, move = M.plan_moves(segments)
    assert "<!-- deleted -->" not in \
        M.render_moved(move, "2026-09-12")


# --------------------------------------------------------------
def test_remaining_deck_still_parses():
    """What is left behind is a valid deck."""
    import move_deleted as M
    preamble, segments = M.split_source(MARKED)
    keep, _ = M.plan_moves(segments)
    deck = parse_deck(M.render_deck(preamble, keep))
    heads = [b.headline for s in deck.sections
             for b in s.blocks]
    assert heads == ["Story that stays", "Layoffs tracker"]
    assert "<!-- deleted -->" not in \
        M.render_deck(preamble, keep)


# --------------------------------------------------------------
def test_nothing_marked_moves_nothing():
    """A deck with no flags loses nothing."""
    import move_deleted as M
    preamble, segments = M.split_source(SIMPLE)
    keep, move = M.plan_moves(segments)
    assert move == []
    assert len(keep) == len(segments)


# --------------------------------------------------------------
def test_notoc_keeps_back_matter_off_the_contents():
    """Author and thank-you pages are not news."""
    deck = parse_deck(
        "# T\n## slide: News\n### Real story\n- x\n"
        "## slide: About the Speaker\n<!-- notoc -->\n"
        "### Lev Selector, Ph.D.\n- y\n"
    )
    assert all_headlines(deck) == ["Real story"]
    assert deck.sections[1].notoc is True


# --------------------------------------------------------------
def test_notoc_does_not_lock_or_delete():
    """Hiding from the contents is not a tombstone."""
    deck = parse_deck(
        "# T\n## slide: X\n<!-- notoc -->\n### H\n- y\n"
    )
    section = deck.sections[0]
    assert section.notoc is True
    assert section.locked is False
    assert section.deleted is False
    assert len(L.paginate(deck.sections)) == 1


# --------------------------------------------------------------
def test_notoc_on_one_item_only():
    """The flag under ### hides just that item."""
    deck = parse_deck(
        "# T\n## slide: X\n### Shown\n- a\n"
        "### Hidden\n<!-- notoc -->\n- b\n"
    )
    assert all_headlines(deck) == ["Shown"]
    assert len(deck.sections[0].blocks) == 2


# --------------------------------------------------------------
def test_epigraph_is_parsed():
    """A blockquote above the sections is the epigraph."""
    deck = parse_deck(
        "# AI News\n\n> Agents got their own computers.\n"
        "\n## slide: News\n### A story\n- x\n"
    )
    assert deck.epigraph == "Agents got their own computers."
    assert deck.title == "AI News"
    assert len(deck.sections) == 1


# --------------------------------------------------------------
def test_epigraph_is_optional():
    """A deck with no blockquote has no epigraph."""
    assert parse_deck(SIMPLE).epigraph == ""


# --------------------------------------------------------------
def test_epigraph_joins_wrapped_lines():
    """Two quote lines become one epigraph."""
    deck = parse_deck(
        "# T\n> First half\n> second half\n"
        "## slide: X\n### H\n- y\n"
    )
    assert deck.epigraph == "First half second half"


# --------------------------------------------------------------
def test_quote_inside_a_section_is_not_the_epigraph():
    """Only the preamble carries the epigraph."""
    deck = parse_deck(
        "# T\n## slide: X\n### H\n> not an epigraph\n- y\n"
    )
    assert deck.epigraph == ""


# --------------------------------------------------------------
def test_bare_heading_makes_a_headless_block():
    """A "###" on its own is a block of only bullets."""
    deck = parse_deck(
        "# T\n## slide: X\n###\n- one\n- two\n"
    )
    block = deck.sections[0].blocks[0]
    assert block.headline == ""
    assert block.bullets == ["one", "two"]


# --------------------------------------------------------------
def test_headless_block_adds_no_contents_entry():
    """An absent headline leaves the contents alone."""
    deck = parse_deck(
        "# T\n## slide: X\n### Real one\n- a\n"
        "###\n- b\n"
    )
    assert all_headlines(deck) == ["Real one"]


# --------------------------------------------------------------
def test_headless_block_is_not_a_deleted_topic():
    """An empty headline matches nothing in the sidecar."""
    from deck_parser import recreated_topics
    deck = parse_deck("# T\n## slide: X\n###\n- a\n")
    assert recreated_topics(deck, ["Some topic"]) == []


# --------------------------------------------------------------
def test_markup_splits_into_styled_pieces():
    """Bold, highlight and code each get their own run."""
    from deck_parser import split_markup
    pieces = split_markup(
        "Run `pip install x` with **care** and ==note=="
    )
    assert pieces == [
        ("Run ", "plain"),
        ("pip install x", "code"),
        (" with ", "plain"),
        ("care", "bold"),
        (" and ", "plain"),
        ("note", "hilite"),
    ]


# --------------------------------------------------------------
def test_strip_markup_leaves_only_words():
    """Measuring must not count the markup characters."""
    from deck_parser import strip_markup
    assert strip_markup("a **b** `c` ==d==") == "a b c d"
    assert strip_markup("plain text") == "plain text"


# --------------------------------------------------------------
def test_markup_styles_agree():
    """deck_parser and pptx_text share one vocabulary."""
    import deck_parser as P
    import pptx_text as T
    assert P.STYLE_BOLD == T.STYLE_BOLD
    assert P.STYLE_HILITE == T.STYLE_HILITE
    assert P.STYLE_CODE == T.STYLE_CODE


# --------------------------------------------------------------
def test_red_markup_is_its_own_style():
    """!!text!! marks a line for bold red."""
    from deck_parser import split_markup, strip_markup
    assert split_markup("a !!loud!! b") == [
        ("a ", "plain"), ("loud", "red"), (" b", "plain")
    ]
    assert strip_markup("a !!loud!! b") == "a loud b"


# --------------------------------------------------------------
def test_link_detection():
    """Only bare URLs count as link bullets."""
    assert is_link("https://example.com")
    assert is_link("http://example.com/a/b?c=1")
    assert not is_link("see https://example.com for more")
    assert not is_link("plain text")


# --------------------------------------------------------------
def expect_error(text, hint):
    """Assert that parsing text raises DeckError."""
    try:
        parse_deck(text)
    except DeckError as exc:
        assert hint in str(exc), f"{hint!r} not in {exc}"
        return
    raise AssertionError(f"expected DeckError for {hint}")


# --------------------------------------------------------------
def test_grammar_errors():
    """Malformed decks fail loudly, never silently."""
    expect_error(
        "# T\n### orphan\n", "before any"
    )
    expect_error(
        "# T\n## AI News\n", "must be written"
    )
    expect_error(
        "# T\n## slide: X\n### H\n![](a.jpg)\n![](b.jpg)\n",
        "already has an image"
    )
    expect_error("# T\n", "no '## slide:'")

# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
