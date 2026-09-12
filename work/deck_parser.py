#!/usr/bin/env python3
"""
Parse a weekly seminar deck Markdown file.

Turns the deck file into plain data objects that carry no
layout and no PowerPoint types. This module is pure text
in, data out: it touches no network, no images, and no
PPTX library, so it can be unit tested on strings alone.

The grammar it accepts (see ADD.md):

    # AI News - Sept 18, 2026      deck title
    ## slide: toc                  generated contents
    ## slide: benchmarks           generated LM Arena
    ## slide: AI News              a content section
    ### Headline of one item       a news block
    ![](images/foo.jpg)            that block's image
    <!-- src: https://... -->      download the image
    <!-- shot: https://... -->     screenshot the page
    > a short line for the week     the epigraph
    <!-- locked -->                hand-edited, keep as is
    <!-- notoc -->                 keep off the contents
    <!-- profile -->               picture left, text right
    - a body bullet                body text
    - https://example.com          rendered as a link

A blockquote line above the first section is the deck
epigraph: one short, pointed sentence about the week, drawn
in red at the top right of the contents page.

"locked" marks text a human wrote or rewrote, so nothing
may change it. Under a "###" it covers that news item;
directly under a "## slide:" heading it covers the whole
section.

Deleting a topic works differently, and does not happen in
this file at all. The author cuts the item out and pastes
it into the sidecar file next to the deck:

    2026-09-18-AI-News.md            the deck
    2026-09-18-AI-News-deleted.md    topics thrown out

Anything listed there must never be written back into the
deck. `read_deleted_topics` returns those headlines, s3
warns if one reappears, and check_preserved.py fails the
edit pass. The older in-file `<!-- deleted -->` flag is
still understood, and still keeps an item off the slides,
but the sidecar is the workflow to use.

A section may hold any number of blocks. Splitting them
across physical slides is deck_layout's job, not this
module's.

Usage:
    from deck_parser import parse_deck, parse_deck_file
    deck = parse_deck_file("2026-09-18-AI-News.md")
    print(deck.title, len(deck.sections))

Created: 2026-09-12
Last updated: 2026-09-12
"""

import os
import re
from dataclasses import dataclass, field

KIND_CONTENT = "content"
KIND_TOC = "toc"
KIND_BENCHMARKS = "benchmarks"

GENERATED_KINDS = (KIND_TOC, KIND_BENCHMARKS)

RE_IMAGE = re.compile(r"^!\[[^\]]*\]\(([^)]+)\)\s*$")
RE_COMMENT = re.compile(r"^<!--\s*(\w+)\s*:\s*(.+?)\s*-->$")
RE_FLAG = re.compile(r"^<!--\s*(\w+)\s*-->$")
RE_BULLET = re.compile(r"^[-*]\s+(.*)$")
RE_QUOTE = re.compile(r"^>\s*(.*)$")

FLAG_LOCKED = "locked"
FLAG_DELETED = "deleted"
FLAG_NOTOC = "notoc"
FLAG_PROFILE = "profile"
FLAGS = (
    FLAG_LOCKED, FLAG_DELETED, FLAG_NOTOC, FLAG_PROFILE,
)

# Sidecar holding the topics a human threw out. Cutting an
# item into this file is how deletion is recorded.
DELETED_SUFFIX = "-deleted.md"


class DeckError(Exception):
    """Raised when the deck file breaks the grammar."""


@dataclass
class Block:
    """One news item: headline, bullets, one image."""

    headline: str
    bullets: list = field(default_factory=list)
    image: str = None
    image_src: str = None
    image_shot: str = None
    locked: bool = False
    deleted: bool = False
    notoc: bool = False
    profile: bool = False


@dataclass
class Section:
    """A titled group of blocks, before pagination."""

    title: str
    kind: str = KIND_CONTENT
    blocks: list = field(default_factory=list)
    locked: bool = False
    deleted: bool = False
    notoc: bool = False
    profile: bool = False


@dataclass
class Deck:
    """A whole week's deck: a title and its sections."""

    title: str = ""
    epigraph: str = ""
    sections: list = field(default_factory=list)


# --------------------------------------------------------------
def is_link(text):
    """Check whether a bullet is a bare URL."""
    stripped = text.strip()
    if " " in stripped:
        return False
    return stripped.startswith(("http://", "https://"))


# --------------------------------------------------------------
def section_kind(title):
    """Map a section title to its slide kind."""
    key = title.strip().lower()
    if key == KIND_TOC:
        return KIND_TOC
    if key == KIND_BENCHMARKS:
        return KIND_BENCHMARKS
    return KIND_CONTENT


# --------------------------------------------------------------
def make_section(title):
    """Build a Section from its heading text."""
    kind = section_kind(title)
    label = "" if kind in GENERATED_KINDS else title.strip()
    return Section(title=label, kind=kind)


# --------------------------------------------------------------
def _need_block(block, lineno, what):
    """Fail when a line appears before any headline."""
    if block is None:
        raise DeckError(
            f"line {lineno}: {what} outside any '###' block"
        )


# --------------------------------------------------------------
def _attach_image(block, path, lineno):
    """Attach an image path to the current block."""
    _need_block(block, lineno, "image")
    if block.image:
        raise DeckError(
            f"line {lineno}: block '{block.headline}' "
            f"already has an image"
        )
    block.image = path.strip()


# --------------------------------------------------------------
def _attach_comment(block, key, value, lineno):
    """Attach a src/shot manifest comment to a block."""
    if key not in ("src", "shot"):
        return
    _need_block(block, lineno, f"{key} comment")
    if key == "src":
        block.image_src = value
    else:
        block.image_shot = value


# --------------------------------------------------------------
def _split_heading(line):
    """Return (level, text) for a Markdown heading."""
    hashes = len(line) - len(line.lstrip("#"))
    return hashes, line[hashes:].strip()


# --------------------------------------------------------------
def _section_title(text, lineno):
    """Strip the 'slide:' prefix from a section heading."""
    if ":" not in text:
        raise DeckError(
            f"line {lineno}: '## {text}' must be written "
            f"as '## slide: <title>'"
        )
    prefix, title = text.split(":", 1)
    if prefix.strip().lower() != "slide":
        raise DeckError(
            f"line {lineno}: unknown section '{prefix}'"
        )
    return title


# --------------------------------------------------------------
def _handle_heading(deck, line, lineno):
    """Apply a heading line, returning the new block."""
    level, text = _split_heading(line)
    if level == 1:
        deck.title = text
        return None
    if level == 2:
        deck.sections.append(make_section(
            _section_title(text, lineno)
        ))
        return None
    if level == 3:
        if not deck.sections:
            raise DeckError(
                f"line {lineno}: '### {text}' before any "
                f"'## slide:' heading"
            )
        section = deck.sections[-1]
        # A flag on the section heading is read before its
        # blocks exist, so new blocks inherit it here.
        block = Block(headline=text, profile=section.profile)
        section.blocks.append(block)
        return block
    raise DeckError(f"line {lineno}: heading too deep")


# --------------------------------------------------------------
def _current_section(deck, lineno, name):
    """Return the open section, or fail loudly."""
    if not deck.sections:
        raise DeckError(
            f"line {lineno}: {name} outside any section"
        )
    return deck.sections[-1]


# --------------------------------------------------------------
def _apply_flag(deck, block, name, lineno):
    """Apply a flag comment to the item it sits under."""
    if name not in FLAGS:
        return
    target = block
    if target is None:
        target = _current_section(deck, lineno, name)

    if name == FLAG_NOTOC:
        target.notoc = True
        return

    if name == FLAG_PROFILE:
        target.profile = True
        for block in getattr(target, "blocks", []):
            block.profile = True
        return

    target.locked = True
    if name == FLAG_DELETED:
        # A tombstone is frozen too. Nothing may edit it,
        # and nothing may bring the topic back.
        target.deleted = True


# --------------------------------------------------------------
def _handle_body(deck, block, line, lineno):
    """Apply an image, comment, flag, or bullet line."""
    match = RE_IMAGE.match(line)
    if match:
        _attach_image(block, match.group(1), lineno)
        return

    match = RE_COMMENT.match(line)
    if match:
        _attach_comment(
            block, match.group(1).lower(),
            match.group(2), lineno
        )
        return

    match = RE_FLAG.match(line)
    if match:
        _apply_flag(
            deck, block, match.group(1).lower(), lineno
        )
        return

    match = RE_BULLET.match(line)
    if match:
        _need_block(block, lineno, "bullet")
        block.bullets.append(match.group(1).strip())


# --------------------------------------------------------------
def parse_deck(text):
    """Parse deck Markdown text into a Deck object."""
    deck = Deck()
    block = None

    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        quote = RE_QUOTE.match(line)
        if quote and not deck.sections:
            deck.epigraph = " ".join(
                filter(None, [deck.epigraph, quote.group(1)])
            ).strip()
            continue
        if line.startswith("#") and not line.startswith("#!"):
            block = _handle_heading(deck, line, lineno)
        else:
            _handle_body(deck, block, line, lineno)

    if not deck.sections:
        raise DeckError("no '## slide:' sections found")
    return deck


# --------------------------------------------------------------
def parse_deck_file(path):
    """Read and parse a deck Markdown file."""
    with open(path, encoding="utf-8") as handle:
        return parse_deck(handle.read())


# --------------------------------------------------------------
def normalize_headline(text):
    """Fold a headline for forgiving comparison."""
    return " ".join(text.lower().split())


# --------------------------------------------------------------
def headline_tokens(text):
    """The words of a headline worth matching on."""
    words = normalize_headline(text).split()
    return {w.strip(".,:;!?()[]'\"") for w in words
            if len(w) > 3}


# --------------------------------------------------------------
def same_topic(one, other):
    """Judge whether two headlines name the same topic.

    Exact comparison is too strict here. A rewritten deck
    rarely reproduces a headline word for word, so a topic
    the author threw out would sneak back under a slightly
    different heading. This also treats one headline being
    a shortening of the other, or sharing most of its
    significant words, as the same topic.

    Kept in step with the copy in check_preserved.py by
    test_checker_agrees_on_topic_matching.
    """
    first = normalize_headline(one)
    second = normalize_headline(other)
    if first == second:
        return True
    if first and second and (first in second or
                             second in first):
        return True

    left, right = headline_tokens(one), headline_tokens(other)
    if not left or not right:
        return False
    shared = len(left & right)
    return shared / min(len(left), len(right)) >= 0.7


# --------------------------------------------------------------
def deleted_sidecar_path(md_path):
    """Build the deleted-topics path beside a deck file."""
    stem = os.path.splitext(md_path)[0]
    return stem + DELETED_SUFFIX


# --------------------------------------------------------------
def parse_deleted_topics(text):
    """List the headlines recorded in a sidecar file.

    Deliberately forgiving. The author pastes whole items
    in, so anything that is not a "###" heading is treated
    as a note and ignored.
    """
    topics = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("### "):
            topics.append(stripped[4:].strip())
    return topics


# --------------------------------------------------------------
def read_deleted_topics(md_path):
    """Read a deck's deleted headlines, or return []."""
    path = deleted_sidecar_path(md_path)
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as handle:
        return parse_deleted_topics(handle.read())


# --------------------------------------------------------------
def recreated_topics(deck, deleted):
    """Pair deck items with the deleted topic they match."""
    found = []
    for section in deck.sections:
        for block in section.blocks:
            for topic in deleted:
                if same_topic(block.headline, topic):
                    found.append((block.headline, topic))
                    break
    return found


# --------------------------------------------------------------
def live_sections(deck):
    """List the sections not struck out by a tombstone."""
    return [s for s in deck.sections if not s.deleted]


# --------------------------------------------------------------
def live_blocks(section):
    """List the blocks not struck out by a tombstone."""
    return [b for b in section.blocks if not b.deleted]


# --------------------------------------------------------------
def deleted_count(deck):
    """Count the tombstoned sections and blocks."""
    total = sum(1 for s in deck.sections if s.deleted)
    for section in deck.sections:
        total += sum(1 for b in section.blocks if b.deleted)
    return total


# --------------------------------------------------------------
def all_headlines(deck):
    """List the news headlines, for the contents page.

    Back matter such as the author page is marked notoc,
    because the contents lists the week's news and not the
    pages that close every deck.
    """
    names = []
    for section in live_sections(deck):
        if section.notoc:
            continue
        for block in live_blocks(section):
            if not block.notoc:
                names.append(block.headline)
    return names


# --------------------------------------------------------------
def image_manifest(deck):
    """List (path, url, mode) for every live image."""
    items = []
    for section in live_sections(deck):
        for block in live_blocks(section):
            if not block.image:
                continue
            if block.image_src:
                items.append(
                    (block.image, block.image_src, "src")
                )
            elif block.image_shot:
                items.append(
                    (block.image, block.image_shot, "shot")
                )
            else:
                items.append((block.image, None, "local"))
    return items
