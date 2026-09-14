#!/usr/bin/env python3
"""
Give every generated shape a stable Google Slides object id.

The id is how the automation recognizes its own work. A box
carrying an id from here was written by the script and may
be rewritten; anything else on the slide belongs to the
human and is left alone.

Object ids are the right place for this because Slides
enforces that they are unique across a whole presentation.
Copy a box and Google is forced to assign the copy a fresh
id, so a duplicate cannot impersonate the original. Alt text
does not have that property: it is copied verbatim, and
three boxes end up claiming the same topic.

An id is built from the topic's own words, so it survives
the human reordering slides, editing the text, or filling
the box with colour. It looks like:

    t-meta-ships-muse-agent-3f9c-b

    t       marks it as ours
    slug    words of the headline, cut to 32 characters
    3f9c    hash of the full headline, so two topics that
            shorten to the same slug stay apart
    b       which shape: b body, p picture, t title

Slides requires ids of 5 to 50 characters, made of letters,
digits, underscores and dashes.

Usage:
    from gslides.ids import assign_ids
    keys = assign_ids(pages)
    keys.page(0)          -> "s-toc"
    keys.block(2, 0)      -> "t-meta-ships-muse-3f9c"

Created: 2026-09-14
Last updated: 2026-09-14
"""

import hashlib
import re

from layout.deck_parser import (
    KIND_BENCHMARKS, KIND_TOC, strip_markup,
)

ID_MIN = 5
ID_MAX = 50
SLUG_MAX = 32
HASH_LEN = 4

OURS = "t"
PAGE_PREFIX = "s"

# Pages that are not a news topic get a fixed name, because
# there is only ever one of each in a deck.
FIXED_KEYS = {
    KIND_TOC: "toc",
    KIND_BENCHMARKS: "bench",
}

RE_NOT_WORD = re.compile(r"[^a-z0-9]+")

# Fixed ids of the skeleton deck g1_new_deck.py builds. They
# carry no hash, while every generated id ends in one, so the
# two families can never collide.
PAGE_TOC = "s-toc"
PAGE_BENCH = "s-bench"
PAGE_AA = "s-aa-index"
PAGE_NEWS_1 = "s-news-1"
PAGE_YOUTUBE = "s-youtube"
PAGE_NEWS_2 = "s-news-2"
PAGE_LAYOFFS = "s-layoffs"
PAGE_ABOUT = "s-about"
PAGE_THANKS = "s-thanks"
PAGE_SEPARATOR = "s-separator"
PAGE_PARKED = "s-parked"

TOPIC_AA = "t-aa-index"
TOPIC_NEWS_1 = "t-news-1"
TOPIC_NEWS_2 = "t-news-2"
TOPIC_YOUTUBE = "t-youtube"
TOPIC_LAYOFFS_FYI = "t-layoffs-fyi"
TOPIC_TRUEUP = "t-trueup"
TOPIC_ABOUT = "t-about"
TOPIC_THANKS = "t-thanks"

BOX = "b"
PICTURE = "p"
TITLE = "title"


# --------------------------------------------------------------
def slug(text, limit=SLUG_MAX):
    """Reduce a headline to lowercase words and dashes."""
    plain = strip_markup(text or "").lower()
    plain = RE_NOT_WORD.sub("-", plain).strip("-")
    if len(plain) <= limit:
        return plain
    # Cut on a word boundary so the id stays readable.
    cut = plain[:limit].rsplit("-", 1)[0]
    return cut or plain[:limit]


# --------------------------------------------------------------
def digest(text, length=HASH_LEN):
    """Short stable hash, so near-identical slugs differ."""
    raw = strip_markup(text or "").strip().lower()
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:length]


# --------------------------------------------------------------
def block_words(block):
    """The text that gives a news block its identity."""
    if block.headline:
        return block.headline
    for line in block.bullets:
        if line.strip():
            return line
    return block.image or ""


# --------------------------------------------------------------
def block_key(block):
    """Stable key for one news topic."""
    words = block_words(block)
    body = slug(words)
    tail = digest(words)
    key = f"{OURS}-{body}-{tail}" if body else f"{OURS}-{tail}"
    return key[:ID_MAX]


# --------------------------------------------------------------
def page_words(page):
    """The text that gives a whole page its identity."""
    if page.kind in FIXED_KEYS:
        return FIXED_KEYS[page.kind]
    if page.items:
        return block_words(page.items[0].block)
    return page.title or ""


# --------------------------------------------------------------
def page_key(page):
    """Stable key for one slide."""
    if page.kind in FIXED_KEYS:
        return f"{PAGE_PREFIX}-{FIXED_KEYS[page.kind]}"
    words = page_words(page)
    body = slug(words, SLUG_MAX - 2)
    tail = digest(words)
    key = f"{PAGE_PREFIX}-{body}-{tail}"
    return key[:ID_MAX]


# --------------------------------------------------------------
def make_unique(key, taken):
    """Add a counter until the key is not already used."""
    if key not in taken:
        taken.add(key)
        return key
    for number in range(2, 100):
        tail = f"-{number}"
        candidate = key[:ID_MAX - len(tail)] + tail
        if candidate not in taken:
            taken.add(candidate)
            return candidate
    raise ValueError(f"cannot make {key!r} unique")


# --------------------------------------------------------------
def valid(key):
    """Does Slides accept this as an object id?"""
    if not ID_MIN <= len(key) <= ID_MAX:
        return False
    return re.fullmatch(r"[A-Za-z0-9_][A-Za-z0-9_-]*",
                        key) is not None


class DeckIds:
    """Every object id for one deck, worked out up front."""

    # --------------------------------------
    def __init__(self, pages):
        """Assign a unique id to each page and block."""
        self.pages = []
        self.blocks = []
        taken = set()
        for page in pages:
            self.pages.append(
                make_unique(page_key(page), taken)
            )
            self.blocks.append([
                make_unique(block_key(item.block), taken)
                for item in page.items
            ])

    # --------------------------------------
    def page(self, index):
        """The slide id for one page."""
        return self.pages[index]

    # --------------------------------------
    def block(self, page_index, block_index):
        """The topic key for one block on a page."""
        return self.blocks[page_index][block_index]

    # --------------------------------------
    def every(self):
        """Every id in the deck, for checking."""
        out = list(self.pages)
        for row in self.blocks:
            out.extend(row)
        return out


# --------------------------------------------------------------
def assign_ids(pages):
    """Work out every object id for a paginated deck."""
    return DeckIds(pages)


# --------------------------------------------------------------
def shape_id(key, suffix):
    """Name one shape belonging to a page or a topic."""
    return f"{key}-{suffix}"[:ID_MAX]


# --------------------------------------------------------------
def owned(object_id):
    """Was this object made by a script?"""
    return object_id.startswith(
        (f"{OURS}-", f"{PAGE_PREFIX}-")
    )


# --------------------------------------------------------------
def box_of(picture_id):
    """The text box a picture belongs to, or None."""
    tail = f"-{PICTURE}"
    if not picture_id.endswith(tail):
        return None
    return picture_id[:-len(tail)] + f"-{BOX}"


# --------------------------------------------------------------
def topic_of(object_id):
    """The topic key inside a box or picture id, or None."""
    for suffix in (BOX, PICTURE):
        tail = f"-{suffix}"
        if object_id.startswith(f"{OURS}-") and \
                object_id.endswith(tail):
            return object_id[:-len(tail)]
    return None
