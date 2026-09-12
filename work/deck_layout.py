#!/usr/bin/env python3
"""
Lay out deck sections onto slides automatically.

Takes the Sections produced by deck_parser and decides,
with no human input, how many physical slides each one
needs, which news blocks land on which slide, how big the
text is, and the exact rectangle for every text box and
every picture.

The author never positions anything. A section may hold two
news items or twenty; this module splits it into as many
slides as the text needs, repeating the section title on
each continuation slide.

How it decides:

  1. Estimate the rendered height of every block from its
     character count and its box width.
  2. Fill a slide with blocks while they still fit the
     content band, up to MAX_BLOCKS per slide.
  3. A block too tall to share a slide gets one of its own,
     with the font stepped down until it fits.
  4. Leftover vertical space is shared out evenly between
     the blocks on a slide, so pages never look top-heavy.

All geometry is in inches. This module imports no
PowerPoint library and does no I/O, so its arithmetic can
be unit tested directly.

Usage:
    from deck_layout import paginate
    pages = paginate(deck.sections)
    for page in pages:
        print(page.title, len(page.items))

Created: 2026-09-12
Last updated: 2026-09-12
"""

from dataclasses import dataclass, field

import text_metrics as M

from deck_parser import GENERATED_KINDS, is_link, live_blocks

SLIDE_W = 10.00
SLIDE_H = 5.625
MARGIN = 0.09

TITLE_Y = 0.00
TITLE_H = 0.36
TITLE_SIZE = 20

BAND_TOP = 0.46
BAND_BOT = 5.53

CONTENT_W = SLIDE_W - 2 * MARGIN
IMG_W = 3.60
COL_GAP = 0.12
TEXT_W_IMG = CONTENT_W - IMG_W - COL_GAP
TEXT_W_FULL = CONTENT_W

# The author page: a large picture on the left with the
# details beside it, the shape used in the hand-made decks.
PROFILE_IMG_W = 1.72
PROFILE_HEADLINE_EXTRA = 8

MAX_BLOCKS = 3
BLOCK_GAP = 0.12

# Tried largest first, so a slide with room to spare uses
# bigger text. Packing still happens at BASE_SIZE, so a
# larger choice never changes how many slides there are.
FONT_LADDER = [16, 14, 12, 11, 10, 9, 8]
BASE_SIZE = 12
LINK_SIZE = 9
LINK_RATIO = 0.75
HEADLINE_EXTRA = 1

LINE_RATIO = 1.24

# A text box keeps 0.05 in of margin on each side, so the
# text itself gets this much less than the box width.
# pptx_text sets those margins; test_box_inset_agrees pins
# the two together.
TEXT_INSET = 0.10

# A bullet's hanging indent eats this much of the width.
BULLET_INSET = 171450 / 914400.0
PARA_GAP = 0.30

BAND_H = BAND_BOT - BAND_TOP


@dataclass
class Rect:
    """A rectangle on a slide, in inches."""

    x: float
    y: float
    w: float
    h: float


@dataclass
class Item:
    """One news block placed on a slide."""

    block: object
    size: int
    text: Rect
    image: Rect = None


@dataclass
class Page:
    """One physical slide ready to be rendered."""

    title: str
    kind: str
    items: list = field(default_factory=list)
    overflow: float = 0.0


# --------------------------------------------------------------
def points_to_inches(points):
    """Convert a point measurement to inches."""
    return points / 72.0


# --------------------------------------------------------------
def wrapped_lines(text, width_in, size_pt, bold=False):
    """Count the lines one string wraps to in a width."""
    return M.wrap_lines(
        text, width_in - TEXT_INSET, size_pt, bold
    )


# --------------------------------------------------------------
def link_size(size_pt):
    """Size for a link line under body text of a size."""
    return max(LINK_SIZE, round(size_pt * LINK_RATIO))


# --------------------------------------------------------------
def bullet_height(text, width_in, size_pt):
    """Estimate the height of one rendered bullet."""
    link = is_link(text)
    size = link_size(size_pt) if link else size_pt
    inset = 0.0 if link else BULLET_INSET
    lines = wrapped_lines(text, width_in - inset, size)
    return lines * points_to_inches(size * LINE_RATIO)


# --------------------------------------------------------------
def headline_height(text, width_in, size_pt):
    """Estimate the height of a bold block headline."""
    lines = wrapped_lines(text, width_in, size_pt, bold=True)
    return lines * points_to_inches(size_pt * LINE_RATIO)


# --------------------------------------------------------------
def is_profile(block):
    """Check for a picture-left profile block."""
    return bool(block.profile) and bool(block.image)


# --------------------------------------------------------------
def headline_size(block, size_pt):
    """Size of one block's bold headline."""
    if is_profile(block):
        return size_pt + PROFILE_HEADLINE_EXTRA
    return size_pt + HEADLINE_EXTRA


# --------------------------------------------------------------
def text_left(block):
    """Left edge of a block's text, in inches."""
    if is_profile(block):
        return MARGIN + PROFILE_IMG_W + COL_GAP
    return MARGIN


# --------------------------------------------------------------
def text_width_for(block):
    """Pick the text column width for one block."""
    if is_profile(block):
        return CONTENT_W - PROFILE_IMG_W - COL_GAP
    return TEXT_W_IMG if block.image else TEXT_W_FULL


# --------------------------------------------------------------
def block_height(block, size_pt):
    """Estimate the full rendered height of one block."""
    width = text_width_for(block)
    total = headline_height(
        block.headline, width, headline_size(block, size_pt)
    )
    for text in block.bullets:
        total += bullet_height(text, width, size_pt)
    total += points_to_inches(PARA_GAP * size_pt)
    return total


# --------------------------------------------------------------
def fitting_size(block):
    """Find the largest font size a lone block fits at."""
    for size in FONT_LADDER:
        if block_height(block, size) <= BAND_H:
            return size
    return FONT_LADDER[-1]


# --------------------------------------------------------------
def group_capacity(heights):
    """Total height a set of blocks needs, gaps included."""
    if not heights:
        return 0.0
    return sum(heights) + BLOCK_GAP * (len(heights) - 1)


# --------------------------------------------------------------
def pack_blocks(blocks):
    """Split blocks into groups, one group per slide."""
    groups = []
    current = []
    heights = []

    for block in blocks:
        height = block_height(block, BASE_SIZE)
        if height > BAND_H and not current:
            groups.append([block])
            continue
        trial = heights + [height]
        too_tall = group_capacity(trial) > BAND_H
        too_many = len(current) >= MAX_BLOCKS
        if current and (too_tall or too_many):
            groups.append(current)
            current, heights = [block], [height]
        else:
            current, heights = current + [block], trial

    if current:
        groups.append(current)
    return groups


# --------------------------------------------------------------
def group_size(group):
    """Pick one shared font size for a slide's blocks."""
    if len(group) == 1:
        return fitting_size(group[0])
    for size in FONT_LADDER:
        heights = [block_height(b, size) for b in group]
        if group_capacity(heights) <= BAND_H:
            return size
    return FONT_LADDER[-1]


# --------------------------------------------------------------
def image_rect(block, top, height):
    """Build the picture box for a block, or None."""
    if not block.image:
        return None
    if is_profile(block):
        # A square box at the top of the row, so the
        # portrait lines up with the first line of text
        # instead of floating halfway down the slide.
        return Rect(MARGIN, top, PROFILE_IMG_W, PROFILE_IMG_W)
    left = MARGIN + TEXT_W_IMG + COL_GAP
    return Rect(left, top, IMG_W, height)


# --------------------------------------------------------------
def place_group(group, size):
    """Turn one group of blocks into placed Items."""
    heights = [block_height(b, size) for b in group]
    slack = BAND_H - group_capacity(heights)
    extra = max(0.0, slack) / len(group)

    items = []
    top = BAND_TOP
    for block, height in zip(group, heights):
        row = height + extra
        # The text box is only as tall as its text, so the
        # border drawn around it hugs the words. The
        # picture still gets the whole row, which keeps it
        # as large as the slide can afford.
        items.append(Item(
            block=block,
            size=size,
            text=Rect(
                text_left(block), top,
                text_width_for(block), height
            ),
            image=image_rect(block, top, row),
        ))
        top += row + BLOCK_GAP
    return items


# --------------------------------------------------------------
def group_overflow(group, size):
    """Height by which a group misses the content band."""
    heights = [block_height(b, size) for b in group]
    return max(0.0, group_capacity(heights) - BAND_H)


# --------------------------------------------------------------
def paginate_section(section):
    """Build every Page needed by one section."""
    if section.kind in GENERATED_KINDS:
        return [Page(title=section.title, kind=section.kind)]

    pages = []
    for group in pack_blocks(live_blocks(section)):
        size = group_size(group)
        pages.append(Page(
            title=section.title,
            kind=section.kind,
            items=place_group(group, size),
            overflow=group_overflow(group, size),
        ))
    return pages


# --------------------------------------------------------------
def paginate(sections):
    """Lay out every live section onto physical slides."""
    pages = []
    for section in sections:
        if section.deleted:
            continue
        pages.extend(paginate_section(section))
    return pages
