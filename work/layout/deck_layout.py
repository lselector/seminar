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
    from layout.deck_layout import paginate
    pages = paginate(deck.sections)
    for page in pages:
        print(page.title, len(page.items))

Created: 2026-09-12
Last updated: 2026-09-12
"""

from dataclasses import dataclass, field

from layout import text_metrics as M

from layout.deck_parser import (
    GENERATED_KINDS, is_link, live_blocks, strip_markup,
)

SLIDE_W = 10.00
SLIDE_H = 5.625
MARGIN = 0.09

TITLE_Y = 0.00
TITLE_H = 0.36
TITLE_SIZE = 20
TITLE_SLACK = 1.10

BAND_TOP = 0.46
BAND_BOT = 5.53

CONTENT_W = SLIDE_W - 2 * MARGIN
IMG_W = 3.60
COL_GAP = 0.12
TEXT_W_IMG = CONTENT_W - IMG_W - COL_GAP
TEXT_W_FULL = CONTENT_W

# The author page: a large picture on the left with the
# details beside it, the shape used in the hand-made decks.
# The pair is centred on the slide rather than stretched
# across it, so the text column is a fixed width instead of
# "whatever is left over".
PROFILE_IMG_W = 1.72
PROFILE_GAP = 0.56
PROFILE_TEXT_W = 4.90
PROFILE_HEADLINE_EXTRA = 8

# The sign-off page: a big centred title a third of the way
# down, with the links centred beneath it and no box drawn
# around either.
CLOSING_TITLE_SIZE = 40
CLOSING_TITLE_Y = SLIDE_H / 3
CLOSING_BODY_GAP = 0.16

MAX_BLOCKS = 3

# White space above, between and below the text boxes on a
# slide. The free height is shared out evenly, so a box
# never sits flush against the title, but the share is
# capped: on a sparse slide the leftover goes to the
# pictures rather than floating one small box mid-page.
BLOCK_GAP = 0.12
MAX_BLOCK_GAP = 0.40

# Tried largest first, never above the ceiling its page
# kind allows. News slides stay at BASE_SIZE, the body size
# the decks have always used. The three standing pages hold
# one short block each and are allowed to go bigger.
FONT_LADDER = [22, 18, 16, 14, 12, 11, 10, 9, 8]
BASE_SIZE = 12
SPECIAL_CEILING = 16
LINK_SIZE = 9
LINK_RATIO = 0.75

# The channel page: a large call to action on the left with
# its screenshots beside it, and no bullet dots.
PROMO_SIZE = 22
PROMO_LINK_SIZE = 18
PROMO_TEXT_W = 4.60
PROMO_IMG_W = CONTENT_W - PROMO_TEXT_W - COL_GAP

# The widest a picture can ever be drawn. s2_clean_images
# scales to this, so no picture is stored larger than any
# box can show, and none is too small for its box.
MAX_IMG_W = max(IMG_W, PROMO_IMG_W)
HEADLINE_EXTRA = 1

LINE_RATIO = 1.24

# The layout plans for 0.05 in of text margin on each side,
# so the text gets this much less than the box width.
# gslides.render grows each box to make up the difference
# with Google Slides' larger fixed padding (PAD_X, PAD_Y).
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
    closing: bool = False
    plain: bool = False


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
def title_width(text, limit=None):
    """Width a slide title needs, never more than allowed.

    The measuring face is Arial, not Calibri, and the bold
    weights differ a little more than the regular ones, so
    the box gets a small margin. The title box also has
    wrapping switched off, which makes an underestimate
    harmless rather than a broken line.
    """
    cap = limit if limit else CONTENT_W
    if not text:
        return cap
    natural = M.text_width(text, TITLE_SIZE, bold=True)
    return min(cap, natural * TITLE_SLACK + TEXT_INSET)


# --------------------------------------------------------------
def link_size(size_pt):
    """Size for a link line under body text of a size."""
    if size_pt >= PROMO_SIZE:
        return PROMO_LINK_SIZE
    return max(LINK_SIZE, round(size_pt * LINK_RATIO))


# --------------------------------------------------------------
def bullet_height(text, width_in, size_pt,
                  bullets=True):
    """Estimate the height of one rendered bullet."""
    link = is_link(text)
    size = link_size(size_pt) if link else size_pt
    inset = 0.0 if link or not bullets else BULLET_INSET
    lines = wrapped_lines(
        strip_markup(text), width_in - inset, size
    )
    return lines * points_to_inches(size * LINE_RATIO)


# --------------------------------------------------------------
def headline_height(text, width_in, size_pt):
    """Estimate the height of a bold block headline."""
    if not text:
        return 0.0
    lines = wrapped_lines(text, width_in, size_pt, bold=True)
    return lines * points_to_inches(size_pt * LINE_RATIO)


# --------------------------------------------------------------
def is_closing(block):
    """Check for a block on the sign-off page."""
    return bool(block.closing)


# --------------------------------------------------------------
def closing_title_height():
    """Height the big centred title takes up."""
    return points_to_inches(
        CLOSING_TITLE_SIZE * LINE_RATIO
    ) + TEXT_INSET


# --------------------------------------------------------------
def closing_title_rect():
    """Where the sign-off title sits, spanning the slide."""
    return Rect(
        MARGIN, CLOSING_TITLE_Y,
        CONTENT_W, closing_title_height()
    )


# --------------------------------------------------------------
def closing_body_top():
    """Top of the links, just under the sign-off title."""
    return (
        CLOSING_TITLE_Y + closing_title_height()
        + CLOSING_BODY_GAP
    )


# --------------------------------------------------------------
def natural_width(block, size_pt):
    """Box width the block's longest line would need."""
    widths = [M.text_width(
        block.headline, headline_size(block, size_pt),
        bold=True
    )]
    for text in block.bullets:
        link = is_link(text)
        size = link_size(size_pt) if link else size_pt
        pad = 0.0 if link else BULLET_INSET
        widths.append(
            M.text_width(strip_markup(text), size) + pad
        )
    return min(CONTENT_W, max(widths) + TEXT_INSET)


# --------------------------------------------------------------
def is_promo(block):
    """Check for a large call-to-action block."""
    return bool(block.promo)


# --------------------------------------------------------------
def size_ceiling(group):
    """Largest font size a group of blocks may use."""
    if any(is_promo(b) for b in group):
        return PROMO_SIZE
    return BASE_SIZE


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
    """Left edge of a block's text, in inches.

    Profile blocks are positioned by place_profile, which
    centres the picture and the text as one pair.
    """
    return MARGIN


# --------------------------------------------------------------
def text_width_for(block):
    """Pick the text column width for one block."""
    if is_profile(block):
        return PROFILE_TEXT_W
    if is_promo(block):
        return PROMO_TEXT_W
    return TEXT_W_IMG if block.image else TEXT_W_FULL


# --------------------------------------------------------------
def image_width_for(block):
    """Widest box this block's picture can be drawn in."""
    if is_promo(block):
        return PROMO_IMG_W
    if is_profile(block):
        return PROFILE_IMG_W
    return IMG_W


# --------------------------------------------------------------
def image_widths(sections):
    """Map each picture's file stem to its widest box.

    s2_clean_images uses this so no picture is stored with
    more pixels than the one box it appears in can show.
    """
    widths = {}
    for section in sections:
        if section.deleted:
            continue
        for block in live_blocks(section):
            if not block.image:
                continue
            name = block.image.split("/")[-1]
            stem = name.rsplit(".", 1)[0]
            widths[stem] = max(
                widths.get(stem, 0.0), image_width_for(block)
            )
    return widths


# --------------------------------------------------------------
def block_height(block, size_pt):
    """Estimate the full rendered height of one block."""
    width = text_width_for(block)
    total = headline_height(
        block.headline, width, headline_size(block, size_pt)
    )
    dotted = not is_promo(block)
    for text in block.bullets:
        total += bullet_height(text, width, size_pt, dotted)
    total += points_to_inches(PARA_GAP * size_pt)
    return total


# --------------------------------------------------------------
def ladder_below(ceiling):
    """Sizes to try, largest first, none above a ceiling."""
    return [s for s in FONT_LADDER if s <= ceiling]


# --------------------------------------------------------------
def fitting_size(block, ceiling=BASE_SIZE):
    """Find the largest font size a lone block fits at."""
    sizes = ladder_below(ceiling)
    for size in sizes:
        if block_height(block, size) <= BAND_H:
            return size
    return sizes[-1]


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
def group_size(group, ceiling=None):
    """Pick one shared font size for a slide's blocks."""
    if ceiling is None:
        ceiling = size_ceiling(group)
    if len(group) == 1:
        return fitting_size(group[0], ceiling)
    sizes = ladder_below(ceiling)
    for size in sizes:
        heights = [block_height(b, size) for b in group]
        if group_capacity(heights) <= BAND_H:
            return size
    return sizes[-1]


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
    if is_promo(block):
        left = MARGIN + PROMO_TEXT_W + COL_GAP
        return Rect(left, top, PROMO_IMG_W, height)
    left = MARGIN + TEXT_W_IMG + COL_GAP
    return Rect(left, top, IMG_W, height)


# --------------------------------------------------------------
def block_gap(free, count):
    """Even white space around a slide's text boxes."""
    if count <= 0:
        return BLOCK_GAP
    even = free / (count + 1)
    return max(BLOCK_GAP, min(even, MAX_BLOCK_GAP))


# --------------------------------------------------------------
def place_group(group, size):
    """Turn one group of blocks into placed Items."""
    heights = [block_height(b, size) for b in group]
    free = BAND_H - sum(heights)
    gap = block_gap(free, len(group))
    leftover = max(0.0, free - gap * (len(group) + 1))
    share = leftover / len(group)

    items = []
    top = BAND_TOP + gap
    for block, height in zip(group, heights):
        row = height + share
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
        top += row + gap
    return items


# --------------------------------------------------------------
def closing_size(group):
    """Font size that fits under the sign-off title."""
    room = BAND_BOT - closing_body_top()
    for size in ladder_below(SPECIAL_CEILING):
        heights = [block_height(b, size) for b in group]
        if group_capacity(heights) <= room:
            return size
    return FONT_LADDER[-1]


# --------------------------------------------------------------
def place_closing(group, size):
    """Centre each sign-off block under the title."""
    items = []
    top = closing_body_top()
    for block in group:
        width = natural_width(block, size)
        height = block_height(block, size)
        items.append(Item(
            block=block,
            size=size,
            text=Rect(
                (SLIDE_W - width) / 2, top, width, height
            ),
        ))
        top += height + BLOCK_GAP
    return items


# --------------------------------------------------------------
def profile_text_width(block, size):
    """Width of the details column beside a portrait."""
    return min(PROFILE_TEXT_W, natural_width(block, size))


# --------------------------------------------------------------
def place_profile(group, size):
    """Centre a portrait and its details as one pair."""
    items = []
    for block in group:
        text_w = profile_text_width(block, size)
        text_h = block_height(block, size)
        pair_w = PROFILE_IMG_W + PROFILE_GAP + text_w
        pair_h = max(PROFILE_IMG_W, text_h)
        left = (SLIDE_W - pair_w) / 2
        top = BAND_TOP + (BAND_H - pair_h) / 2
        items.append(Item(
            block=block,
            size=size,
            text=Rect(
                left + PROFILE_IMG_W + PROFILE_GAP,
                top + (pair_h - text_h) / 2,
                text_w, text_h,
            ),
            image=Rect(
                left, top + (pair_h - PROFILE_IMG_W) / 2,
                PROFILE_IMG_W, PROFILE_IMG_W,
            ),
        ))
    return items


# --------------------------------------------------------------
def paginate_profile(section):
    """Build the single page an author section needs."""
    blocks = live_blocks(section)
    size = group_size(blocks, SPECIAL_CEILING)
    return [Page(
        title=section.title,
        kind=section.kind,
        items=place_profile(blocks, size),
        plain=True,
    )]


# --------------------------------------------------------------
def paginate_closing(section):
    """Build the single page a sign-off section needs."""
    blocks = live_blocks(section)
    size = closing_size(blocks)
    return [Page(
        title=section.title,
        kind=section.kind,
        items=place_closing(blocks, size),
        closing=True,
        plain=True,
    )]


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
    if section.closing and live_blocks(section):
        return paginate_closing(section)
    if any(is_profile(b) for b in live_blocks(section)):
        return paginate_profile(section)

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
