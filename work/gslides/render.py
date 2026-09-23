#!/usr/bin/env python3
"""
Turn laid-out deck content into Google Slides requests.

The drawing half of the Slides pipeline. deck_layout decides
every rectangle; this module decides what goes in it and how
it is styled, and returns plain request dictionaries. It
never talks to Google, so every function here can be tested
by looking at what it returns.

Used by the g*_ scripts that build and refresh a live deck.

Two choices are passed in rather than fixed:

    fill    the background of a text box. Scripts pass None,
            because in a live deck a fill means "a human
            accepted this, hands off".

    locate  a function from a picture's local path to a URL
            Google can fetch, or None to leave the picture
            out. The scripts upload to Drive (gslides.host.Host).

Text offsets are counted in UTF-16 code units, which is what
the Slides API uses, so an emoji in a headline does not
shift every style that follows it.

Usage:
    from gslides.render import render_content, Body
    reqs = render_content(slide_id, slide_id, page, keys,
                          fill=None, locate=my_locator)

Created: 2026-09-14
Last updated: 2026-09-16
"""

import os
from dataclasses import dataclass, field

from layout import deck_layout as L
from gslides import api as S
from gslides import ids as T
from layout.deck_parser import is_link, split_markup
from layout.text_metrics import text_width

FONT = "Calibri"
CODE_FONT = "Consolas"
CODE_SIZE = 9

BLACK = (0x00, 0x00, 0x00)
RED = (0xFF, 0x00, 0x00)
BLUE = (0x05, 0x63, 0xC1)
CODE_BLUE = (0x3C, 0x78, 0xD8)
YELLOW = (0xFF, 0xF2, 0xCC)
GREY = (0x59, 0x59, 0x59)

# Real list bullets carry no text. BULLET stands in for the
# glyph when measuring a contents line, and is what decks
# written before 2026-09-16 have typed at a line's start.
BULLET = "● "
TOC_BLUE = (0x44, 0x72, 0xC4)
TOC_COLUMNS = 2
# The contents is set at 14 pt, and steps down to 12 pt when
# that is what it takes for every item to fit on the slide.
TOC_SIZE = 14
TOC_MIN_SIZE = 12
EPIGRAPH_SIZE = 18
EPIGRAPH_X = 5.90
TOC_GAP = 0.14
# The contents is four boxes, two a side, each its own tint:
# left upper, left lower, right upper, right lower.
LIGHT_GREEN = (0xE2, 0xF0, 0xD9)
LIGHT_BLUE = (0xDE, 0xEB, 0xF7)
TOC_FILLS = [YELLOW, LIGHT_GREEN, LIGHT_BLUE, YELLOW]
TOC_STACK_GAP = 0.10
# Each contents box holds 8 to 12 lines: as many as fit its
# side of the slide. An empty box shows TOC_EMPTY.
TOC_BOX_MIN = 8
TOC_BOX_MAX = 12
TOC_EMPTY = "xxx"

# Google Slides pads text away from the box edge by 0.10in
# left and right and 0.05in top and bottom, and the API has
# no field to change it. deck_layout sizes its boxes for the
# smaller PPTX padding, so every box is grown by the
# difference and moved out by half of it. The text then
# lands exactly where the layout intended and the wrapping
# matches what was measured. Probed on 2026-09-14.
PAD_X = 0.05
PAD_Y = 0.03

# A benchmark column is only as wide as its widest line, and
# that line is often the leaderboard URL. text_metrics is
# calibrated on words, so it comes up a few thousandths
# short on a string of slashes and dots, and the link wraps.
# This buys the column back that margin.
BENCH_SLACK = 0.06

PICTURE_PPI = 150

# Breathing room under the last line of a resized box.
TEXT_SLACK = 0.06


# --------------------------------------------------------------
def u16(text):
    """Length of text as the Slides API counts it."""
    return len(text.encode("utf-16-le")) // 2


@dataclass
class Body:
    """A box's text, plus the spans and lines to style."""

    text: str = ""
    spans: list = field(default_factory=list)
    lines: list = field(default_factory=list)

    # --------------------------------------
    def run(self, chunk, **style):
        """Append one styled piece of text."""
        start = u16(self.text)
        self.text += chunk
        if chunk.strip():
            self.spans.append((start, u16(self.text), style))

    # --------------------------------------
    def end_line(self, bullet=False):
        """Close a paragraph and remember its extent."""
        start = self.lines[-1][1] if self.lines else 0
        self.text += "\n"
        self.lines.append((start, u16(self.text), bullet))


# --------------------------------------------------------------
def piece_style(style, size):
    """Font settings for one piece of marked-up text."""
    if style == "code":
        return {"size": CODE_SIZE, "colour": CODE_BLUE,
                "font": CODE_FONT}
    if style == "red":
        return {"size": size, "bold": True, "colour": RED,
                "font": FONT}
    if style == "bold":
        return {"size": size, "bold": True,
                "colour": BLACK, "font": FONT}
    if style == "hilite":
        return {"size": size, "colour": BLACK, "font": FONT,
                "background": YELLOW}
    return {"size": size, "colour": BLACK, "font": FONT}


# --------------------------------------------------------------
def add_headline(body, item):
    """Write a block's bold red headline, if it has one."""
    text = item.block.headline
    if not text:
        return
    body.run(text, size=L.headline_size(item.block, item.size),
             bold=True, colour=RED, font=FONT)
    body.end_line()


# --------------------------------------------------------------
def add_bullet(body, text, item, dotted):
    """Write one body line, as a list bullet or plain."""
    if is_link(text):
        body.run(text, size=L.link_size(item.size),
                 colour=BLUE, font=FONT, link=text)
        body.end_line()
        return
    for chunk, style in split_markup(text):
        body.run(chunk, **piece_style(style, item.size))
    body.end_line(bullet=dotted)


# --------------------------------------------------------------
def block_body(item):
    """Build the full text and styling of one news block."""
    body = Body()
    add_headline(body, item)
    dotted = not L.is_promo(item.block)
    for text in item.block.bullets:
        add_bullet(body, text, item, dotted)
    return body


# --------------------------------------------------------------
def write_body(box_id, body, centred=False):
    """Emit the requests that fill and style one box."""
    text = body.text.rstrip("\n")
    if not text:
        return []
    reqs = S.insert_text(box_id, text)
    limit = u16(text)
    for start, end, style in body.spans:
        reqs += S.style_span(
            box_id, start, min(end, limit), **style
        )
    for start, end, bullet in body.lines:
        stop = min(end, limit)
        if bullet:
            reqs += S.bullets(box_id, start, stop)
            reqs += S.hanging_indent(box_id, start, stop)
        else:
            reqs += S.tighten(box_id, start, stop)
        if centred:
            reqs += S.centre(box_id, start, stop)
    return reqs


# --------------------------------------------------------------
def plain_body(text, size, colour=BLACK, bold=False,
               italic=False):
    """A body holding one uniformly styled paragraph set."""
    body = Body()
    for line in text.split("\n"):
        body.run(line, size=size, colour=colour, bold=bold,
                 italic=italic, font=FONT)
        body.end_line()
    return body


# --------------------------------------------------------------
def height_in_box(block, width, size):
    """How tall a block's text is at a width and size."""
    total = L.headline_height(block.headline, width,
                              L.headline_size(block, size))
    dotted = not L.is_promo(block)
    for text in block.bullets:
        total += L.bullet_height(text, width, size, dotted)
    return total


# --------------------------------------------------------------
def body_paragraphs(body):
    """(text, size, bold, bullet) for each line of a Body."""
    raw = body.text.encode("utf-16-le")
    out = []
    for start, end, bullet in body.lines:
        weight = {}
        for s_start, s_end, style in body.spans:
            overlap = min(end, s_end) - max(start, s_start)
            if overlap > 0:
                key = (style.get("size") or 12,
                       bool(style.get("bold")))
                weight[key] = weight.get(key, 0) + overlap
        size, bold = max(weight, key=weight.get) if weight \
            else (12, False)
        text = raw[2 * start:2 * end].decode("utf-16-le")
        out.append((text.rstrip("\n"), size, bold, bullet))
    return out


# --------------------------------------------------------------
def text_height(paragraphs, width):
    """Height of paragraphs at an inner width, with slack.

    paragraphs are (text, size, bold, bullet). A bulleted
    line wraps under its first word, so it has less width.
    """
    total = 0.0
    for text, size, bold, bullet in paragraphs:
        room = width - (L.BULLET_INSET if bullet else 0.0)
        lines = max(1, L.wrapped_lines(text, room, size, bold))
        total += lines * L.points_to_inches(size * L.LINE_RATIO)
    return total + TEXT_SLACK


# --------------------------------------------------------------
def box_height(paragraphs, box_width):
    """Height a box as read from the deck needs, padding in."""
    return text_height(paragraphs, box_width - 2 * PAD_X) \
        + 2 * PAD_Y


# --------------------------------------------------------------
def fit_rect(rect, body):
    """A layout rect made exactly as tall as its text."""
    return L.Rect(rect.x, rect.y, rect.w,
                  text_height(body_paragraphs(body), rect.w))


# --------------------------------------------------------------
def fit_size(block, rect, ceiling=None):
    """Largest ladder size at which a block fits a box.

    rect is the box as read back from the deck, so it
    includes the Slides padding, which is taken off first.
    """
    ceiling = ceiling or L.size_ceiling([block])
    width = rect.w - 2 * PAD_X - L.TEXT_INSET
    height = rect.h - 2 * PAD_Y
    sizes = L.ladder_below(ceiling)
    for size in sizes:
        if height_in_box(block, width, size) <= height:
            return size
    return sizes[-1]


# --------------------------------------------------------------
def make_box(box_id, slide_id, rect, border=False, fill=None):
    """Create a text box, padded for the Slides insets."""
    grown = L.Rect(rect.x - PAD_X, rect.y - PAD_Y,
                   rect.w + 2 * PAD_X, rect.h + 2 * PAD_Y)
    reqs = S.textbox(box_id, slide_id, grown)
    if border or fill:
        reqs += S.paint_box(box_id, fill=fill,
                            border=RED if border else None)
    return reqs


# --------------------------------------------------------------
def fitted(path, rect):
    """Shrink a rect to the picture's own aspect ratio."""
    try:
        from PIL import Image
        with Image.open(path) as img:
            src_w, src_h = img.size
    except Exception:
        return rect
    scale = min(rect.w / src_w, rect.h / src_h,
                1.0 / PICTURE_PPI)
    width, height = src_w * scale, src_h * scale
    return L.Rect(rect.x + (rect.w - width) / 2,
                  rect.y + (rect.h - height) / 2,
                  width, height)


# --------------------------------------------------------------
def render_picture(slide_id, key, path, rect, locate):
    """Place one picture with its red outline, if possible."""
    if not (path and rect and locate and os.path.isfile(path)):
        return []
    url = locate(path)
    if not url:
        return []
    art = T.shape_id(key, T.PICTURE)
    reqs = S.picture(art, slide_id, fitted(path, rect), url)
    return reqs + S.outline_picture(art, RED)


# --------------------------------------------------------------
def render_title(slide_id, tag, text, width=None):
    """Draw the slide title at the top left."""
    if not text:
        return []
    rect = L.Rect(L.MARGIN, L.TITLE_Y,
                  L.title_width(text, width), L.TITLE_H)
    box = T.shape_id(tag, T.TITLE)
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, u16(text),
                         size=L.TITLE_SIZE, bold=True,
                         colour=BLACK, font=FONT)
    return reqs


# --------------------------------------------------------------
def render_block(slide_id, key, item, boxed, fill=YELLOW,
                 locate=None):
    """Draw one news block: its text box and its picture."""
    box = T.shape_id(key, T.BOX)
    reqs = make_box(box, slide_id, item.text, boxed,
                    fill if boxed else None)
    reqs += write_body(box, block_body(item))
    reqs += render_picture(slide_id, key, item.block.image,
                           item.image, locate)
    return reqs


# --------------------------------------------------------------
def render_closing_title(slide_id, tag, text):
    """Draw the big centred title of the sign-off page."""
    if not text:
        return []
    rect = L.closing_title_rect()
    box = T.shape_id(tag, T.TITLE)
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, u16(text),
                         size=L.CLOSING_TITLE_SIZE,
                         bold=True, colour=BLACK, font=FONT)
    reqs += S.centre(box, 0, u16(text))
    return reqs


# --------------------------------------------------------------
def render_content(slide_id, tag, page, keys, fill=YELLOW,
                   locate=None):
    """Draw a title and every block placed on the page."""
    if page.closing:
        reqs = render_closing_title(slide_id, tag, page.title)
    else:
        reqs = render_title(slide_id, tag, page.title)
    for index, item in enumerate(page.items):
        reqs += render_block(slide_id, keys[index], item,
                             not page.plain, fill, locate)
    return reqs


# --------------------------------------------------------------
def toc_column_width(count=TOC_COLUMNS):
    """Width of one contents column."""
    return (L.CONTENT_W - TOC_GAP * (count - 1)) / max(1, count)


# --------------------------------------------------------------
def toc_fits(name, width=None):
    """Does one contents item, dot included, fit one line?"""
    width = width or toc_column_width()
    return L.wrapped_lines(BULLET + name, width, TOC_SIZE,
                           bold=True) <= 1


# --------------------------------------------------------------
def toc_columns(names, count=TOC_COLUMNS):
    """Split the headlines into equal top-down columns."""
    if not names:
        return []
    per = -(-len(names) // count)
    return [names[i:i + per]
            for i in range(0, len(names), per)]


# --------------------------------------------------------------
def epigraph_rect(text):
    """Where the week's epigraph sits, top right."""
    width = L.SLIDE_W - L.MARGIN - EPIGRAPH_X
    lines = max(1, L.wrapped_lines(text, width, EPIGRAPH_SIZE,
                                   bold=True))
    tall = lines * L.points_to_inches(
        EPIGRAPH_SIZE * L.LINE_RATIO) + 0.06
    return L.Rect(EPIGRAPH_X, 0.06, width, tall)


# --------------------------------------------------------------
def render_epigraph(slide_id, tag, text):
    """Draw the week's red epigraph, returning its bottom."""
    if not text:
        return [], L.BAND_TOP
    rect = epigraph_rect(text)
    box = T.shape_id(tag, "epi")
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, u16(text),
                         size=EPIGRAPH_SIZE, bold=True,
                         italic=True, colour=RED, font=FONT)
    return reqs, max(L.BAND_TOP, rect.y + rect.h + 0.10)


# --------------------------------------------------------------
def toc_column_rect(index, count, column, top, size=TOC_SIZE):
    """Where one contents column sits and how tall it is."""
    width = toc_column_width(count)
    tall = sum(
        L.wrapped_lines(BULLET + n, width, size, bold=True)
        for n in column
    ) * L.points_to_inches(size * L.LINE_RATIO)
    return L.Rect(L.MARGIN + index * (width + TOC_GAP), top,
                  width, min(L.BAND_H, tall + 0.06))


# --------------------------------------------------------------
def render_toc_columns(slide_id, tag, names, top, fill=YELLOW):
    """Draw the contents as bold blue bulleted columns."""
    columns = toc_columns(names)
    reqs = []
    style = {"size": TOC_SIZE, "colour": TOC_BLUE,
             "bold": True, "font": FONT}
    for index, column in enumerate(columns):
        body = Body()
        for name in column:
            body.run(name, **style)
            body.end_line(bullet=True)
        rect = toc_column_rect(index, len(columns), column,
                               top)
        box = T.shape_id(tag, f"c{index}")
        reqs += make_box(box, slide_id, rect, True, fill)
        reqs += write_body(box, body)
    return reqs


# --------------------------------------------------------------
def toc_capacity(top, boxes=2, size=TOC_SIZE):
    """Lines one box holds when a side starts at top."""
    line = L.points_to_inches(size * L.LINE_RATIO)
    room = (L.BAND_BOT - top - TOC_STACK_GAP * (boxes - 1))
    fits = int((room / boxes - 0.06) // line)
    return max(TOC_BOX_MIN, min(TOC_BOX_MAX, fits))


# --------------------------------------------------------------
def toc_capacities(tops, size=TOC_SIZE):
    """Lines each of the four boxes holds, c0 to c3."""
    return [toc_capacity(top, size=size)
            for top in tops for _ in (0, 1)]


# --------------------------------------------------------------
def toc_size(count, tops):
    """Largest font, 14 down to 12 pt, that fits every item.

    If even 12 pt is too big, 12 pt is used and the last box
    runs long, as before.
    """
    for size in range(TOC_SIZE, TOC_MIN_SIZE, -1):
        if sum(toc_capacities(tops, size)) >= count:
            return size
    return TOC_MIN_SIZE


# --------------------------------------------------------------
def toc_quarters(names, capacities):
    """Items in slide order, filling each box before the next.

    Boxes go left upper, left lower, right upper, right
    lower. Anything past the last box's share still goes in
    the last box, so no item is ever dropped.
    """
    parts, rest = [], list(names)
    for index, cap in enumerate(capacities):
        last = index == len(capacities) - 1
        parts.append(rest if last else rest[:cap])
        rest = [] if last else rest[cap:]
    return parts


# --------------------------------------------------------------
def toc_body(names, size=TOC_SIZE):
    """The bold blue bulleted lines of one contents box."""
    body = Body()
    for name in names:
        body.run(name, size=size, colour=TOC_BLUE,
                 bold=True, font=FONT)
        body.end_line(bullet=True)
    return body


# --------------------------------------------------------------
def render_toc_boxes(slide_id, tag, names, tops):
    """Draw the contents as four tinted boxes, two a side.

    tops is (left, right): where each side starts, so the
    right side can sit below the epigraph. Box c0 is left
    upper, c1 left lower, c2 right upper, c3 right lower,
    each filled with its colour from TOC_FILLS. Items run in
    slide order and fill each box (8 to 12 lines, as many as
    fit that side) before the next; an empty box says xxx.
    The font is the largest from 14 down to 12 pt at which
    every item fits on the slide.
    """
    reqs = []
    y = list(tops)
    size = toc_size(len(names), tops)
    caps = toc_capacities(tops, size)
    for index, part in enumerate(toc_quarters(names, caps)):
        part = part or [TOC_EMPTY]
        side = index // 2
        rect = toc_column_rect(side, TOC_COLUMNS, part, y[side],
                               size)
        y[side] = rect.y + rect.h + TOC_STACK_GAP
        box = T.shape_id(tag, f"c{index}")
        reqs += make_box(box, slide_id, rect, True,
                         TOC_FILLS[index])
        reqs += write_body(box, toc_body(part, size))
    return reqs


# --------------------------------------------------------------
def render_toc(slide_id, tag, page, names, epigraph,
               fill=YELLOW):
    """Draw the generated contents page."""
    limit = EPIGRAPH_X - L.MARGIN - 0.10 if epigraph else None
    reqs = render_title(slide_id, tag, page.title, limit)
    extra, top = render_epigraph(slide_id, tag, epigraph)
    reqs += extra
    return reqs + render_toc_columns(slide_id, tag, names,
                                     top, fill)


# --------------------------------------------------------------
def render_legend(slide_id, tag, B, fill=YELLOW):
    """Draw the vendor colour key on the header row."""
    body = Body()
    for vendor, label in B.VENDOR_LABELS:
        body.run(f"{B.MARKER} ", size=B.LEGEND_SIZE,
                 colour=B.VENDOR_COLORS[vendor],
                 font=FONT)
        body.run(f"{label}    ", size=B.LEGEND_SIZE,
                 colour=GREY, font=FONT)
    body.end_line()
    width = L.TEXT_INSET + sum(
        text_width(f"{B.MARKER} {label}    ", B.LEGEND_SIZE)
        for _, label in B.VENDOR_LABELS
    )
    rect = L.Rect(L.MARGIN, B.HEADER_Y, width,
                  B.header_height())
    box = T.shape_id(tag, "leg")
    reqs = make_box(box, slide_id, rect, True, fill)
    return reqs + write_body(box, body)


# --------------------------------------------------------------
def render_cutoff(slide_id, tag, boards, B, fill=YELLOW):
    """Draw the vote cutoff note beside the legend."""
    note = B.cutoff_note(boards)
    if not note:
        return []
    width = text_width(note, B.DATE_SIZE) + L.TEXT_INSET
    rect = L.Rect(L.SLIDE_W - L.MARGIN - width, B.HEADER_Y,
                  width, B.header_height())
    box = T.shape_id(tag, "date")
    reqs = make_box(box, slide_id, rect, True, fill)
    reqs += S.insert_text(box, note)
    reqs += S.style_span(box, 0, u16(note),
                         size=B.DATE_SIZE, colour=BLACK,
                         font=FONT)
    reqs += S.tighten(box, 0, u16(note))
    return reqs


# --------------------------------------------------------------
def bench_column_body(board, size, B):
    """Build the caption, link and rows of one column."""
    body = Body()
    body.run(board.get("label", ""),
             size=B.BENCH_CAPTION_SIZE, bold=True,
             colour=RED, font=FONT)
    body.end_line()
    url = board.get("url", "")
    if url:
        body.run(url, size=L.LINK_SIZE, colour=BLUE,
                 font=FONT, link=url)
        body.end_line()
    for entry in board.get("entries", [])[:B.BENCH_ROWS]:
        marker, score, name = B.row_text(entry)
        vendor = entry.get("vendor", "other")
        body.run(marker, size=size, font=FONT,
                 colour=B.VENDOR_COLORS[vendor])
        body.run(score, size=size, bold=True,
                 colour=BLACK, font=FONT)
        body.run(name, size=size, colour=BLACK, font=FONT)
        body.end_line()
    return body


# --------------------------------------------------------------
def bench_geometry(boards, B):
    """Font size and column rectangles for the columns."""
    height = L.BAND_BOT - B.bench_top()
    rows = max(len(b.get("entries", [])[:B.BENCH_ROWS])
               for b in boards)
    size = B.bench_font_size(
        rows, height - B.caption_height() - B.BOX_PAD
    )
    widths = [min((L.CONTENT_W - B.BENCH_GAP) / 2,
                  B.column_width(b, size) + BENCH_SLACK)
              for b in boards]
    total = sum(widths) + B.BENCH_GAP * (len(widths) - 1)
    left = (L.SLIDE_W - total) / 2
    rects = []
    for index, board in enumerate(boards):
        count = len(board.get("entries", [])[:B.BENCH_ROWS])
        rects.append(L.Rect(left, B.bench_top(), widths[index],
                            B.column_height(count, size)))
        left += widths[index] + B.BENCH_GAP
    return size, rects


# --------------------------------------------------------------
def render_bench(slide_id, tag, data, B, fill=YELLOW):
    """Draw the LM Arena page from the leaderboard JSON."""
    reqs = render_title(slide_id, tag, "Benchmarks")
    boards = (data or {}).get("boards", [])[:2]
    if not boards:
        return reqs
    reqs += render_legend(slide_id, tag, B, fill)
    reqs += render_cutoff(slide_id, tag, boards, B, fill)
    size, rects = bench_geometry(boards, B)
    for index, board in enumerate(boards):
        box = T.shape_id(tag, f"k{index}")
        reqs += make_box(box, slide_id, rects[index], True,
                         fill)
        reqs += write_body(
            box, bench_column_body(board, size, B)
        )
    return reqs
