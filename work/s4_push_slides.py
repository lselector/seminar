#!/usr/bin/env python3
"""
Build a Google Slides deck from the weekly Markdown file.

The same deck, rendered somewhere else. deck_parser reads
the Markdown and deck_layout decides every rectangle, both
untouched: a Slides page is 10 x 5.625 inches, exactly the
PPTX page, so the geometry needs no translation at all.
Only the drawing differs, and that lives in slides_api.py.

Pictures are pulled from the public GitHub copy of this
repository. Google fetches image URLs itself and cannot see
your Drive, so the files have to be reachable without
credentials, and work/images is already published there.

The deck is created inside a folder you name. Under the
drive.file scope this script can only ever see decks it
created itself, so it can never touch the rest of your
Drive.

Usage:
    python3 s4_push_slides.py <FOLDER_ID>
    python3 s4_push_slides.py <FOLDER_ID> <deck.md>

Run g_auth.py once first to create credentials/token.json.

Created: 2026-09-14
Last updated: 2026-09-14
"""

import glob
import json
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime

import deck_layout as L
import slides_api as S
import topic_ids as T
from deck_parser import (
    KIND_BENCHMARKS, KIND_TOC, all_headlines, is_link,
    parse_deck_file, split_markup,
)

CRED_DIR = "credentials"
TOKEN_FILE = os.path.join(CRED_DIR, "token.json")
DECK_GLOB = "*-AI-News.md"

RAW_BASE = ("https://raw.githubusercontent.com/lselector/"
            "seminar/master/work/")

FONT = "Calibri"
CODE_FONT = "Consolas"
CODE_SIZE = 9

BLACK = (0x00, 0x00, 0x00)
RED = (0xFF, 0x00, 0x00)
BLUE = (0x05, 0x63, 0xC1)
CODE_BLUE = (0x3C, 0x78, 0xD8)
YELLOW = (0xFF, 0xF2, 0xCC)
GREY = (0x59, 0x59, 0x59)

BULLET = "● "
TOC_SIZE = 14
EPIGRAPH_SIZE = 18

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


@dataclass
class Body:
    """A box's text, plus the spans and lines to style."""

    text: str = ""
    spans: list = field(default_factory=list)
    lines: list = field(default_factory=list)

    # --------------------------------------
    def run(self, chunk, **style):
        """Append one styled piece of text."""
        start = len(self.text)
        self.text += chunk
        if chunk.strip():
            self.spans.append((start, len(self.text), style))

    # --------------------------------------
    def end_line(self, bullet=False):
        """Close a paragraph and remember its extent."""
        start = self.lines[-1][1] if self.lines else 0
        self.text += "\n"
        self.lines.append((start, len(self.text), bullet))


# --------------------------------------------------------------
def log(message):
    """Print one timestamped line."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}")


# --------------------------------------------------------------
def credentials():
    """Load the token g_auth.py saved, or explain."""
    if not os.path.isfile(TOKEN_FILE):
        log(f"ERROR: {TOKEN_FILE} not found.")
        log("  Run: python3 g_auth.py <FOLDER_ID>")
        sys.exit(1)
    from google.oauth2.credentials import Credentials
    scopes = json.load(open(TOKEN_FILE))["scopes"]
    return Credentials.from_authorized_user_file(
        TOKEN_FILE, scopes
    )


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
    """Write one body line, with or without its dot."""
    if is_link(text):
        body.run(text, size=L.link_size(item.size),
                 colour=BLUE, font=FONT, link=text)
        body.end_line()
        return
    if dotted:
        body.run(BULLET, size=item.size, colour=RED,
                 font=FONT)
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
    limit = len(text)
    for start, end, style in body.spans:
        reqs += S.style_span(
            box_id, start, min(end, limit), **style
        )
    for start, end, bullet in body.lines:
        stop = min(end, limit)
        if bullet:
            reqs += S.hanging_indent(box_id, start, stop)
        else:
            reqs += S.tighten(box_id, start, stop)
        if centred:
            reqs += S.centre(box_id, start, stop)
    return reqs


# --------------------------------------------------------------
def make_box(box_id, slide_id, rect, boxed=False):
    """Create a text box, padded for the Slides insets."""
    grown = L.Rect(rect.x - PAD_X, rect.y - PAD_Y,
                   rect.w + 2 * PAD_X, rect.h + 2 * PAD_Y)
    reqs = S.textbox(box_id, slide_id, grown)
    if boxed:
        reqs += S.paint_box(box_id, fill=YELLOW, border=RED)
    return reqs


# --------------------------------------------------------------
def image_url(path):
    """Public URL Google can fetch a picture from."""
    return RAW_BASE + path.lstrip("./")


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
                1.0 / 150)
    width, height = src_w * scale, src_h * scale
    return L.Rect(rect.x + (rect.w - width) / 2,
                  rect.y + (rect.h - height) / 2,
                  width, height)


# --------------------------------------------------------------
def render_title(slide_id, tag, text, width=None):
    """Draw the slide title at the top left."""
    if not text:
        return []
    rect = L.Rect(L.MARGIN, L.TITLE_Y,
                  L.title_width(text, width), L.TITLE_H)
    box = T.shape_id(tag, "title")
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, len(text),
                         size=L.TITLE_SIZE, bold=True,
                         colour=BLACK, font=FONT)
    return reqs


# --------------------------------------------------------------
def render_block(slide_id, key, item, boxed):
    """Draw one news block: its text box and its picture."""
    box = T.shape_id(key, "b")
    reqs = make_box(box, slide_id, item.text, boxed)
    reqs += write_body(box, block_body(item))

    path = item.block.image
    if path and item.image and os.path.isfile(path):
        art = T.shape_id(key, "p")
        reqs += S.picture(art, slide_id,
                          fitted(path, item.image),
                          image_url(path))
        reqs += S.outline_picture(art, RED)
    return reqs


# --------------------------------------------------------------
def render_closing_title(slide_id, tag, text):
    """Draw the big centred title of the sign-off page."""
    if not text:
        return []
    rect = L.closing_title_rect()
    box = T.shape_id(tag, "title")
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, len(text),
                         size=L.CLOSING_TITLE_SIZE,
                         bold=True, colour=BLACK, font=FONT)
    reqs += S.centre(box, 0, len(text))
    return reqs


# --------------------------------------------------------------
def render_content(slide_id, tag, page, keys):
    """Draw a title and every block placed on the page."""
    if page.closing:
        reqs = render_closing_title(slide_id, tag, page.title)
    else:
        reqs = render_title(slide_id, tag, page.title)
    for index, item in enumerate(page.items):
        reqs += render_block(slide_id, keys[index], item,
                             not page.plain)
    return reqs


# --------------------------------------------------------------
def toc_columns(names, count=3):
    """Split the headlines into equal top-down columns."""
    if not names:
        return []
    per = -(-len(names) // count)
    return [names[i:i + per]
            for i in range(0, len(names), per)]


# --------------------------------------------------------------
def render_epigraph(slide_id, tag, text):
    """Draw the week's red epigraph, returning its bottom."""
    if not text:
        return [], L.BAND_TOP
    width = L.SLIDE_W - L.MARGIN - 5.90
    lines = L.wrapped_lines(text, width, EPIGRAPH_SIZE,
                            bold=True)
    tall = lines * L.points_to_inches(
        EPIGRAPH_SIZE * L.LINE_RATIO) + 0.06
    rect = L.Rect(5.90, 0.06, width, tall)
    box = T.shape_id(tag, "epi")
    reqs = make_box(box, slide_id, rect)
    reqs += S.insert_text(box, text)
    reqs += S.style_span(box, 0, len(text),
                         size=EPIGRAPH_SIZE, bold=True,
                         italic=True, colour=RED, font=FONT)
    return reqs, max(L.BAND_TOP, rect.y + rect.h + 0.10)


# --------------------------------------------------------------
def render_toc(slide_id, tag, page, names, epigraph):
    """Draw the generated contents page."""
    limit = 5.90 - L.MARGIN - 0.10 if epigraph else None
    reqs = render_title(slide_id, tag, page.title, limit)
    extra, top = render_epigraph(slide_id, tag, epigraph)
    reqs += extra

    columns = toc_columns(names)
    gap = 0.14
    width = (L.CONTENT_W - gap * (len(columns) - 1))
    width = width / max(1, len(columns))
    for index, column in enumerate(columns):
        body = Body()
        for name in column:
            body.run(name, size=TOC_SIZE, colour=BLACK,
                     font=FONT)
            body.end_line()
        tall = sum(
            L.wrapped_lines(n, width, TOC_SIZE)
            for n in column
        ) * L.points_to_inches(TOC_SIZE * L.LINE_RATIO)
        rect = L.Rect(L.MARGIN + index * (width + gap), top,
                      width, min(L.BAND_H, tall + 0.06))
        box = T.shape_id(tag, f"c{index}")
        reqs += make_box(box, slide_id, rect, True)
        reqs += write_body(box, body)
    return reqs


# --------------------------------------------------------------
def as_rgb(colour):
    """Turn a python-pptx RGBColor into a byte triple."""
    text = str(colour)
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


# --------------------------------------------------------------
def leaderboard():
    """Read the JSON leaderboard.py writes, or None."""
    path = os.path.join("resources_raw", "leaderboard.json")
    if not os.path.isfile(path):
        log(f"WARNING: {path} not found. Run leaderboard.py")
        return None
    with open(path, encoding="utf-8") as handle:
        return json.load(handle)


# --------------------------------------------------------------
def render_legend(slide_id, tag, B):
    """Draw the vendor colour key on the header row."""
    body = Body()
    for vendor, label in B.VENDOR_LABELS:
        body.run(f"{B.MARKER} ", size=B.LEGEND_SIZE,
                 colour=as_rgb(B.VENDOR_COLORS[vendor]),
                 font=FONT)
        body.run(f"{label}    ", size=B.LEGEND_SIZE,
                 colour=GREY, font=FONT)
    body.end_line()
    from text_metrics import text_width
    width = L.TEXT_INSET + sum(
        text_width(f"{B.MARKER} {label}    ", B.LEGEND_SIZE)
        for _, label in B.VENDOR_LABELS
    )
    rect = L.Rect(L.MARGIN, B.HEADER_Y, width,
                  B.header_height())
    box = T.shape_id(tag, "leg")
    reqs = make_box(box, slide_id, rect, True)
    reqs += write_body(box, body)
    return reqs


# --------------------------------------------------------------
def render_cutoff(slide_id, tag, boards, B):
    """Draw the vote cutoff note beside the legend."""
    note = B.cutoff_note(boards)
    if not note:
        return []
    from text_metrics import text_width
    width = text_width(note, B.DATE_SIZE) + L.TEXT_INSET
    rect = L.Rect(L.SLIDE_W - L.MARGIN - width, B.HEADER_Y,
                  width, B.header_height())
    box = T.shape_id(tag, "date")
    reqs = make_box(box, slide_id, rect, True)
    reqs += S.insert_text(box, note)
    reqs += S.style_span(box, 0, len(note),
                         size=B.DATE_SIZE, colour=BLACK,
                         font=FONT)
    reqs += S.tighten(box, 0, len(note))
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
                 colour=as_rgb(B.VENDOR_COLORS[vendor]))
        body.run(score, size=size, bold=True,
                 colour=BLACK, font=FONT)
        body.run(name, size=size, colour=BLACK, font=FONT)
        body.end_line()
    return body


# --------------------------------------------------------------
def render_bench(slide_id, tag, data, B):
    """Draw the LM Arena page from the leaderboard JSON."""
    reqs = render_title(slide_id, tag, "Benchmarks")
    if not data:
        return reqs
    boards = data.get("boards", [])[:2]
    if not boards:
        return reqs
    reqs += render_legend(slide_id, tag, B)
    reqs += render_cutoff(slide_id, tag, boards, B)

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
    for index, board in enumerate(boards):
        rect = L.Rect(left, B.bench_top(), widths[index],
                      B.column_height(
                          len(board.get("entries", [])
                              [:B.BENCH_ROWS]), size))
        box = T.shape_id(tag, f"k{index}")
        reqs += make_box(box, slide_id, rect, True)
        reqs += write_body(
            box, bench_column_body(board, size, B)
        )
        left += widths[index] + B.BENCH_GAP
    return reqs


# --------------------------------------------------------------
def page_requests(page, index, names, epigraph, bench, B,
                  ids):
    """Build every request for one page of the deck."""
    # The slide id doubles as the prefix for the shapes that
    # belong to the page rather than to a topic.
    slide_id = ids.page(index)
    reqs = list(S.new_slide(slide_id))
    reqs[0]["createSlide"]["insertionIndex"] = index
    if page.kind == KIND_TOC:
        return reqs + render_toc(slide_id, slide_id, page,
                                 names, epigraph)
    if page.kind == KIND_BENCHMARKS:
        return reqs + render_bench(slide_id, slide_id,
                                   bench, B)
    return reqs + render_content(slide_id, slide_id, page,
                                 ids.blocks[index])


# --------------------------------------------------------------
def check_ids(ids):
    """Stop before the API does if an id is unusable."""
    every = ids.every()
    bad = [k for k in every if not T.valid(k)]
    if bad:
        log(f"ERROR: unusable object ids: {bad[:5]}")
        sys.exit(1)
    if len(set(every)) != len(every):
        log("ERROR: two topics claim the same object id")
        sys.exit(1)


# --------------------------------------------------------------
def deck_name(source):
    """Name the presentation after the deck file."""
    stem = os.path.basename(source)
    stem = os.path.splitext(stem)[0]
    return f"{stem}-Template"


# --------------------------------------------------------------
def newest_deck():
    """Find the newest deck Markdown in this directory."""
    names = sorted(glob.glob(DECK_GLOB))
    return names[-1] if names else None


# --------------------------------------------------------------
def resolve(argv):
    """Read the folder id and deck path from the command."""
    named = [a for a in argv[1:] if not a.startswith("-")]
    if not named:
        log("Usage: python3 s4_push_slides.py <FOLDER_ID>")
        sys.exit(1)
    folder = named[0]
    source = named[1] if len(named) > 1 else newest_deck()
    if not source or not os.path.isfile(source):
        log(f"ERROR: deck file not found: {source}")
        sys.exit(1)
    return folder, source


# --------------------------------------------------------------
def create_deck(drive, name, folder):
    """Create an empty presentation inside a folder."""
    made = drive.files().create(body={
        "name": name,
        "mimeType": ("application/vnd.google-apps"
                     ".presentation"),
        "parents": [folder],
    }, fields="id, webViewLink").execute()
    return made["id"], made.get("webViewLink", "")


# --------------------------------------------------------------
def send(slides, deck_id, requests, label):
    """Send one batch of requests, reporting failure."""
    from googleapiclient.errors import HttpError
    if not requests:
        return True
    try:
        slides.presentations().batchUpdate(
            presentationId=deck_id,
            body={"requests": requests},
        ).execute()
        return True
    except HttpError as exc:
        log(f"  FAILED {label}: {str(exc)[:220]}")
        return False


# --------------------------------------------------------------
def build_slides(slides, deck_id, pages, parts, B, ids):
    """Send one batch per page, then drop the blank slide."""
    names, epigraph, bench = parts
    ok = 0
    for index, page in enumerate(pages):
        label = f"slide {index + 1} ({page.kind})"
        if send(slides, deck_id,
                page_requests(page, index, names, epigraph,
                              bench, B, ids), label):
            ok += 1
            log(f"  {label}  {ids.page(index)}")
    deck = slides.presentations().get(
        presentationId=deck_id
    ).execute()
    mine = set(ids.pages)
    extra = [s["objectId"] for s in deck["slides"]
             if s["objectId"] not in mine]
    for object_id in extra:
        send(slides, deck_id, S.delete_object(object_id),
             "remove starting slide")
    return ok


# --------------------------------------------------------------
def main():
    """Render the weekly Markdown into a Slides deck."""
    import bench_page as B
    from googleapiclient.discovery import build

    folder, source = resolve(sys.argv)
    creds = credentials()
    drive = build("drive", "v3", credentials=creds)
    slides = build("slides", "v1", credentials=creds)

    log(f"Reading {source}")
    deck = parse_deck_file(source)
    pages = L.paginate(deck.sections)
    for page in pages:
        if page.kind == KIND_TOC and not page.title:
            page.title = deck.title
    parts = (all_headlines(deck), deck.epigraph,
             leaderboard())
    ids = T.assign_ids(pages)
    check_ids(ids)

    name = deck_name(source)
    deck_id, link = create_deck(drive, name, folder)
    log(f"Created '{name}'")

    ok = build_slides(slides, deck_id, pages, parts, B, ids)
    log("=" * 50)
    log(f"Slides written: {ok} of {len(pages)}")
    log(link)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
