#!/usr/bin/env python3
"""
Draw the LM Arena benchmarks slide.

Turns the JSON that leaderboard.py writes into two plain
columns, each line reading "marker score name" with the
marker coloured by vendor, plus a legend across the top.

This page was a real PPTX table first, and that was wrong.
PowerPoint enforces a minimum row height regardless of what
the file asks for, so 25 rows always spilled off the slide.
Ordinary paragraphs have no such floor: the line spacing
can be set exactly and the font size is computed from the
space available, so the page fits by arithmetic.

Reads nothing and writes nothing. It is handed a slide and
the already-loaded leaderboard data.

Usage:
    from bench_page import render_benchmarks_page
    render_benchmarks_page(slide, page, data)

Created: 2026-09-12
Last updated: 2026-09-12
"""

import deck_layout as L
from pptx_text import (
    BOX_PAD, COLOR_LINK, COLOR_MUTED, add_textbox,
    first_or_new, style_run,
)
from text_metrics import text_width
from pptx.dml.color import RGBColor
from pptx.util import Pt

BENCH_ROWS = 25
BENCH_CAPTION_SIZE = 11
BENCH_GAP = 0.40
BENCH_TOP = 0.62
BENCH_SIZE_MAX = 11
BENCH_SIZE_MIN = 6
BENCH_LINE_RATIO = 1.20
LEGEND_SIZE = 9

# The vendor swatch. A coloured character in an ordinary
# paragraph has no minimum height, unlike a table cell.
MARKER = "\u25a0"

VENDOR_COLORS = {
    "anthropic": RGBColor(0x44, 0x72, 0xC4),
    "google": RGBColor(0xFF, 0x00, 0x00),
    "openai": RGBColor(0xFF, 0xD7, 0x00),
    "opensource": RGBColor(0x00, 0xB0, 0x50),
    "other": RGBColor(0xA6, 0xA6, 0xA6),
}

# The project's spelling of months, so this date matches the
# deck title. deck_init.py keeps the same list for the same
# reason; test_month_names_agree pins them together.
MONTHS = [
    "Jan", "Feb", "March", "April", "May", "June",
    "July", "Aug", "Sept", "Oct", "Nov", "Dec",
]

VENDOR_LABELS = [
    ("anthropic", "Claude"),
    ("google", "Gemini"),
    ("openai", "OpenAI"),
    ("opensource", "Open source"),
    ("other", "Other"),
]


# --------------------------------------------------------------
# --------------------------------------------------------------
def bench_font_size(rows, height_in):
    """Largest size at which the rows fit their column."""
    if rows <= 0:
        return BENCH_SIZE_MAX
    available = height_in * 72.0
    fits = available / (rows * BENCH_LINE_RATIO)
    return max(BENCH_SIZE_MIN, min(BENCH_SIZE_MAX, fits))


# --------------------------------------------------------------
def tighten(para, size):
    """Remove the space a paragraph adds around itself."""
    para.line_spacing = Pt(size * BENCH_LINE_RATIO)
    para.space_before = Pt(0)
    para.space_after = Pt(0)


# --------------------------------------------------------------
def row_text(entry):
    """The three pieces of one benchmark line."""
    return (
        f"{MARKER} ",
        f"{entry.get('score', '')}  ",
        entry.get("name", ""),
    )


# --------------------------------------------------------------
def row_width(entry, size):
    """Width of one rendered benchmark line, in inches."""
    marker, score, name = row_text(entry)
    return (
        text_width(marker, size)
        + text_width(score, size, bold=True)
        + text_width(name, size)
    )


# --------------------------------------------------------------
def column_width(board, size):
    """Width the widest line of a column needs."""
    entries = board.get("entries", [])[:BENCH_ROWS]
    widths = [row_width(e, size) for e in entries]
    widths.append(text_width(
        board.get("label", ""), BENCH_CAPTION_SIZE, bold=True
    ))
    widths.append(
        text_width(board.get("url", ""), L.LINK_SIZE)
    )
    return max(widths) + L.TEXT_INSET + BOX_PAD


# --------------------------------------------------------------
def add_bench_row(frame, entry, size, used):
    """Write one model as marker, score, then name."""
    para = first_or_new(frame, used)
    tighten(para, size)

    vendor = entry.get("vendor", "other")
    marker_text, score_text, name_text = row_text(entry)

    marker = para.add_run()
    marker.text = marker_text
    style_run(
        marker, size,
        color=VENDOR_COLORS.get(
            vendor, VENDOR_COLORS["other"]
        )
    )

    score = para.add_run()
    score.text = score_text
    style_run(score, size, bold=True)

    name = para.add_run()
    name.text = name_text
    style_run(name, size)


# --------------------------------------------------------------
def add_bench_caption(frame, board):
    """Write the board's name and source link on top."""
    label = first_or_new(frame, 0)
    tighten(label, BENCH_CAPTION_SIZE)
    run = label.add_run()
    run.text = board.get("label", "")
    style_run(run, BENCH_CAPTION_SIZE, bold=True)

    url = board.get("url", "")
    if not url:
        return 1
    link = first_or_new(frame, 1)
    tighten(link, L.LINK_SIZE)
    run = link.add_run()
    run.text = url
    style_run(run, L.LINK_SIZE, color=COLOR_LINK)
    run.hyperlink.address = url
    return 2


# --------------------------------------------------------------
def caption_height():
    """Height the caption and its link take up."""
    points = (BENCH_CAPTION_SIZE + L.LINK_SIZE)
    return L.points_to_inches(points * BENCH_LINE_RATIO)


# --------------------------------------------------------------
def build_bench_column(slide, left, width, board, size):
    """Draw one leaderboard as a single boxed column."""
    entries = board.get("entries", [])[:BENCH_ROWS]
    if not entries:
        return
    height = L.BAND_BOT - BENCH_TOP

    frame = add_textbox(
        slide, L.Rect(left, BENCH_TOP, width, height),
        boxed=True
    )
    frame.word_wrap = False
    used = add_bench_caption(frame, board)
    for entry in entries:
        add_bench_row(frame, entry, size, used)
        used += 1


# --------------------------------------------------------------
def pretty_date(iso):
    """Turn 2026-09-11 into Sept 11, 2026."""
    try:
        year, month, day = (int(p) for p in iso.split("-"))
        return f"{MONTHS[month - 1]} {day}, {year}"
    except (AttributeError, ValueError, IndexError):
        return ""


# --------------------------------------------------------------
def cutoff_note(boards):
    """One line saying what the numbers are current to."""
    dates = sorted({
        b.get("cutoff") for b in boards if b.get("cutoff")
    })
    if not dates:
        return ""
    if len(dates) == 1:
        return f"Votes counted through {pretty_date(dates[0])}"
    return "  ".join(
        f"{b['label']}: {pretty_date(b.get('cutoff'))}"
        for b in boards if b.get("cutoff")
    )


# --------------------------------------------------------------
def render_cutoff(slide, boards):
    """Draw the vote cutoff at the right of the legend."""
    note = cutoff_note(boards)
    if not note:
        return
    width = text_width(note, LEGEND_SIZE) + L.TEXT_INSET
    rect = L.Rect(
        L.SLIDE_W - L.MARGIN - width,
        L.TITLE_H + 0.02, width, 0.22
    )
    frame = add_textbox(slide, rect)
    para = frame.paragraphs[0]
    tighten(para, LEGEND_SIZE)
    run = para.add_run()
    run.text = note
    style_run(run, LEGEND_SIZE, color=COLOR_MUTED)


# --------------------------------------------------------------
def render_legend(slide):
    """Name the colours used by the vendor markers."""
    width = sum(
        text_width(f"{MARKER} ", LEGEND_SIZE)
        + text_width(f"{label}    ", LEGEND_SIZE)
        for _, label in VENDOR_LABELS
    ) + L.TEXT_INSET
    rect = L.Rect(
        L.MARGIN, L.TITLE_H + 0.02, width, 0.22
    )
    frame = add_textbox(slide, rect, boxed=True, fit=True)
    para = frame.paragraphs[0]
    tighten(para, LEGEND_SIZE)
    for vendor, label in VENDOR_LABELS:
        marker = para.add_run()
        marker.text = f"{MARKER} "
        style_run(
            marker, LEGEND_SIZE,
            color=VENDOR_COLORS[vendor]
        )
        text = para.add_run()
        text.text = f"{label}    "
        style_run(text, LEGEND_SIZE, color=COLOR_MUTED)


# --------------------------------------------------------------
def render_benchmarks_page(slide, data):
    """Draw the benchmark columns under an existing title."""
    if not data:
        return
    boards = data.get("boards", [])[:2]
    if not boards:
        return
    render_legend(slide)
    render_cutoff(slide, boards)

    height = L.BAND_BOT - BENCH_TOP
    rows = max(
        len(b.get("entries", [])[:BENCH_ROWS]) for b in boards
    )
    size = bench_font_size(
        rows, height - caption_height() - BOX_PAD
    )
    limit = (L.CONTENT_W - BENCH_GAP) / 2
    widths = [
        min(limit, column_width(b, size)) for b in boards
    ]

    # First column on the left edge, second on the right, so
    # the space they no longer need reads as a gutter rather
    # than a gap at the end of the slide.
    lefts = [L.MARGIN]
    if len(boards) > 1:
        lefts.append(L.SLIDE_W - L.MARGIN - widths[1])
    for board, left, width in zip(boards, lefts, widths):
        build_bench_column(slide, left, width, board, size)
