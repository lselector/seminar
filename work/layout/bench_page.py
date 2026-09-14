#!/usr/bin/env python3
"""
Measurements and wording for the LM Arena benchmarks slide.

The page shows the JSON sources/leaderboard.py writes as
two plain text columns, each line reading "marker score
name" with the marker coloured by vendor, a legend across
the top, and a
note saying which votes the numbers count.

This module holds only the numbers and words: sizes, vendor
colours, how wide and tall a column must be, and the cutoff
note. gslides/render.py turns them into Slides requests.

Plain paragraphs rather than a table, because a table has a
minimum row height and 25 rows then spill off the slide.
Paragraphs fit by arithmetic.

Usage:
    from layout import bench_page as B
    size = B.bench_font_size(25, 4.5)
    width = B.column_width(board, size)
    note = B.cutoff_note(boards)

Created: 2026-09-12
Last updated: 2026-09-14 (PPTX drawing removed)
"""

from layout import deck_layout as L
from layout.text_metrics import text_width

BENCH_ROWS = 25

# This page carries 50 lines of data, so it runs smaller
# than the rest of the deck. The caption follows the usual
# "one point above the body, bold and red" rule.
BENCH_SIZE = 9
BENCH_CAPTION_SIZE = BENCH_SIZE + 1
BENCH_SIZE_MIN = 6
BENCH_LINE_RATIO = 1.20

BENCH_GAP = 0.50
LEGEND_SIZE = 9

# The vote-cutoff note reads at the normal body size, on the
# header row beside the legend.
DATE_SIZE = 12
HEADER_Y = 0.38
HEADER_GAP = 0.08

# Vertical breathing room inside a box around its text.
BOX_PAD = 0.06

# The vendor swatch: a coloured character.
MARKER = "■"

VENDOR_COLORS = {
    "anthropic": (0x44, 0x72, 0xC4),
    "google": (0xFF, 0x00, 0x00),
    "openai": (0xFF, 0xD7, 0x00),
    "opensource": (0x00, 0xB0, 0x50),
    "other": (0xA6, 0xA6, 0xA6),
}

# The project's spelling of months, so this date matches the
# deck title. gslides/deck.py keeps the same list for the same
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
def header_height():
    """Height of the legend and date row under the title."""
    tall = DATE_SIZE * BENCH_LINE_RATIO / 72.0
    return tall + BOX_PAD


# --------------------------------------------------------------
def bench_top():
    """Top of the two columns, clear of the header row."""
    return HEADER_Y + header_height() + HEADER_GAP


# --------------------------------------------------------------
def bench_font_size(rows, height_in):
    """Body size, stepped down only if the rows will not fit."""
    if rows <= 0:
        return BENCH_SIZE
    available = height_in * 72.0
    fits = available / (rows * BENCH_LINE_RATIO)
    return max(BENCH_SIZE_MIN, min(BENCH_SIZE, fits))


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
def caption_height():
    """Height the caption and its link take up."""
    points = (BENCH_CAPTION_SIZE + L.LINK_SIZE)
    return L.points_to_inches(points * BENCH_LINE_RATIO)


# --------------------------------------------------------------
def column_height(rows, size):
    """Height a column needs, caption and rows together."""
    body = rows * size * BENCH_LINE_RATIO / 72.0
    return caption_height() + body + BOX_PAD


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
