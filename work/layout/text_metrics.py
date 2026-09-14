#!/usr/bin/env python3
"""
Measure how much room a piece of text will take.

The layout has to know, before PowerPoint ever opens the
file, how many lines a sentence wraps to and how wide the
longest of them is. Every box is sized from those two
numbers, and since the boxes now carry a visible border,
guessing wrong shows up as a ragged edge or empty space.

The earlier version multiplied the character count by an
average glyph width. That is fine for deciding a font size
and useless for drawing a box: "Artificial Analysis
Intelligence Index" counted as two lines and rendered as
one, leaving a finger of blank yellow under the text.

So this measures real glyphs. Calibri is not installed on
macOS, but Arial is, and the two are close enough in shape
that a single width factor converts between them. The
factor was fitted against twelve line breaks read off
rendered slides at two font sizes and two column widths,
which brackets it to [0.9070, 0.9225]. The midpoint is used.

If Arial is missing the module falls back to the old
average-width estimate, so nothing breaks on another
machine, it merely gets less exact.

Usage:
    from layout.text_metrics import text_width, wrap_lines
    text_width("hello", 12)            # inches
    wrap_lines("a long sentence", 3.0, 12)

Created: 2026-09-12
Last updated: 2026-09-12
"""

from functools import lru_cache

# Width of Calibri relative to Arial, fitted from rendered
# slides. See the module docstring for the bracket.
CALIBRI_OVER_ARIAL = 0.9148

# Measuring at a large pixel size keeps rounding out of the
# ratio; the result is scaled back to the real font size.
SAMPLE_PX = 400

FACES = {
    False: "Arial.ttf",
    True: "Arial Bold.ttf",
}

# Used only when no TrueType face can be loaded.
FALLBACK_RATIO = {False: 0.46, True: 0.50}


# --------------------------------------------------------------
@lru_cache(maxsize=4)
def _face(bold):
    """Load a measuring face once, or None if absent."""
    try:
        from PIL import ImageFont
        return ImageFont.truetype(FACES[bold], SAMPLE_PX)
    except Exception:
        return None


# --------------------------------------------------------------
@lru_cache(maxsize=8192)
def _em_width(text, bold):
    """Width of text in multiples of the font size."""
    face = _face(bold)
    if face is None:
        return len(text) * FALLBACK_RATIO[bold]
    return face.getlength(text) / SAMPLE_PX


# --------------------------------------------------------------
def text_width(text, size_pt, bold=False):
    """Width of one unwrapped string, in inches."""
    em = _em_width(text, bold) * CALIBRI_OVER_ARIAL
    return em * size_pt / 72.0


# --------------------------------------------------------------
def wrap(text, width_in, size_pt, bold=False):
    """Break text into the lines PowerPoint would draw."""
    words = text.split()
    if not words:
        return [""]

    lines = []
    current = words[0]
    for word in words[1:]:
        trial = f"{current} {word}"
        if text_width(trial, size_pt, bold) <= width_in:
            current = trial
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


# --------------------------------------------------------------
def wrap_lines(text, width_in, size_pt, bold=False):
    """Count the lines text wraps to in a given width."""
    return len(wrap(text, width_in, size_pt, bold))


# --------------------------------------------------------------
def wrapped_width(text, width_in, size_pt, bold=False):
    """Width of the longest line after wrapping, in inches."""
    lines = wrap(text, width_in, size_pt, bold)
    return max(
        text_width(line, size_pt, bold) for line in lines
    )


# --------------------------------------------------------------
def widest(texts, size_pt, bold=False):
    """Width of the widest string in a list, in inches."""
    if not texts:
        return 0.0
    return max(text_width(t, size_pt, bold) for t in texts)
