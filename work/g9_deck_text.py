#!/usr/bin/env python3
"""
Extract the text of a live deck as plain Markdown.

Reads the talk, slide 1 up to the separator, and writes each
slide's title and the words in its boxes, top to bottom.
Tables come out one row per line, cells joined by " | ".
Parked slides and speaker notes are left out. Changes
nothing in the deck.

Used by the gslides-epigraphs skill, which reads the text to
write epigraphs for the week, and handy whenever the words of
a deck are wanted without opening it.

Usage:
    python3 g9_deck_text.py                      to stdout
    python3 g9_deck_text.py --out deck.md
    python3 g9_deck_text.py 2026-09-18 --out deck.md

Created: 2026-09-18
Last updated: 2026-09-18
"""

import sys

from gslides import deck as G
from gslides.client import log
from gslides.step import start


# --------------------------------------------------------------
def split_out(argv):
    """The --out path, and the rest of the command line."""
    if "--out" not in argv:
        return None, argv
    at = argv.index("--out")
    if at + 1 >= len(argv):
        log("ERROR: --out needs a file name")
        sys.exit(1)
    return argv[at + 1], argv[:at] + argv[at + 2:]


# --------------------------------------------------------------
def shape_lines(shape):
    """The non-empty lines one shape holds."""
    if shape.kind == G.KIND_TABLE:
        rows = [" | ".join(c.strip() for c in row)
                for row in shape.cells]
        return [r for r in rows if r.strip(" |")]
    return shape.paragraphs()


# --------------------------------------------------------------
def slide_text(slide, number):
    """One slide as a Markdown section."""
    shapes = sorted((s for s in slide.shapes
                     if s.kind in (G.KIND_TEXT, G.KIND_TABLE)),
                    key=lambda s: (s.rect.y, s.rect.x))
    titles = [s for s in shapes
              if s.kind == G.KIND_TEXT and G.is_title(s)]
    titles.sort(key=lambda s: not s.id.endswith("-title"))
    title = titles[0].first_line() if titles else ""
    parts = [f"## Slide {number}: {title}".rstrip(": ")]
    for shape in shapes:
        lines = shape_lines(shape)
        if not lines or (lines == [title]):
            continue
        parts.append("\n".join(lines))
    return "\n\n".join(parts)


# --------------------------------------------------------------
def deck_text(deck):
    """The whole talk as one Markdown document."""
    sections = [f"# {deck.title}"]
    for number, slide in enumerate(deck.main(), start=1):
        sections.append(slide_text(slide, number))
    return "\n\n".join(sections) + "\n"


# --------------------------------------------------------------
def main():
    """Write the text of one deck to a file or stdout."""
    out, argv = split_out(sys.argv[1:])
    step = start("Extract the text of a deck", argv)
    text = deck_text(G.read_deck(step.slides, step.deck_id))
    if not out:
        print(text, end="")
        return
    with open(out, "w", encoding="utf-8") as handle:
        handle.write(text)
    log(f"deck text: {len(text.split())} words -> {out}")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
