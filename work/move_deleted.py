#!/usr/bin/env python3
"""
Move items marked deleted out of a deck, into its sidecar.

There are two ways to throw a topic out of a deck. Cut the
item out by hand and paste it into the "-deleted.md" file,
or mark it in the deck and let this script do the moving.
This is the second way, and it is easier when several items
go at once.

Mark an item by putting the flag on its own line:

    ### A story I do not want
    <!-- deleted -->
    - bullets stay with it

Then run this script. Every flagged item is cut out of the
deck and appended to the sidecar file, and the flag line
itself is dropped, because in the sidecar the heading alone
is the record.

The flag also works on a "## slide:" heading, in which case
the whole section and all of its items move together.

The deck is edited in place and a copy of the previous
version is left in "<deck>.bak". Use --dry-run first to see
what would move without changing anything.

Text is moved verbatim. This works on the file as lines,
not through the parser, so bullets, images, links, and
blank lines survive the trip unchanged.

Usage:
    python3 move_deleted.py --dry-run
    python3 move_deleted.py
    python3 move_deleted.py 2026-09-18-AI-News.md

Run s3_make_pptx.py afterwards to rebuild the slides.

Created: 2026-09-12
Last updated: 2026-09-12
"""

import datetime as dt
import glob
import os
import shutil
import sys
from dataclasses import dataclass, field

from deck_parser import (
    DeckError, FLAG_DELETED, RE_FLAG, deleted_sidecar_path,
    parse_deck_file,
)

DECK_GLOB = "*-AI-News.md"
SECTION_PREFIX = "## "
BLOCK_PREFIX = "### "


@dataclass
class Segment:
    """One heading of a deck file and its raw lines."""

    kind: str
    heading: str
    body: list = field(default_factory=list)

    # --------------------------------------
    def text(self):
        """Rebuild this segment's original lines."""
        return "\n".join([self.heading] + self.body)


# --------------------------------------------------------------
def log_message(message):
    """Print timestamped log message."""
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}")


# --------------------------------------------------------------
def segment_kind(line):
    """Return 'block', 'section', or None for a line."""
    if line.startswith(BLOCK_PREFIX):
        return "block"
    if line.startswith(SECTION_PREFIX):
        return "section"
    return None


# --------------------------------------------------------------
def split_source(text):
    """Split deck text into a preamble and its segments."""
    preamble = []
    segments = []
    current = None

    for line in text.splitlines():
        kind = segment_kind(line)
        if kind:
            if current:
                segments.append(current)
            current = Segment(kind=kind, heading=line)
        elif current:
            current.body.append(line)
        else:
            preamble.append(line)

    if current:
        segments.append(current)
    return preamble, segments


# --------------------------------------------------------------
def is_deleted_flag(line):
    """Check whether one line is the deleted marker."""
    match = RE_FLAG.match(line.strip())
    return bool(match) and \
        match.group(1).lower() == FLAG_DELETED


# --------------------------------------------------------------
def has_deleted_flag(body):
    """Check whether a segment body carries the marker."""
    return any(is_deleted_flag(line) for line in body)


# --------------------------------------------------------------
def strip_deleted_flag(body):
    """Drop the marker line from a segment body."""
    return [
        line for line in body if not is_deleted_flag(line)
    ]


# --------------------------------------------------------------
def plan_moves(segments):
    """Split segments into those kept and those moved."""
    keep, move = [], []
    in_deleted_section = False

    for segment in segments:
        if segment.kind == "section":
            in_deleted_section = has_deleted_flag(segment.body)
            flagged = in_deleted_section
        else:
            flagged = in_deleted_section or \
                has_deleted_flag(segment.body)
        (move if flagged else keep).append(segment)

    return keep, move


# --------------------------------------------------------------
def heading_label(segment):
    """Readable name of a segment, for the log."""
    prefix = BLOCK_PREFIX if segment.kind == "block" \
        else SECTION_PREFIX
    return segment.heading[len(prefix):].strip()


# --------------------------------------------------------------
def render_deck(preamble, keep):
    """Rebuild the deck file from what stays in it."""
    parts = list(preamble)
    for segment in keep:
        parts.append(segment.text())
    return "\n".join(parts).rstrip("\n") + "\n"


# --------------------------------------------------------------
def render_moved(move, today):
    """Build the text appended to the sidecar file."""
    parts = []
    for segment in move:
        body = strip_deleted_flag(segment.body)
        block = "\n".join(
            [segment.heading, f"<!-- moved: {today} -->"]
            + body
        )
        parts.append(block.rstrip("\n"))
    return "\n\n".join(parts) + "\n"


# --------------------------------------------------------------
def sidecar_header(title):
    """Opening text for a sidecar created on demand."""
    return (
        f"# Topics thrown out of {title}\n\n"
        "Cut an item out of the deck and paste it in here to "
        "delete it,\nor mark it in the deck and run "
        "move_deleted.py.\n\n"
        "Keep the \"###\" heading: that heading is the "
        "record. Nothing here\ngoes on a slide, and nothing "
        "here may be written back into the deck.\n"
    )


# --------------------------------------------------------------
def append_to_sidecar(path, title, text):
    """Append moved items, creating the file if needed."""
    fresh = not os.path.isfile(path)
    with open(path, "a", encoding="utf-8") as handle:
        if fresh:
            handle.write(sidecar_header(title))
        handle.write("\n" + text)
    return fresh


# --------------------------------------------------------------
def deck_title(preamble):
    """Take the deck title from its level-one heading."""
    for line in preamble:
        if line.startswith("# "):
            return line[2:].strip()
    return "this deck"


# --------------------------------------------------------------
def newest_deck():
    """Find the newest deck Markdown in this directory."""
    names = sorted(glob.glob(DECK_GLOB))
    return names[-1] if names else None


# --------------------------------------------------------------
def resolve_deck(argv):
    """Take the deck path from the command line."""
    named = [a for a in argv[1:] if not a.startswith("-")]
    source = named[0] if named else newest_deck()
    if not source:
        log_message(
            f"ERROR: no {DECK_GLOB} file here. "
            f"Usage: python3 move_deleted.py deck.md"
        )
        sys.exit(1)
    if not os.path.isfile(source):
        log_message(f"ERROR: {source} not found")
        sys.exit(1)
    if not named:
        log_message(f"No file given, using {source}")
    return source


# --------------------------------------------------------------
def validate_deck(source):
    """Refuse to operate on a deck that will not parse."""
    try:
        parse_deck_file(source)
    except DeckError as exc:
        log_message(f"ERROR in {source}: {exc}")
        log_message(
            "Nothing was moved. Fix that line and re-run."
        )
        sys.exit(1)


# --------------------------------------------------------------
def report_plan(move):
    """List what is about to move, or say nothing is."""
    if not move:
        log_message("Nothing is marked deleted.")
        log_message(
            f"Mark an item with <!-- {FLAG_DELETED} --> "
            f"and run this again."
        )
        return False
    for segment in move:
        log_message(
            f"  {segment.kind:7s} {heading_label(segment)}"
        )
    return True


# --------------------------------------------------------------
def write_changes(source, preamble, keep, move, title):
    """Move the flagged items, keeping a backup."""
    sidecar = deleted_sidecar_path(source)
    backup = source + ".bak"
    shutil.copy2(source, backup)

    today = dt.date.today().isoformat()
    fresh = append_to_sidecar(
        sidecar, title, render_moved(move, today)
    )

    with open(source, "w", encoding="utf-8") as handle:
        handle.write(render_deck(preamble, keep))

    log_message(f"Backup of the deck: {backup}")
    if fresh:
        log_message(f"Created {sidecar}")
    log_message(f"Appended {len(move)} item(s) to {sidecar}")


# --------------------------------------------------------------
def main():
    """Move every flagged item into the sidecar file."""
    source = resolve_deck(sys.argv)
    dry_run = "--dry-run" in sys.argv

    validate_deck(source)
    with open(source, encoding="utf-8") as handle:
        preamble, segments = split_source(handle.read())

    keep, move = plan_moves(segments)
    log_message(f"Marked deleted in {source}: {len(move)}")
    if not report_plan(move):
        return

    if dry_run:
        log_message("Dry run, nothing was changed.")
        return

    write_changes(
        source, preamble, keep, move, deck_title(preamble)
    )
    log_message("Next: python3 s3_make_pptx.py")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
