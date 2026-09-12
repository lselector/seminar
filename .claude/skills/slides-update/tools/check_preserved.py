#!/usr/bin/env python3
"""
Prove that an edit pass did not undo a human's work.

Compares a deck Markdown file before and after an editing
pass and fails when protected content changed. It exists
because "the AI promised not to touch that" is not a
guarantee, and a script that checks is.

What counts as a violation:

  CHANGED      a locked item's text is different
  REMOVED      an item that existed is gone
  RECREATED    a topic in the deleted sidecar is back
  RESURRECTED  an in-file tombstone was taken off

Additions are fine and are only listed for information.
A locked "## slide:" section protects every news item
inside it, not just its own heading.

The sidecar is found next to the new file: for
2026-09-18-AI-News.md it is 2026-09-18-AI-News-deleted.md.
Every "###" headline in it is a topic the author threw out,
and none of them may appear in the deck. Headlines are
compared case-insensitively with whitespace collapsed, so a
small rewording is still caught.

Items are matched by heading text, not by position, so
inserting a new story in the middle is not mistaken for a
rewrite. Trailing whitespace and blank lines are ignored,
since neither changes what a slide says.

This tool deliberately does its own lightweight splitting
of the file rather than importing deck_parser. A checker
that shares code with the thing it checks shares its bugs.

Exit code 0 means every protected item survived.

Usage:
    python3 check_preserved.py before.md after.md

Created: 2026-09-12
Last updated: 2026-09-12
"""

import os
import sys
from dataclasses import dataclass

FLAG_LOCKED = "<!-- locked -->"
FLAG_DELETED = "<!-- deleted -->"

SECTION_PREFIX = "## "
BLOCK_PREFIX = "### "

DELETED_SUFFIX = "-deleted.md"

PREAMBLE_HEADING = "(deck title and epigraph)"


@dataclass
class Segment:
    """One heading in a deck file and the lines under it."""

    kind: str
    heading: str
    body: str
    locked: bool
    deleted: bool


# --------------------------------------------------------------
def segment_kind(line):
    """Return 'block', 'section', or None for a line."""
    if line.startswith(BLOCK_PREFIX):
        return "block"
    if line.startswith(SECTION_PREFIX):
        return "section"
    return None


# --------------------------------------------------------------
def clean_body(lines):
    """Normalize a segment body for comparison."""
    trimmed = [line.rstrip() for line in lines]
    while trimmed and not trimmed[-1]:
        trimmed.pop()
    return "\n".join(trimmed)


# --------------------------------------------------------------
def make_segment(kind, heading, lines):
    """Build one Segment from its heading and body."""
    body = clean_body(lines)
    return Segment(
        kind=kind,
        heading=heading.strip(),
        body=body,
        locked=FLAG_LOCKED in body,
        deleted=FLAG_DELETED in body,
    )


# --------------------------------------------------------------
def raw_segments(text):
    """Split a deck file into Segments, in file order.

    The lines above the first heading become a segment of
    their own, so the deck title and the epigraph are
    covered by the same rules as everything else.
    """
    found = []
    kind = heading = None
    lines = []
    preamble = []

    for line in text.splitlines():
        this_kind = segment_kind(line)
        if this_kind:
            if kind:
                found.append(make_segment(kind, heading, lines))
            kind, heading, lines = this_kind, line, []
        elif kind:
            lines.append(line)
        else:
            preamble.append(line)

    if kind:
        found.append(make_segment(kind, heading, lines))
    if any(line.strip() for line in preamble):
        found.insert(0, make_segment(
            "preamble", PREAMBLE_HEADING, preamble
        ))
    return found


# --------------------------------------------------------------
def apply_section_flags(segments):
    """Let a locked or deleted section cover its blocks."""
    locked = deleted = False
    for segment in segments:
        if segment.kind == "section":
            locked, deleted = segment.locked, segment.deleted
            continue
        segment.locked = segment.locked or locked
        segment.deleted = segment.deleted or deleted
    return segments


# --------------------------------------------------------------
def read_segments(path):
    """Read one deck file into flag-resolved Segments."""
    try:
        with open(path, encoding="utf-8") as handle:
            text = handle.read()
    except OSError as exc:
        print(f"ERROR: cannot read {path}: {exc}")
        sys.exit(2)
    return apply_section_flags(raw_segments(text))


# --------------------------------------------------------------
def index_by_heading(segments):
    """Group Segments by their heading text."""
    index = {}
    for segment in segments:
        index.setdefault(segment.heading, []).append(segment)
    return index


# --------------------------------------------------------------
def bodies(segments):
    """Sorted bodies of a group, for order-free compare."""
    return sorted(s.body for s in segments)


# --------------------------------------------------------------
def any_flag(segments, name):
    """Check whether any segment in a group has a flag."""
    return any(getattr(s, name) for s in segments)


# --------------------------------------------------------------
def check_locked(heading, before, after):
    """Report a locked item that changed or vanished."""
    if not after:
        return [("REMOVED", heading, "locked item is gone")]
    if bodies(before) != bodies(after):
        return [(
            "CHANGED", heading,
            "locked item's text was edited"
        )]
    return []


# --------------------------------------------------------------
def check_deleted(heading, before, after):
    """Report a tombstoned item brought back to life."""
    if not after:
        return [(
            "REMOVED", heading,
            "tombstone was dropped, so the topic can "
            "come back later"
        )]
    if not any_flag(after, "deleted"):
        return [(
            "RESURRECTED", heading,
            "the deleted marker was taken off"
        )]
    return []


# --------------------------------------------------------------
def check_one(heading, before, after):
    """Check one heading group for violations."""
    if any_flag(before, "deleted"):
        return check_deleted(heading, before, after)
    if any_flag(before, "locked"):
        return check_locked(heading, before, after)
    if not after:
        return [(
            "REMOVED", heading,
            "item existed before and is gone now"
        )]
    return []


# --------------------------------------------------------------
def compare(before, after):
    """Collect every violation between two deck files."""
    old = index_by_heading(before)
    new = index_by_heading(after)

    problems = []
    for heading, group in old.items():
        problems.extend(
            check_one(heading, group, new.get(heading, []))
        )
    return problems


# --------------------------------------------------------------
def normalize(text):
    """Fold a heading for forgiving comparison."""
    return " ".join(text.lower().split())


# --------------------------------------------------------------
def tokens(text):
    """The words of a headline worth matching on."""
    words = normalize(text).split()
    return {w.strip(".,:;!?()[]'\"") for w in words
            if len(w) > 3}


# --------------------------------------------------------------
def same_topic(one, other):
    """Judge whether two headlines name the same topic.

    A deliberate copy of deck_parser.same_topic. This tool
    keeps its own implementation so it cannot inherit a bug
    from the code it checks. The two are pinned together by
    test_checker_agrees_on_topic_matching in test_deck.py,
    so change them as a pair.
    """
    first, second = normalize(one), normalize(other)
    if first == second:
        return True
    if first and second and (first in second or
                             second in first):
        return True

    left, right = tokens(one), tokens(other)
    if not left or not right:
        return False
    shared = len(left & right)
    return shared / min(len(left), len(right)) >= 0.7


# --------------------------------------------------------------
def sidecar_path(after_path):
    """Build the deleted-topics path beside a deck file."""
    stem = after_path
    if stem.endswith(".md"):
        stem = stem[:-3]
    return stem + DELETED_SUFFIX


# --------------------------------------------------------------
def read_sidecar(after_path):
    """Read the thrown-out headings, or return []."""
    path = sidecar_path(after_path)
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as handle:
        return [
            s.heading[len(BLOCK_PREFIX):].strip()
            for s in raw_segments(handle.read())
            if s.kind == "block"
        ]


# --------------------------------------------------------------
def check_recreated(after, deleted):
    """Report deck items that the author had thrown out."""
    problems = []
    for segment in after:
        if segment.kind != "block":
            continue
        name = segment.heading[len(BLOCK_PREFIX):].strip()
        for topic in deleted:
            if not same_topic(name, topic):
                continue
            problems.append((
                "RECREATED", segment.heading,
                f"the author threw this out: '{topic}'"
            ))
            break
    return problems


# --------------------------------------------------------------
def added_headings(before, after):
    """List headings that appear only in the new file."""
    old = index_by_heading(before)
    return [
        s.heading for s in after
        if s.heading not in old
    ]


# --------------------------------------------------------------
def count_protected(segments):
    """Count locked and tombstoned segments."""
    locked = sum(1 for s in segments if s.locked)
    deleted = sum(1 for s in segments if s.deleted)
    return locked, deleted


# --------------------------------------------------------------
def report(problems, added, before, dropped):
    """Print the outcome and return an exit code."""
    locked, deleted = count_protected(before)
    print(f"Protected before the edit: {locked} locked, "
          f"{deleted} deleted")
    print(f"Topics in the deleted file: {dropped}")

    for heading in added:
        print(f"  added        {heading}")

    if not problems:
        print("OK: every protected item survived")
        return 0

    print(f"\n{len(problems)} violation(s):")
    for kind, heading, why in problems:
        print(f"  {kind:12s} {heading}")
        print(f"               {why}")
    print("\nRestore the items above before continuing.")
    return 1


# --------------------------------------------------------------
def main():
    """Compare two versions of one deck Markdown file."""
    if len(sys.argv) != 3:
        print(
            "Usage: python3 check_preserved.py "
            "before.md after.md"
        )
        sys.exit(2)

    before = read_segments(sys.argv[1])
    after = read_segments(sys.argv[2])
    deleted = read_sidecar(sys.argv[2])

    problems = compare(before, after)
    problems.extend(check_recreated(after, deleted))
    added = added_headings(before, after)
    sys.exit(report(problems, added, before, len(deleted)))


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
