#!/usr/bin/env python3
"""
Work out which weekly deck to edit, and create it if new.

Resolves the seminar date, builds the three file names the
project uses, reports whether they exist, and optionally
writes fresh skeletons from the templates.

The three files for one week:

    2026-09-18-AI-News.md              the deck
    2026-09-18-AI-News-generated.pptx  built from it
    2026-09-18-AI-News-deleted.md      topics thrown out

Date rules:
  - an explicit YYYY-MM-DD argument always wins
  - otherwise the next Friday
  - on a Friday that means today, not a week later

Output is one JSON line, so the caller never has to parse
prose:

  {"date": "2026-09-18", "title": "AI News - Sept 18,
   2026", "md": "work/2026-09-18-AI-News.md",
   "pptx": "work/2026-09-18-AI-News-generated.pptx",
   "deleted": "work/2026-09-18-AI-News-deleted.md",
   "md_exists": true, "pptx_exists": false,
   "deleted_exists": true, "created": false}

Run it from the repository root.

Usage:
    python3 deck_init.py
    python3 deck_init.py 2026-10-02
    python3 deck_init.py --create
    python3 deck_init.py 2026-10-02 --create --dir work

Created: 2026-09-12
Last updated: 2026-09-12
"""

import argparse
import datetime as dt
import json
import os
import sys
from pathlib import Path

FRIDAY = 4
SUFFIX_MD = "-AI-News.md"
SUFFIX_PPTX = "-AI-News-generated.pptx"
SUFFIX_DELETED = "-AI-News-deleted.md"
TEMPLATE = "reference/deck_skeleton.md"
TEMPLATE_DELETED = "reference/deleted_skeleton.md"

# The month spellings used in the existing decks, so a
# generated title matches five years of archive.
MONTHS = [
    "Jan", "Feb", "March", "April", "May", "June",
    "July", "Aug", "Sept", "Oct", "Nov", "Dec",
]


# --------------------------------------------------------------
def next_friday(today=None):
    """Return the coming Friday, or today if it is one."""
    today = today or dt.date.today()
    return today + dt.timedelta(
        days=(FRIDAY - today.weekday()) % 7
    )


# --------------------------------------------------------------
def parse_date(text):
    """Parse a YYYY-MM-DD argument, or exit with help."""
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        print(
            f"ERROR: '{text}' is not a date in "
            f"YYYY-MM-DD form",
            file=sys.stderr
        )
        sys.exit(1)


# --------------------------------------------------------------
def resolve_date(argument):
    """Pick the seminar date from the argument or today."""
    if argument:
        return parse_date(argument)
    return next_friday()


# --------------------------------------------------------------
def deck_title(date):
    """Build the deck title in the project's own style."""
    month = MONTHS[date.month - 1]
    return f"AI News - {month} {date.day}, {date.year}"


# --------------------------------------------------------------
def deck_paths(date, directory):
    """Build the Markdown, PPTX, and deleted paths."""
    stamp = date.isoformat()
    return (
        os.path.join(directory, f"{stamp}{SUFFIX_MD}"),
        os.path.join(directory, f"{stamp}{SUFFIX_PPTX}"),
        os.path.join(directory, f"{stamp}{SUFFIX_DELETED}"),
    )


# --------------------------------------------------------------
def template_path(name):
    """Locate a skeleton template inside this skill."""
    return Path(__file__).resolve().parent.parent / name


# --------------------------------------------------------------
def render_skeleton(name, title):
    """Read a template and fill in the deck title."""
    path = template_path(name)
    if not path.is_file():
        print(f"ERROR: template missing: {path}",
              file=sys.stderr)
        sys.exit(1)
    text = path.read_text(encoding="utf-8")
    return text.replace("{{DECK_TITLE}}", title)


# --------------------------------------------------------------
def create_from_template(path, template, title):
    """Write a fresh file, refusing to overwrite."""
    if os.path.exists(path):
        return False
    directory = os.path.dirname(path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(render_skeleton(template, title))
    return True


# --------------------------------------------------------------
def parse_args():
    """Read the command line arguments."""
    parser = argparse.ArgumentParser(
        description="Resolve and bootstrap a weekly deck."
    )
    parser.add_argument(
        "date", nargs="?",
        help="seminar date as YYYY-MM-DD"
    )
    parser.add_argument(
        "--dir", default="work",
        help="directory holding the decks"
    )
    parser.add_argument(
        "--create", action="store_true",
        help="write a skeleton if the Markdown is missing"
    )
    return parser.parse_args()


# --------------------------------------------------------------
def main():
    """Report the deck paths, creating the file if asked."""
    args = parse_args()
    date = resolve_date(args.date)
    title = deck_title(date)
    md_path, pptx_path, del_path = deck_paths(
        date, args.dir
    )

    created = False
    if args.create:
        created = create_from_template(
            md_path, TEMPLATE, title
        )
        create_from_template(
            del_path, TEMPLATE_DELETED, title
        )

    print(json.dumps({
        "date": date.isoformat(),
        "weekday": date.strftime("%A"),
        "title": title,
        "md": md_path,
        "pptx": pptx_path,
        "deleted": del_path,
        "md_exists": os.path.isfile(md_path),
        "pptx_exists": os.path.isfile(pptx_path),
        "deleted_exists": os.path.isfile(del_path),
        "created": created,
    }))


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
