#!/usr/bin/env python3
"""
Step 3: rebuild the table of contents on the first slide.

Walks the talk from the first slide to the separator, in the
order the slides are in right now, and collects the first
line of every text box: the script's and yours alike. Titles
and empty placeholders are skipped, and so are the slides
the archive decks leave out of the contents: the script's
Benchmarks page, the YouTube promo, About the Speaker and
Thank You. On your own benchmarks page (the Code | Model |
Score tables) only the board captions count, the boxes with
an arena.ai leaderboard link.

Each item is a short label on one line: a bold blue bulleted
list with no links or details, in four boxes, two a side:
left upper light yellow, left lower light green, right upper
light blue, right lower light yellow. Items run in slide
order and fill one box before the next, in that order; a box
holds 8 to 12 lines, as many as fit its side, and an empty
box shows "xxx". The left side starts under the title; the
right side starts under the epigraph at top right, so the
two never overlap. A skill can
pass labels for headlines; "-" leaves a headline out, and
headlines given the same label are listed once:

    {"toc": {"English - https://lmarena.ai/...": "LM Arena",
             "Sept 10": "-"},
     "epigraph": "Agents got their own computers."}

Labels are kept in the contents slide's speaker notes, so a
later plain run uses them too. A headline with no label has
its links and trailing marks removed and is cut at a word
until it fits one line.

The script paints those four colours itself, so a fill does
not freeze these boxes the way it freezes others. Instead,
give any contents box a different colour, or clear its fill,
and the whole contents is yours and left alone, so the boxes
never disagree with each other.
The epigraph is written only while its box still holds its
placeholder.

Usage:
    python3 g3_update_toc.py                  this week's Friday
    python3 g3_update_toc.py 2026-09-18
    python3 g3_update_toc.py --json toc.json  labels, epigraph

Created: 2026-09-14
Last updated: 2026-09-19
"""

import re

from layout import deck_layout as L
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import api as S
from gslides import bench as BK
from gslides import ids as T
from gslides.client import log
from gslides.step import load_json, page_or_stop, start

LABEL = "table of contents"
MAX_COLUMNS = 4
EPIGRAPH_BOX = T.shape_id(T.PAGE_TOC, "epi")
LABELS_HEAD = "Contents labels. Kept by g3_update_toc.py."
LABEL_SEP = " => "
LEAVE_OUT = "-"
URL = re.compile(r"https?://\S+")
TRIM = " -–—:|,;.●"
DANGLING = {"a", "an", "and", "as", "at", "by", "for", "from",
            "in", "of", "on", "or", "the", "to", "with"}

# On your benchmarks page only the board captions count
# ("English - https://lmarena.ai/..."); the Elo note, the date
# and the model sizes are not topics.
ARENA_LINK = "arena.ai/leaderboard"

# Pages the archive decks keep out of the contents.
NOT_IN_TOC = {T.PAGE_TOC, T.PAGE_BENCH, T.PAGE_YOUTUBE,
              T.PAGE_ABOUT, T.PAGE_THANKS, T.PAGE_SEPARATOR,
              T.PAGE_PARKED}


# --------------------------------------------------------------
def column_ids():
    """Every id a contents box may have: c0 to c3."""
    return [T.shape_id(T.PAGE_TOC, f"c{i}")
            for i in range(MAX_COLUMNS)]


# --------------------------------------------------------------
def is_yours(shape):
    """A contents box you recoloured or cleared of its fill.

    The script fills each box with its own colour; a box
    showing any other colour, or none after the script
    painted it, is the author's. An unfilled box from an
    older deck is still the script's.
    """
    wanted = dict(zip(column_ids(), R.TOC_FILLS))
    if not shape.filled:
        return False
    return shape.fill != wanted.get(shape.id)


# --------------------------------------------------------------
def headline(shape):
    """The line a box contributes to the contents."""
    line = shape.first_line()
    if line.startswith(R.BULLET.strip()):
        line = line[len(R.BULLET.strip()):].strip()
    return line


# --------------------------------------------------------------
def shorten(line):
    """A one-line item: no links, cut at a word to fit."""
    words = URL.sub(" ", line).replace("●", " ").split()
    text = " ".join(words).strip(TRIM)
    while len(words) > 1 and not R.toc_fits(text):
        words.pop()
        while len(words) > 1 and \
                words[-1].strip(TRIM).lower() in DANGLING:
            words.pop()
        text = " ".join(words).strip(TRIM)
    return text


# --------------------------------------------------------------
def read_labels(deck):
    """Headline-to-label pairs from the TOC notes; last wins."""
    page = deck.slide(T.PAGE_TOC)
    labels = {}
    for line in (page.notes if page else "").split("\n"):
        if LABEL_SEP in line:
            key, label = line.split(LABEL_SEP, 1)
            labels[key.strip()] = label.strip()
    return labels


# --------------------------------------------------------------
def title_bottom(slide):
    """Where the slide's own title ends, or 0 if it has none.

    The title is the script's -title box, else the topmost
    box in the title band.
    """
    titles = [s for s in slide.text_shapes() if G.is_title(s)]
    titles.sort(key=lambda s: (not s.id.endswith("-title"),
                               s.rect.y))
    if not titles:
        return 0.0
    return titles[0].rect.y + titles[0].rect.h


# --------------------------------------------------------------
def in_title_row(shape, bottom):
    """A title, or a box beside it: not a topic.

    A box of yours that starts high up but below the title's
    bottom edge is content, such as the notes box under the
    Intelligence Index title.
    """
    return G.is_title(shape) and shape.rect.y < bottom - 0.01


# --------------------------------------------------------------
def headlines(deck):
    """First lines of the talk, in the current slide order."""
    lines = []
    bench = BK.table_page(deck)
    for slide in deck.main():
        if slide.id in NOT_IN_TOC:
            continue
        bottom = title_bottom(slide)
        for shape in slide.text_shapes():
            if in_title_row(shape, bottom) or \
                    G.is_placeholder(shape.text):
                continue
            if bench and slide.id == bench.id and \
                    ARENA_LINK not in shape.text:
                continue
            line = headline(shape)
            if line and line not in lines:
                lines.append(line)
    return lines


# --------------------------------------------------------------
def collect(deck, labels=None):
    """Contents items of the talk, in the current order."""
    labels = labels or {}
    names = []
    for line in headlines(deck):
        label = labels.get(line, line).strip() or LEAVE_OUT
        if label == LEAVE_OUT:
            continue
        name = shorten(label)
        if name and name not in names:
            names.append(name)
    return names


# --------------------------------------------------------------
def label_requests(deck, given):
    """Append new or changed labels to the TOC notes."""
    page = deck.slide(T.PAGE_TOC)
    known = read_labels(deck)
    lines = [f"{k}{LABEL_SEP}{v.strip() or LEAVE_OUT}"
             for k, v in given.items()
             if known.get(k) != (v.strip() or LEAVE_OUT)]
    if not lines:
        return []
    if not page or not page.notes_id:
        log(f"{LABEL}: no notes on the TOC slide, labels not "
            f"kept")
        return []
    text = "\n".join(lines)
    head = "\n" if page.notes.strip() else LABELS_HEAD + "\n"
    return [{"insertText": {
        "objectId": page.notes_id,
        "insertionIndex": R.u16(page.notes),
        "text": head + text,
    }}]


# --------------------------------------------------------------
def right_top(deck, epigraph=None):
    """Where the right side starts: under the epigraph.

    The box's own height, or the height its text needs if
    that is more, so a long epigraph pushes the side down.
    An epigraph written in this same run counts.
    """
    epi = deck.shape(EPIGRAPH_BOX)
    if not epi:
        return L.BAND_TOP
    text = epigraph if (epigraph and
                        G.is_placeholder(epi.text)) else epi.text
    tall = max(epi.rect.h, R.epigraph_rect(text).h)
    return max(L.BAND_TOP, epi.rect.y + tall + 0.10)


# --------------------------------------------------------------
def toc_requests(deck, names, epigraph=None):
    """Rebuild the four boxes, or nothing if any is yours."""
    existing = [deck.shape(i) for i in column_ids()
                if deck.shape(i)]
    yours = [s.id for s in existing if is_yours(s)]
    if yours:
        log(f"{LABEL}: left alone, you recoloured "
            f"{', '.join(yours)}")
        return []
    sid = T.PAGE_TOC
    tops = (L.BAND_TOP, right_top(deck, epigraph))
    wanted = R.render_toc_boxes(sid, sid, names, tops)
    reqs = W.replace_owned(deck, sid, wanted, sweep=False,
                           allow=column_ids())
    new_ids = set(W.created(wanted))
    for shape in existing:
        if shape.id not in new_ids:
            reqs = S.delete_object(shape.id) + reqs
    return reqs


# --------------------------------------------------------------
def epigraph_requests(deck, text):
    """Write the epigraph into a still-empty epigraph box."""
    box = deck.shape(EPIGRAPH_BOX)
    if not text or not box:
        return []
    if not G.is_placeholder(box.text):
        log(f"{LABEL}: epigraph already written, kept")
        return []
    body = R.plain_body(text, R.EPIGRAPH_SIZE, colour=R.RED,
                        bold=True, italic=True)
    return W.refresh_text(deck, EPIGRAPH_BOX, body)


# --------------------------------------------------------------
def log_items(deck, labels):
    """Show each headline and the item it became."""
    for line in headlines(deck):
        label = labels.get(line, line).strip() or LEAVE_OUT
        item = "(left out)" if label == LEAVE_OUT \
            else shorten(label)
        log(f"    {item}" if item == line
            else f"    {item}  <=  {line}")


# --------------------------------------------------------------
def plan_for(extra):
    """Build the planning function for run_plan."""
    given = extra.get("toc") or {}

    # --------------------------------------
    def plan(deck):
        """Contents, labels and, if given, the epigraph."""
        if not page_or_stop(deck, T.PAGE_TOC, LABEL):
            return []
        labels = {**read_labels(deck), **given}
        names = collect(deck, labels)
        log(f"{LABEL}: {len(names)} items")
        log_items(deck, labels)
        return [("epigraph",
                 epigraph_requests(deck, extra.get("epigraph"))),
                ("labels", label_requests(deck, given)),
                (LABEL, toc_requests(deck, names,
                                     extra.get("epigraph")))]

    return plan


# --------------------------------------------------------------
def main():
    """Rebuild the contents of one deck."""
    step = start("Update the table of contents")
    extra = load_json(step.args.json) if step.args.json else {}
    sent = W.run_plan(step.slides, step.deck_id,
                      plan_for(extra), LABEL,
                      allow=tuple(column_ids()))
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
