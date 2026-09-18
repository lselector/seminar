#!/usr/bin/env python3
"""
Write to a live deck without trampling a human's edits.

Three protections, described in ADD.md:

    Revision check  every batch carries the revision id the
                    plan was made from. If someone typed in
                    between, Google refuses the batch, and
                    run_plan reads the deck again and makes a
                    fresh plan, once.

    Ownership       the planning helpers below drop every
                    request aimed at a shape the rules say a
                    script may not touch. A plan cannot write
                    to a frozen box even by mistake.

    Geometry        refresh_text and refresh_picture keep a
                    shape's current position and width, so a
                    box a human moved stays where they put it.
                    Every text write sets the box height to
                    its text: exactly for a rewrite
                    (fit_rewrite), by the change for an edit
                    (fit_edit).

Planning helpers are pure: a Deck and requests in, requests
out. Only apply and run_plan talk to Google.

Usage:
    from gslides.write import run_plan, refresh_text
    def plan(deck):
        return refresh_text(deck, "s-toc-c0", body)
    run_plan(slides, deck_id, plan, "update TOC")

Created: 2026-09-14
Last updated: 2026-09-14
"""

import re

from layout import deck_layout as L
from gslides import render as R
from gslides import api as S
from gslides import ids as T
from gslides.client import log
from gslides.deck import MONTHS, read_deck

REVISION_CLASH = "does not match the latest revision"

# A date written the deck's way inside a box: "Sept 10",
# "Data for Sept 02". Group 2 is the day.
DATE_WORDS = re.compile(
    r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
    r"[a-z]*\.? (\d{1,2})\b", re.IGNORECASE)

# Where the target object id sits in each request type.
TARGET_KEYS = {
    "createShape": "objectId",
    "createImage": "objectId",
    "createSlide": "objectId",
    "insertText": "objectId",
    "deleteText": "objectId",
    "updateTextStyle": "objectId",
    "updateParagraphStyle": "objectId",
    "updateShapeProperties": "objectId",
    "updateImageProperties": "objectId",
    "updateTableCellProperties": "objectId",
    "updatePageElementAltText": "objectId",
    "updatePageElementTransform": "objectId",
    "deleteObject": "objectId",
    "replaceImage": "imageObjectId",
    "duplicateObject": "objectId",
}


class Conflict(Exception):
    """The deck changed after the plan was made."""


# --------------------------------------------------------------
def target(request):
    """The object id a request acts on, or None."""
    for kind, body in request.items():
        key = TARGET_KEYS.get(kind)
        if key:
            return body.get(key)
    return None


# --------------------------------------------------------------
def created(requests):
    """Ids of the shapes a list of requests creates."""
    return [target(r) for r in requests
            if "createShape" in r or "createImage" in r]


# --------------------------------------------------------------
def blocked_ids(deck, ids):
    """Of these ids, the ones that exist and may not change."""
    blocked = set()
    for object_id in ids:
        shape = deck.shape(object_id)
        if shape and not deck.editable(shape):
            blocked.add(object_id)
    return blocked


# --------------------------------------------------------------
def replace_owned(deck, slide_id, requests, sweep=True):
    """Rebuild a slide's script shapes, sparing frozen ones.

    Every shape the requests create is deleted first if it
    already exists and is editable. A frozen one keeps its
    old self, and every request aimed at it is dropped. With
    sweep, editable script shapes on the slide that the new
    requests no longer create are removed too.
    """
    new_ids = created(requests)
    blocked = blocked_ids(deck, new_ids)
    deletes = []
    for object_id in new_ids:
        if deck.shape(object_id) and object_id not in blocked:
            deletes += S.delete_object(object_id)
    home = deck.slide(slide_id)
    if sweep and home:
        for shape in home.shapes:
            if shape.id not in new_ids and \
                    deck.editable(shape):
                deletes += S.delete_object(shape.id)
    kept = [r for r in requests if target(r) not in blocked]
    return deletes + kept


# --------------------------------------------------------------
def height_request(shape, height):
    """Set a box's height, unless it is already that tall."""
    if height <= 0 or abs(height - shape.rect.h) < 0.02:
        return []
    return S.set_height(shape.id, shape.size_emu,
                        shape.transform, height)


# --------------------------------------------------------------
def fit_rewrite(shape, body):
    """Make a box exactly as tall as the body written into it."""
    return height_request(shape, R.box_height(
        R.body_paragraphs(body), shape.rect.w))


# --------------------------------------------------------------
def shape_paragraphs(shape, text=None):
    """(text, size, bold, bullet) per paragraph of a shape.

    text replaces the shape's own words; each paragraph keeps
    the style its position had.
    """
    text = shape.text if text is None else text
    styles = shape.styles or [(shape.size_pt or 12, False)]
    return [(line, *styles[min(i, len(styles) - 1)], False)
            for i, line in enumerate(text.split("\n"))]


# --------------------------------------------------------------
def fit_edit(shape, new_text):
    """Grow or shrink a box by what an edit did to its text.

    The change in measured text height is added to the box's
    height as it is. Padding and spacing set by hand, which
    the API cannot read, therefore stay as they were, and a
    box that fitted its text still does.
    """
    if not shape.rect:
        return []
    width = shape.rect.w
    new = R.box_height(shape_paragraphs(shape, new_text), width)
    change = new - R.box_height(shape_paragraphs(shape), width)
    floor = new - 2 * R.PAD_Y - R.TEXT_SLACK
    return height_request(shape, max(floor,
                                     shape.rect.h + change))


# --------------------------------------------------------------
def refresh_text(deck, box_id, body, fit=True):
    """Replace the words in a box and fit its height to them."""
    shape = deck.shape(box_id)
    if not shape or not deck.editable(shape):
        return []
    reqs = []
    if shape.text:
        reqs.append({"deleteText": {
            "objectId": box_id,
            "textRange": {"type": "ALL"},
        }})
    reqs += R.write_body(box_id, body)
    return reqs + (fit_rewrite(shape, body) if fit else [])


# --------------------------------------------------------------
def refresh_picture(deck, image_id, url, allow=()):
    """Swap a picture's content, keeping its frame."""
    shape = deck.shape(image_id)
    if not shape or not url:
        return []
    if not (deck.editable(shape) or image_id in allow):
        return []
    return [{"replaceImage": {
        "imageObjectId": image_id,
        "url": url,
        "imageReplaceMethod": "CENTER_INSIDE",
    }}]


# --------------------------------------------------------------
def room_rect(deck, box):
    """The box stretched down as far as free space allows."""
    rect = box.rect
    return L.Rect(rect.x, rect.y, rect.w,
                  deck.room_below(box))


# --------------------------------------------------------------
def topic_size(deck, box, block):
    """The font size a topic gets in its box."""
    return R.fit_size(block, room_rect(deck, box))


# --------------------------------------------------------------
def topic_body(deck, box, block):
    """The styled text of a topic sized for its box."""
    size = topic_size(deck, box, block)
    item = L.Item(block=block, size=size, text=box.rect)
    return R.block_body(item)


# --------------------------------------------------------------
def fit_height(deck, box, block):
    """Make a box as tall as its new text, within its room."""
    size = topic_size(deck, box, block)
    width = box.rect.w - 2 * R.PAD_X - L.TEXT_INSET
    needed = (R.height_in_box(block, width, size)
              + 2 * R.PAD_Y + R.TEXT_SLACK)
    height = min(needed, deck.room_below(box))
    if abs(height - box.rect.h) < 0.02:
        return []
    return S.set_height(box.id, box.size_emu, box.transform,
                        height)


# --------------------------------------------------------------
def refresh_topic(deck, key, block, url=None):
    """Rewrite a topic's box and picture where they stand.

    The box keeps its position and width. Its height follows
    the new text, growing only into free space below, so it
    never covers another shape; the font shrinks only when
    there is no room left to grow.
    """
    box_id = T.shape_id(key, T.BOX)
    box = deck.shape(box_id)
    if not box:
        return []
    reqs = refresh_text(deck, box_id, topic_body(deck, box,
                                                 block),
                        fit=False)
    if reqs:
        reqs += fit_height(deck, box, block)
    if url:
        reqs += refresh_picture(
            deck, T.shape_id(key, T.PICTURE), url)
    return reqs


# --------------------------------------------------------------
def replace_line(deck, box_id, marker, new_line, allow=()):
    """Swap the one line of a box that contains marker.

    Only that line changes; the rest of the box, and the
    style of the line itself, stay as they are. The box
    height follows the change. A box in allow is rewritten
    even when a human has filled it: the step was told to
    keep that one line current.
    """
    shape = deck.shape(box_id)
    if not shape or not (deck.editable(shape)
                         or box_id in allow):
        return []
    lines = shape.text.split("\n")
    start = 0
    for index, line in enumerate(lines):
        if marker in line:
            if line == new_line:
                return []
            end = start + R.u16(line)
            new_text = "\n".join(lines[:index] + [new_line]
                                 + lines[index + 1:])
            return [{"deleteText": {
                "objectId": box_id,
                "textRange": {"type": "FIXED_RANGE",
                              "startIndex": start,
                              "endIndex": end}}},
                    {"insertText": {"objectId": box_id,
                                    "insertionIndex": start,
                                    "text": new_line}}] \
                + fit_edit(shape, new_text)
        start += R.u16(line) + 1
    return []


# --------------------------------------------------------------
def date_label(day, padded=False):
    """A date the deck's way: Sept 14, or Sept 04 if padded."""
    number = f"{day.day:02d}" if padded else str(day.day)
    return f"{MONTHS[day.month - 1]} {number}"


# --------------------------------------------------------------
def date_edit(text, day):
    """(start, end, new) that sets the date in text, or None.

    Only "Month day" changes; the words around it stay. A
    zero-padded day (Sept 02) stays zero-padded.
    """
    match = DATE_WORDS.search(text)
    if not match:
        return None
    old_day = match.group(2)
    new = date_label(day, padded=len(old_day) == 2
                     and old_day.startswith("0"))
    return (match.start(), match.end(), new)


# --------------------------------------------------------------
def replace_spans(box_id, text, edits):
    """Swap spans of a box's text, keeping each span's style.

    edits are (start, end, new) in string positions of text.
    Each new piece goes in after the old one, taking its
    style, and then the old one is deleted. Edits run from
    the end of the text backwards so positions stay valid.
    """
    reqs = []
    for start, end, new in sorted(edits, reverse=True):
        if text[start:end] == new:
            continue
        first, last = R.u16(text[:start]), R.u16(text[:end])
        reqs += [{"insertText": {"objectId": box_id,
                                 "insertionIndex": last,
                                 "text": new}},
                 {"deleteText": {"objectId": box_id,
                                 "textRange": {
                                     "type": "FIXED_RANGE",
                                     "startIndex": first,
                                     "endIndex": last}}}]
    return reqs


# --------------------------------------------------------------
def apply_edits(text, edits):
    """What text becomes after (start, end, new) edits."""
    for start, end, new in sorted(edits, reverse=True):
        text = text[:start] + new + text[end:]
    return text


# --------------------------------------------------------------
def edit_spans(shape, edits):
    """Swap spans of a box's text and fit its height."""
    reqs = replace_spans(shape.id, shape.text, edits)
    if not reqs:
        return []
    return reqs + fit_edit(shape, apply_edits(shape.text,
                                              edits))


# --------------------------------------------------------------
def replace_date(shape, day):
    """Swap the date inside a box's text, keeping its style."""
    edit = date_edit(shape.text, day) if shape else None
    return edit_spans(shape, [edit]) if edit else []


# --------------------------------------------------------------
def guard(deck, requests, allow=()):
    """Last line of defence: drop writes to protected shapes.

    Creating a brand-new id is always allowed. Anything
    aimed at an existing shape must pass deck.editable, or
    be in allow: shapes a step was explicitly told to change.
    """
    kept, dropped = [], []
    fresh = set(created(requests))
    for request in requests:
        object_id = target(request)
        shape = deck.shape(object_id) if object_id else None
        page = deck.slide(object_id or "")
        if object_id in allow and shape is not None:
            allowed = True
        elif page is not None:
            allowed = (T.owned(object_id)
                       and page.index < deck.separator())
        else:
            allowed = (shape is None or object_id in fresh
                       or deck.editable(shape))
        if allowed:
            kept.append(request)
        else:
            dropped.append(object_id)
    return kept, sorted(set(dropped))


# --------------------------------------------------------------
def apply(slides, deck, requests, label, allow=()):
    """Send one batch at the deck's revision."""
    from googleapiclient.errors import HttpError
    kept, dropped = guard(deck, requests, allow)
    if dropped:
        log(f"  guard kept {label} off: {dropped}")
    if not kept:
        return 0
    control = {"requiredRevisionId": deck.revision}
    body = {"requests": kept, "writeControl": control}
    try:
        reply = slides.presentations().batchUpdate(
            presentationId=deck.id, body=body).execute()
    except HttpError as exc:
        if REVISION_CLASH in str(exc):
            raise Conflict(label) from exc
        raise
    deck.revision = reply.get("writeControl", {}).get(
        "requiredRevisionId", deck.revision)
    return len(kept)


# --------------------------------------------------------------
def run_plan(slides, deck_id, plan, label, tries=2, allow=()):
    """Read, plan and write, re-planning once on a clash.

    plan(deck) returns a list of (label, requests) batches.
    Batches go in order, each at the revision the previous
    one produced. allow lists shapes the step may change
    even though a human made them.
    """
    for attempt in range(1, tries + 1):
        deck = read_deck(slides, deck_id)
        batches = plan(deck)
        try:
            sent = 0
            for part, requests in batches:
                sent += apply(slides, deck, requests, part,
                              allow)
            return sent
        except Conflict:
            log(f"  {label}: the deck changed while writing "
                f"(attempt {attempt}), reading it again")
    log(f"ERROR: {label}: the deck kept changing. Nothing "
        f"more was written. Try again in a moment.")
    return -1
