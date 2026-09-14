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
                    shape's current position and size, so a
                    box a human moved stays where they put it.

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

from layout import deck_layout as L
from gslides import render as R
from gslides import api as S
from gslides import ids as T
from gslides.client import log
from gslides.deck import read_deck

REVISION_CLASH = "does not match the latest revision"

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
def refresh_text(deck, box_id, body):
    """Replace the words in a box, keeping its place."""
    shape = deck.shape(box_id)
    if not shape or not deck.editable(shape):
        return []
    reqs = []
    if shape.text:
        reqs.append({"deleteText": {
            "objectId": box_id,
            "textRange": {"type": "ALL"},
        }})
    return reqs + R.write_body(box_id, body)


# --------------------------------------------------------------
def refresh_picture(deck, image_id, url):
    """Swap a picture's content, keeping its frame."""
    shape = deck.shape(image_id)
    if not shape or not url or not deck.editable(shape):
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
                                                 block))
    if reqs:
        reqs += fit_height(deck, box, block)
    if url:
        reqs += refresh_picture(
            deck, T.shape_id(key, T.PICTURE), url)
    return reqs


# --------------------------------------------------------------
def replace_line(deck, box_id, marker, new_line):
    """Swap the one line of a box that contains marker.

    Only that line changes; the rest of the box, and the
    style of the line itself, stay as they are.
    """
    shape = deck.shape(box_id)
    if not shape or not deck.editable(shape):
        return []
    start = 0
    for line in shape.text.split("\n"):
        if marker in line:
            if line == new_line:
                return []
            end = start + R.u16(line)
            return [{"deleteText": {
                "objectId": box_id,
                "textRange": {"type": "FIXED_RANGE",
                              "startIndex": start,
                              "endIndex": end}}},
                    {"insertText": {"objectId": box_id,
                                    "insertionIndex": start,
                                    "text": new_line}}]
        start += R.u16(line) + 1
    return []


# --------------------------------------------------------------
def guard(deck, requests):
    """Last line of defence: drop writes to protected shapes.

    Creating a brand-new id is always allowed. Anything
    aimed at an existing shape must pass deck.editable.
    """
    kept, dropped = [], []
    fresh = set(created(requests))
    for request in requests:
        object_id = target(request)
        shape = deck.shape(object_id) if object_id else None
        page = deck.slide(object_id or "")
        if page is not None:
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
def apply(slides, deck, requests, label):
    """Send one batch at the deck's revision."""
    from googleapiclient.errors import HttpError
    kept, dropped = guard(deck, requests)
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
def run_plan(slides, deck_id, plan, label, tries=2):
    """Read, plan and write, re-planning once on a clash.

    plan(deck) returns a list of (label, requests) batches.
    Batches go in order, each at the revision the previous
    one produced.
    """
    for attempt in range(1, tries + 1):
        deck = read_deck(slides, deck_id)
        batches = plan(deck)
        try:
            sent = 0
            for part, requests in batches:
                sent += apply(slides, deck, requests, part)
            return sent
        except Conflict:
            log(f"  {label}: the deck changed while writing "
                f"(attempt {attempt}), reading it again")
    log(f"ERROR: {label}: the deck kept changing. Nothing "
        f"more was written. Try again in a moment.")
    return -1
