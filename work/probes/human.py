#!/usr/bin/env python3
"""
Imitate a human editing a live deck, for integration checks.

Each action is what the Slides UI does underneath, sent
through the API with no revision check, the way a person's
edit arrives.

Usage:
    python3 probes/human.py <DECK_ID> fill <SHAPE_ID>
    python3 probes/human.py <DECK_ID> type <SHAPE_ID> "text"
    python3 probes/human.py <DECK_ID> addbox <SLIDE_ID> "text"
    python3 probes/human.py <DECK_ID> move <SLIDE_ID> <INDEX>
    python3 probes/human.py <DECK_ID> copy <SHAPE_ID> <SLIDE_ID>
    python3 probes/human.py <DECK_ID> show <SLIDE_ID>

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

from layout import deck_layout as L
from gslides import deck as G
from gslides import api as S
from gslides.client import services

YELLOW = (0xFF, 0xF2, 0xCC)


# --------------------------------------------------------------
def send(slides, deck_id, requests):
    """Apply edits the way the UI would, unconditionally."""
    reply = slides.presentations().batchUpdate(
        presentationId=deck_id,
        body={"requests": requests}).execute()
    return reply.get("replies", [])


# --------------------------------------------------------------
def show(slides, deck_id, slide_id):
    """Print every shape on one slide with its state."""
    deck = G.read_deck(slides, deck_id)
    page = deck.slide(slide_id)
    print(f"{slide_id} at index {page.index} "
          f"(separator {deck.separator()})")
    for shape in page.shapes:
        line = shape.first_line()[:50]
        print(f"  {shape.id:32s} {shape.kind:5s} "
              f"filled={shape.filled!s:5s} "
              f"editable={deck.editable(shape)!s:5s} {line!r}")


# --------------------------------------------------------------
def copy_to(slides, deck_id, shape_id, slide_id):
    """Duplicate a shape, then move the copy to a slide."""
    reply = send(slides, deck_id, [{"duplicateObject": {
        "objectId": shape_id}}])
    new_id = reply[0]["duplicateObject"]["objectId"]
    deck = G.read_deck(slides, deck_id)
    text = deck.shape(new_id).text
    rect = L.Rect(0.5, 1.0, 5.0, 1.5)
    reqs = S.delete_object(new_id)
    reqs += S.textbox(new_id + "x", slide_id, rect)
    reqs += S.insert_text(new_id + "x", text)
    send(slides, deck_id, reqs)
    print(f"copied {shape_id} to {slide_id} as {new_id}x")


# --------------------------------------------------------------
def main():
    """Run one imitated human action."""
    deck_id, action = sys.argv[1], sys.argv[2]
    rest = sys.argv[3:]
    _, slides = services()
    if action == "fill":
        send(slides, deck_id, S.paint_box(rest[0], fill=YELLOW))
    elif action == "type":
        send(slides, deck_id, [{"deleteText": {
            "objectId": rest[0], "textRange": {"type": "ALL"}}}]
            + S.insert_text(rest[0], rest[1]))
    elif action == "addbox":
        send(slides, deck_id, [{"createShape": {
            "shapeType": "TEXT_BOX",
            "elementProperties": S.placement(
                rest[0], L.Rect(6.5, 4.6, 3.0, 0.6))}}])
        deck = G.read_deck(slides, deck_id)
        box = [s for s in deck.slide(rest[0]).shapes
               if not s.text and not s.id.startswith("s-")
               and not s.id.startswith("t-")]
        send(slides, deck_id, S.insert_text(box[-1].id, rest[1]))
        print("human box", box[-1].id)
    elif action == "move":
        send(slides, deck_id, [{"updateSlidesPosition": {
            "slideObjectIds": [rest[0]],
            "insertionIndex": int(rest[1])}}])
    elif action == "copy":
        copy_to(slides, deck_id, rest[0], rest[1])
    deck = G.read_deck(slides, deck_id)
    target = rest[1] if action == "copy" else rest[0]
    shape = deck.shape(target)
    show(slides, deck_id, shape.slide if shape else target)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
