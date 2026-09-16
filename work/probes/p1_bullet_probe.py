#!/usr/bin/env python3
"""
Probe how real Slides bullets behave through the API.

Creates a throwaway deck in the workflow folder, writes one
box with a headline and two bulleted lines, and prints what
Slides stores: the bullet glyph's style and the paragraph
indents. Tries making the glyph red by styling only the
newline that ends its paragraph. Deletes the deck at the end.

Usage:
    python3 probes/p1_bullet_probe.py

Created: 2026-09-16
Last updated: 2026-09-16
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

from gslides import api as S
from gslides.client import services, settings

TEXT = "Headline\nBlack text line\nhttps://example.com"
BOX = "probe-box-1"
DECK_TYPE = "application/vnd.google-apps.presentation"


# --------------------------------------------------------------
def make_deck(drive, slides):
    """A new empty deck in the folder; its id and slide id."""
    meta = {"name": "probe-bullets (delete me)",
            "mimeType": DECK_TYPE,
            "parents": [settings()["folder_id"]]}
    deck_id = drive.files().create(body=meta,
                                   fields="id").execute()["id"]
    pres = slides.presentations().get(
        presentationId=deck_id).execute()
    return deck_id, pres["slides"][0]["objectId"]


# --------------------------------------------------------------
def requests(slide_id):
    """Box, text, bullets on lines 2-3, red newline, indent."""
    reqs = [{"createShape": {
        "objectId": BOX, "shapeType": "TEXT_BOX",
        "elementProperties": {
            "pageObjectId": slide_id,
            "size": {
                "width": {"magnitude": 3e6, "unit": "EMU"},
                "height": {"magnitude": 1e6, "unit": "EMU"}},
        }}}]
    reqs += S.insert_text(BOX, TEXT)
    start, end = len("Headline\n"), len(TEXT)
    reqs.append({"createParagraphBullets": {
        "objectId": BOX,
        "textRange": {"type": "FIXED_RANGE",
                      "startIndex": start, "endIndex": end},
        "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"}})
    newline = len("Headline\nBlack text line")
    reqs += S.style_span(BOX, newline, newline + 1,
                         colour=(0xFF, 0, 0), size=14)
    reqs += S.hanging_indent(BOX, start, end)
    return reqs


# --------------------------------------------------------------
def report(slides, deck_id):
    """Print each paragraph's bullet and indents."""
    pres = slides.presentations().get(
        presentationId=deck_id).execute()
    for element in pres["slides"][0]["pageElements"]:
        if element["objectId"] != BOX:
            continue
        shape = element["shape"]
        for item in shape["text"]["textElements"]:
            marker = item.get("paragraphMarker")
            if marker:
                print(json.dumps(marker, indent=1))
        print(json.dumps(shape["text"].get("lists"), indent=1))


# --------------------------------------------------------------
def main():
    """Run the probe and always remove the deck."""
    drive, slides = services()
    deck_id, slide_id = make_deck(drive, slides)
    try:
        slides.presentations().batchUpdate(
            presentationId=deck_id,
            body={"requests": requests(slide_id)}).execute()
        report(slides, deck_id)
    finally:
        drive.files().delete(fileId=deck_id).execute()


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
