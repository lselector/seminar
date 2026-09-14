#!/usr/bin/env python3
"""
Hand-edit probes behind ADD.md: fill, cut and copy by hand.

The API probes showed what a fill or a duplicate looks like
when a script makes it. These confirm the same holds when a
person does it in the Slides UI in a browser.

    setup   make a scratch deck with three labelled boxes
    check   read it back and report what the UI did
    clean   delete the scratch deck

In the deck, by hand:

    box 1  fill it with any colour
    box 2  select it, cut (Cmd+X), paste (Cmd+V)
    box 3  select it, copy (Cmd+C), paste (Cmd+V)

Usage:
    python3 probes/p0_ui_probe.py setup
    python3 probes/p0_ui_probe.py check
    python3 probes/p0_ui_probe.py clean

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

from layout import deck_layout as L
from gslides import deck as G
from gslides import render as R
from gslides import api as S
from gslides.client import services, settings

TAG = {"seminarProbe": "ui"}
BOXES = [
    ("t-probe-fill-b", "1. Fill this box with any colour"),
    ("t-probe-cut-b", "2. Cut this box (Cmd+X), then paste it"),
    ("t-probe-copy-b",
     "3. Copy this box (Cmd+C), then paste it"),
]


# --------------------------------------------------------------
def find(drive):
    """The scratch probe deck, or None."""
    got = drive.files().list(
        q="appProperties has { key='seminarProbe' and "
          "value='ui' } and trashed=false",
        fields="files(id,webViewLink)").execute()["files"]
    return got[0] if got else None


# --------------------------------------------------------------
def setup(drive, slides):
    """Create the scratch deck with three boxes."""
    if find(drive):
        print("Already set up:", find(drive)["webViewLink"])
        return
    made = drive.files().create(body={
        "name": "PROBE - three edits by hand, then tell Claude",
        "mimeType": G.DECK_MIME, "appProperties": TAG,
        "parents": [settings()["folder_id"]],
    }, fields="id,webViewLink").execute()
    reqs = S.new_slide("s-probe-ui")
    for n, (box, text) in enumerate(BOXES):
        rect = L.Rect(0.6, 0.6 + n * 1.5, 6.0, 0.8)
        reqs += R.make_box(box, "s-probe-ui", rect, border=True)
        reqs += R.write_body(box, R.plain_body(text, 20))
    slides.presentations().batchUpdate(
        presentationId=made["id"],
        body={"requests": reqs}).execute()
    print("Open this, make the three edits, then run check:")
    print(made["webViewLink"])


# --------------------------------------------------------------
def check(drive, slides):
    """Report what the three hand edits look like."""
    found = find(drive)
    if not found:
        print("No probe deck. Run setup first.")
        return
    deck = G.read_deck(slides, found["id"])
    shapes = {s.id: s for sl in deck.slides for s in sl.shapes}
    fill = shapes.get("t-probe-fill-b")
    print(f"a. filled by hand reads as filled: "
          f"{bool(fill and fill.filled)}")
    for box, text in BOXES[1:]:
        same = [s for s in shapes.values()
                if s.first_line() == text]
        ids = sorted(s.id for s in same)
        new = [i for i in ids if i != box]
        print(f"b. {text.split('(')[0].strip()}: ids {ids}; "
              f"pasted box has a new id: {bool(new)}")


# --------------------------------------------------------------
def main():
    """Run setup, check or clean."""
    action = sys.argv[1] if len(sys.argv) > 1 else "check"
    drive, slides = services()
    if action == "setup":
        setup(drive, slides)
    elif action == "clean":
        found = find(drive)
        if found:
            drive.files().delete(fileId=found["id"]).execute()
        print("removed" if found else "nothing to remove")
    else:
        check(drive, slides)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
