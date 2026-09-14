#!/usr/bin/env python3
"""
Export a live deck to PNG pages, to look at what Google drew.

Read-only. Exports through Drive as PDF, then rasterizes each
page with ImageMagick into /tmp/gcheck/<prefix>-NN.png.

Usage:
    python3 probes/render_deck.py <DECK_ID> [prefix]

Created: 2026-09-14
Last updated: 2026-09-14
"""

import glob
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

from gslides.client import services

OUT = "/tmp/gcheck"


# --------------------------------------------------------------
def main():
    """Export one deck and split it into page images."""
    deck_id = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else "page"
    os.makedirs(OUT, exist_ok=True)
    for old in glob.glob(os.path.join(OUT, f"{prefix}-*.png")):
        os.remove(old)
    drive, _ = services()
    pdf = os.path.join(OUT, f"{prefix}.pdf")
    with open(pdf, "wb") as handle:
        handle.write(drive.files().export(
            fileId=deck_id, mimeType="application/pdf"
        ).execute())
    subprocess.run(["magick", "-density", "110", pdf,
                    "-background", "white", "-alpha", "remove",
                    os.path.join(OUT, f"{prefix}-%02d.png")],
                   check=True)
    pages = sorted(glob.glob(os.path.join(OUT,
                                          f"{prefix}-*.png")))
    print(f"{len(pages)} pages in {OUT}/{prefix}-NN.png")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
