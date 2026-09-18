#!/usr/bin/env python3
"""
The standing wording of the fixed slides.

config/skeleton.json holds the words of the pages that do not
come from the news: the Intelligence Index page, the channel
promo, layoffs, About and Thank You. This module is the one
way to read that file and turn a page of it into a Section
the layout can paginate.

Step 1 builds every fixed page from here. A later step uses
it to draw one of those boxes again if it goes missing.

Usage:
    from layout import skeleton as K
    entry = K.page("youtube")
    section = K.section(entry, ROOT)

Created: 2026-09-17
Last updated: 2026-09-17
"""

import json
import os

from layout.deck_parser import Block, Section

WORK = os.path.dirname(os.path.dirname(os.path.abspath(
    __file__)))
SKELETON = os.path.join(WORK, "config", "skeleton.json")


# --------------------------------------------------------------
def load():
    """Every fixed page, as config/skeleton.json holds it."""
    with open(SKELETON, encoding="utf-8") as handle:
        return json.load(handle)


# --------------------------------------------------------------
def page(name):
    """One fixed page by its key, such as "youtube"."""
    return load()[name]


# --------------------------------------------------------------
def section(entry, root=None):
    """One page as a Section of Blocks.

    A page kind (promo, profile, closing) is set on the
    section and on each of its blocks, as the layout reads
    it from both. Image paths are joined to root.
    """
    flags = {entry["kind"]: True} if entry.get("kind") else {}
    blocks = []
    for topic in entry["topics"]:
        image = topic.get("image")
        blocks.append(Block(
            headline=topic["headline"],
            bullets=list(topic["bullets"]),
            image=os.path.join(root, image)
            if (image and root) else image,
            **flags))
    return Section(title=entry["title"], blocks=blocks,
                   **flags)
