#!/usr/bin/env python3
"""
Step 2: add this week's news topics to a live deck.

A skill finds the news and writes it as JSON; this script
decides what is new, lays it out and puts it in the deck:

    {"topics": [
      {"headline": "Mistral raises EUR 3 billion",
       "bullets": ["Series D led by ...", "..."],
       "url": "https://source.example.com/article",
       "image": {"shot": "https://source.example.com/article"}}
    ]}

A topic is skipped when any of these is true:

    its headline matches the first line of any text box
    anywhere in the deck, parked section included, so a
    topic you copied past the separator never comes back

    it is in the topic ledger: the speaker notes of the
    separator slide, where every topic this script ever
    added is recorded, so a topic you deleted outright
    never comes back either

    it repeats another topic in the same JSON

Matching is on the topic, not the exact wording (the same
same_topic rule the Markdown workflow uses).

New topics fill the "AI News" placeholder slides first, while
their placeholder is still empty and nobody else has put
anything on them. The rest go in as new slides just before
the layoffs slide, wherever it is now.

Usage:
    python3 g2_add_news.py --list              what is known
    python3 g2_add_news.py --json news.json --dry-run
    python3 g2_add_news.py --json news.json
    python3 g2_add_news.py 2026-09-18 --json news.json
    python3 g2_add_news.py --json news.json --redraw

--redraw draws topics added before again, in the current box
format. It touches a topic only when this script added it
(it is in the ledger), its box is still unfilled and in the
talk, and its slide holds nothing but such boxes. Those
slides are removed and the topics laid out again in their
place, with the same ids, so the ledger stays as it is.

Created: 2026-09-14
Last updated: 2026-10-20
"""

import sys

from layout import deck_layout as L
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import api as S
from gslides import ids as T
from layout.deck_parser import Block, Section, same_topic
from gslides.client import log
from gslides.host import Host
from sources.pictures import prepare, remove_workdir, workdir
from gslides.step import load_json, start

LABEL = "news"
NEWS_TITLE = "AI News"
LEDGER_HEAD = "Topic ledger. Kept by g2_add_news.py."
LEDGER_SEP = " | "

PLACEHOLDER_PAGES = [(T.PAGE_NEWS_1, T.TOPIC_NEWS_1),
                     (T.PAGE_NEWS_2, T.TOPIC_NEWS_2)]
ANCHORS = [T.PAGE_LAYOFFS, T.PAGE_ABOUT, T.PAGE_SEPARATOR]
SKIP_SLIDES = {T.PAGE_TOC, T.PAGE_SEPARATOR}


# --------------------------------------------------------------
def to_topic(item):
    """One JSON topic as a Block and a picture spec."""
    bullets = [b for b in item.get("bullets", []) if b.strip()]
    url = item.get("url")
    if url and url not in bullets:
        bullets.append(url)
    spec = item.get("image") or ({"shot": url} if url else None)
    return Block(headline=item["headline"].strip(),
                 bullets=bullets), spec


# --------------------------------------------------------------
def read_ledger(deck):
    """(key, headline) pairs from the separator's notes."""
    sep = deck.slide(T.PAGE_SEPARATOR)
    entries = []
    for line in (sep.notes if sep else "").split("\n"):
        if line.startswith(f"{T.OURS}-") and LEDGER_SEP in line:
            key, headline = line.split(LEDGER_SEP, 1)
            entries.append((key.strip(), headline.strip()))
    return entries


# --------------------------------------------------------------
def deck_headlines(deck):
    """First lines of every content box, parked included."""
    found = []
    for slide in deck.slides:
        if slide.id in SKIP_SLIDES:
            continue
        for shape in slide.text_shapes():
            if G.is_title(shape) or G.is_placeholder(shape.text):
                continue
            found.append((shape.first_line(), slide.id))
    return found


# --------------------------------------------------------------
def why_known(block, deck, ledger, chosen):
    """The reason a topic is not new, or None."""
    key = T.block_key(block)
    if key in {k for k, _ in ledger}:
        return "added before (ledger)"
    for line, slide_id in deck_headlines(deck):
        if same_topic(block.headline, line):
            where = ("parked" if deck.slide(slide_id).index
                     > deck.separator() else "in the deck")
            return f"{where}: {line!r}"
    for _, headline in ledger:
        if same_topic(block.headline, headline):
            return f"added before as {headline!r}"
    for other in chosen:
        if same_topic(block.headline, other.headline):
            return "repeats another topic in this batch"
    return None


# --------------------------------------------------------------
def select(topics, deck):
    """Split candidates into new ones and skipped ones."""
    ledger = read_ledger(deck)
    fresh, skipped = [], []
    for block, spec in topics:
        reason = why_known(block, deck, ledger,
                           [b for b, _ in fresh])
        if reason:
            skipped.append((block.headline, reason))
        else:
            fresh.append((block, spec))
    return fresh, skipped


# --------------------------------------------------------------
def free_placeholder(deck, slide_id, key):
    """Is this AI News slide still just its placeholder?"""
    page = deck.slide(slide_id)
    box = deck.shape(T.shape_id(key, T.BOX))
    if not page or not box or page.index >= deck.separator():
        return False
    if not deck.editable(box) or not G.is_placeholder(box.text):
        return False
    others = [s for s in page.shapes
              if s.id != box.id and not G.is_title(s)]
    return not others


# --------------------------------------------------------------
def anchor_index(deck):
    """Where new slides go: before layoffs, or a fallback."""
    for slide_id in ANCHORS:
        page = deck.slide(slide_id)
        if page and page.index <= deck.separator():
            return page.index
    return deck.separator()


# --------------------------------------------------------------
def fill_placeholder(deck, slide_id, page, keys, locate):
    """Put a laid-out page onto a placeholder slide."""
    wanted = R.render_content(slide_id, slide_id, page, keys,
                              fill=None, locate=locate)
    return W.replace_owned(deck, slide_id, wanted)


# --------------------------------------------------------------
def insert_page(slide_id, index, page, keys, locate):
    """A new slide at a given place, with its content."""
    head = S.new_slide(slide_id)
    head[0]["createSlide"]["insertionIndex"] = index
    return head + R.render_content(slide_id, slide_id, page,
                                   keys, fill=None,
                                   locate=locate)


# --------------------------------------------------------------
def ledger_requests(deck, blocks, keys):
    """Record the added topics in the separator's notes."""
    sep = deck.slide(T.PAGE_SEPARATOR)
    if not sep or not sep.notes_id or not blocks:
        log(f"{LABEL}: no separator slide, ledger not kept")
        return []
    lines = [f"{k}{LEDGER_SEP}{b.headline}"
             for b, k in zip(blocks, keys)]
    text = "\n".join(lines)
    if sep.notes.strip():
        text = "\n" + text
    else:
        text = LEDGER_HEAD + "\n" + text
    return [{"insertText": {
        "objectId": sep.notes_id,
        "insertionIndex": R.u16(sep.notes),
        "text": text,
    }}]


# --------------------------------------------------------------
def place(deck, pages, locate):
    """Batches that put every laid-out page in the deck."""
    taken = set(deck.ids())
    slots = [sid for sid, key in PLACEHOLDER_PAGES
             if free_placeholder(deck, sid, key)]
    anchor, inserted = anchor_index(deck), 0
    batches, all_keys = [], []
    for page in pages:
        keys = [T.make_unique(T.block_key(i.block), taken)
                for i in page.items]
        all_keys += keys
        if slots:
            sid = slots.pop(0)
            reqs = fill_placeholder(deck, sid, page, keys,
                                    locate)
        else:
            sid = T.make_unique(T.page_key(page), taken)
            reqs = insert_page(sid, anchor + inserted, page,
                               keys, locate)
            inserted += 1
        batches.append((sid, reqs))
        log(f"  {sid}: {len(page.items)} topic(s)")
    return batches, all_keys


# --------------------------------------------------------------
def report(fresh, skipped, verb="new"):
    """Say what will be added and what was skipped."""
    for block, _ in fresh:
        log(f"{LABEL}: {verb:8s} {block.headline}")
    for headline, reason in skipped:
        log(f"{LABEL}: skipped  {headline}  ({reason})")


# --------------------------------------------------------------
def plan_for(topics, locate):
    """Build the planning function for run_plan."""

    # --------------------------------------
    def plan(deck):
        """Choose, lay out, place, and record."""
        fresh, skipped = select(topics, deck)
        report(fresh, skipped)
        blocks = [b for b, _ in fresh]
        if not blocks:
            return []
        section = Section(title=NEWS_TITLE, blocks=blocks)
        pages = L.paginate([section])
        batches, keys = place(deck, pages, locate)
        placed = [i.block for p in pages for i in p.items]
        batches.append(("ledger",
                        ledger_requests(deck, placed, keys)))
        return batches

    return plan


# --------------------------------------------------------------
def redraw_blocker(deck, key, ledger):
    """Why a topic's box cannot be redrawn, or None."""
    if key not in ledger:
        return "not added by this script (not in the ledger)"
    box = deck.shape(T.shape_id(key, T.BOX))
    if not box:
        return "its box is gone (deleted or renamed)"
    if not deck.editable(box):
        return "filled by you, or parked"
    if box.slide in {sid for sid, _ in PLACEHOLDER_PAGES}:
        return "on a placeholder slide, which is never removed"
    return None


# --------------------------------------------------------------
def slide_is_ours(deck, slide_id, keys):
    """Does a slide hold only the script's shapes of keys?"""
    slide = deck.slide(slide_id)
    if not slide or not T.owned(slide_id):
        return False
    for shape in slide.shapes:
        if shape.id.endswith(f"-{T.TITLE}"):
            continue
        if not deck.editable(shape) or \
                T.topic_of(shape.id) not in keys:
            return False
    return True


# --------------------------------------------------------------
def select_redraw(topics, deck):
    """Topics to redraw, skipped ones, and their slides.

    A topic is redrawn only when this script added it, its
    box is unfilled and in the talk, and its slide holds
    nothing but the script's boxes of topics being redrawn,
    so deleting the slide loses nothing of the author's.
    """
    ledger = {k for k, _ in read_ledger(deck)}
    chosen, skipped = [], []
    for block, spec in topics:
        key = T.block_key(block)
        reason = redraw_blocker(deck, key, ledger)
        if reason:
            skipped.append((block.headline, reason))
        else:
            chosen.append((block, spec))
    keys = {T.block_key(b) for b, _ in chosen}
    homes = {b.headline: deck.shape(
        T.shape_id(T.block_key(b), T.BOX)).slide
        for b, _ in chosen}
    mixed = {sid for sid in homes.values()
             if not slide_is_ours(deck, sid, keys)}
    for block, _ in [c for c in chosen
                     if homes[c[0].headline] in mixed]:
        skipped.append((block.headline,
                        "its slide holds other things"))
    chosen = [c for c in chosen
              if homes[c[0].headline] not in mixed]
    slides = {homes[b.headline] for b, _ in chosen}
    return chosen, skipped, slides


# --------------------------------------------------------------
def redraw_for(topics, locate):
    """Planning function that rebuilds topics' slides."""

    # --------------------------------------
    def plan(deck):
        """Remove the topics' slides, then lay them out anew."""
        chosen, skipped, slides = select_redraw(topics, deck)
        report(chosen, skipped, "redraw")
        if not chosen:
            return []
        at = min(deck.slide(sid).index for sid in slides)
        gone = set(slides)
        for sid in slides:
            gone |= {s.id for s in deck.slide(sid).shapes}
        taken = set(deck.ids()) - gone
        section = Section(title=NEWS_TITLE,
                          blocks=[b for b, _ in chosen])
        draws = []
        for n, page in enumerate(L.paginate([section])):
            keys = [T.block_key(i.block) for i in page.items]
            sid = T.make_unique(T.page_key(page), taken)
            draws += insert_page(sid, at + n, page, keys, locate)
        removes = [r for sid in sorted(slides)
                   for r in S.delete_object(sid)]
        return [("redraw: remove", removes),
                ("redraw: draw", draws)]

    return plan


# --------------------------------------------------------------
def fetch_pictures(fresh, folder):
    """Screenshot or download a picture for each new topic."""
    for block, spec in fresh:
        if not spec:
            continue
        log(f"  picture for {block.headline[:50]}")
        block.image = prepare(spec, L.IMG_W, folder)
        if not block.image:
            log("    not available; the topic goes in without")


# --------------------------------------------------------------
def locator(host):
    """Upload each picture once, however often it is asked."""
    cache = {}

    # --------------------------------------
    def locate(path):
        """The hosted URL for one local picture."""
        if path not in cache:
            cache[path] = host.put(path)
        return cache[path]

    return locate


# --------------------------------------------------------------
def list_known(step):
    """Print every topic the deck already knows about."""
    deck = G.read_deck(step.slides, step.deck_id)
    for line, slide_id in deck_headlines(deck):
        place_name = ("parked" if deck.slide(slide_id).index
                      > deck.separator() else "deck")
        print(f"{place_name:6s} {line}")
    for _, headline in read_ledger(deck):
        print(f"ledger {headline}")


# --------------------------------------------------------------
def run(step, topics, dry_run, redraw=False):
    """Fetch pictures for new topics, then write them.

    With redraw, the topics are ones added before, and their
    slides are drawn again in the current box format.
    """
    deck = G.read_deck(step.slides, step.deck_id)
    if redraw:
        fresh, skipped, _ = select_redraw(topics, deck)
    else:
        fresh, skipped = select(topics, deck)
    if dry_run or not fresh:
        verb = "redraw" if redraw else "new"
        report(fresh, skipped, verb)
        log(f"{LABEL}: {len(fresh)} {verb}, {len(skipped)} "
            f"skipped{' (dry run)' if dry_run else ''}")
        return
    folder = workdir()
    host = Host(step.drive, step.folder_id)
    try:
        fetch_pictures(fresh, folder)
        make = redraw_for if redraw else plan_for
        sent = W.run_plan(step.slides, step.deck_id,
                          make(fresh, locator(host)), LABEL)
    finally:
        host.clean()
        remove_workdir(folder)
    if sent > 0:
        done = "redrawn" if redraw else "added"
        log(f"{LABEL}: {done} ({sent} requests). Run "
            f"g3_update_toc.py to refresh the contents.")


# --------------------------------------------------------------
def main():
    """Add news topics to one deck."""
    flags = {"--list", "--dry-run", "--redraw"}
    argv = [a for a in sys.argv[1:] if a not in flags]
    step = start("Add news topics", argv)
    if "--list" in sys.argv:
        list_known(step)
        return
    data = load_json(step.args.json)
    items = data["topics"] if isinstance(data, dict) else data
    topics = [to_topic(item) for item in items]
    run(step, topics, "--dry-run" in sys.argv,
        "--redraw" in sys.argv)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
