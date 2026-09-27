#!/usr/bin/env python3
"""
Step 10: refresh the Tech Hiring slide.

The slide sits right after Jobs and Layoffs and holds two
script topics:

    t-hiring-jobs    Software job postings on Indeed. The
                     script reads the numbers itself (Indeed's
                     postings indexes via FRED, and Indeed's
                     share of AI postings) and writes the
                     bullets; the picture is FRED's chart of
                     software postings against all postings.
    t-hiring-skills  Skills in demand. The words need
                     judgement, so a skill researches them and
                     passes them as JSON. The picture is Dice's
                     current "Top 50 Skills" table, shot in the
                     frame's shape, unless the JSON names
                     another ("image", a picture spec):

    {"skills": {"headline": "Skills in demand",
                "bullets": ["AI and ML postings **+101%**",
                            "..."],
                "url": "https://www.dice.com/..."}}

Without --json only the job postings topic is refreshed;
the skills words and picture change together.
A filled box is frozen, with its picture. If the slide is
missing (a deck made before this step existed, or deleted
by hand), it is drawn again from config/skeleton.json first.

Usage:
    python3 g10_update_hiring.py
    python3 g10_update_hiring.py --json hiring.json
    python3 g10_update_hiring.py 2026-10-02 --json hiring.json

Created: 2026-09-27
Last updated: 2026-09-27
"""

from layout import skeleton as K
from layout.deck_parser import Block
from gslides import api as S
from gslides import deck as G
from gslides import render as R
from gslides import write as W
from gslides import ids as T
from gslides.client import ROOT, log
from gslides.host import Host
from gslides.step import (
    block_from, load_json, report_frozen, start,
    with_pictures,
)
from sources import hiring as SRC
from sources.pictures import remove_workdir, workdir

# How the Dice table is shot: the whole embed page from the
# top, at a viewport as narrow as the table reads well.
DICE_REGION = "body"
DICE_WIDTH = 700

LABEL = "hiring"
PAGE_KEY = "hiring"
KEYS = [T.TOPIC_HIRING_JOBS, T.TOPIC_HIRING_SKILLS]
JOBS_HEADLINE = "Software job postings on Indeed"
SKILLS_HEADLINE = "Skills in demand"


# --------------------------------------------------------------
def title():
    """The slide's title, as config/skeleton.json has it."""
    return K.page(PAGE_KEY)["title"]


# --------------------------------------------------------------
def find_page(deck):
    """The hiring slide in the talk, or None."""
    return G.titled_page(deck, T.PAGE_HIRING, title())


# --------------------------------------------------------------
def insert_index(deck):
    """Where a missing hiring slide goes: after layoffs.

    Falls back to just before About, then to the separator.
    """
    layoffs = G.titled_page(deck, T.PAGE_LAYOFFS,
                            K.page("layoffs")["title"])
    if layoffs:
        return layoffs.index + 1
    about = deck.slide(T.PAGE_ABOUT)
    if about and about.index < deck.separator():
        return about.index
    return deck.separator()


# --------------------------------------------------------------
def draw_page(deck, locate):
    """Requests that add the skeleton hiring slide."""
    head = S.new_slide(T.PAGE_HIRING)
    head[0]["createSlide"]["insertionIndex"] = \
        insert_index(deck)
    section = K.section(K.page(PAGE_KEY), ROOT)
    return head + R.render_fixed(T.PAGE_HIRING, KEYS, section,
                                 locate)


# --------------------------------------------------------------
def ensure_page(step):
    """Draw the slide if the deck has none; True if it has."""
    if find_page(G.read_deck(step.slides, step.deck_id)):
        return True
    log(f"{LABEL}: no '{title()}' slide; drawing it")
    host = Host(step.drive, step.folder_id)
    try:
        sent = W.run_plan(
            step.slides, step.deck_id,
            lambda deck: [(LABEL, draw_page(deck, host.put))],
            LABEL)
    finally:
        host.clean()
    return sent > 0


# --------------------------------------------------------------
def change(now, then):
    """Percent change from then to now, signed: "+19%"."""
    return f"{(now / then - 1) * 100:+.0f}%"


# --------------------------------------------------------------
def month_year(day):
    """A date as the deck writes a month: "Feb 2022"."""
    return f"{G.MONTHS[day.month - 1]} {day.year}"


# --------------------------------------------------------------
def software_lines(series):
    """Bullets about software development postings."""
    day, now = series.last
    _, then = series.year_ago
    peak_day, peak = series.peak
    return [
        f"Software development postings: **{now:.0f}** "
        f"(Feb 2020 = 100), {W.date_label(day)}",
        f"=={change(now, then)}== from a year ago ({then:.0f})",
        f"Peak was {peak:.0f} in {month_year(peak_day)}; "
        f"now {change(now, peak)} from the peak",
    ]


# --------------------------------------------------------------
def other_lines(facts):
    """Bullets about all postings and the AI share."""
    lines = []
    if facts.all_jobs:
        _, now = facts.all_jobs.last
        _, then = facts.all_jobs.year_ago
        lines.append(f"All US job postings: {now:.0f}, "
                     f"{change(now, then)} from a year ago")
    if facts.ai_share:
        _, now = facts.ai_share.last
        _, then = facts.ai_share.year_ago
        way = "up" if now >= then else "down"
        lines.append(f"Postings that mention AI: "
                     f"**{now:.1f}%** of all, {way} from "
                     f"{then:.1f}% a year ago")
    return lines


# --------------------------------------------------------------
def jobs_block(facts):
    """The job postings topic, or None without the numbers."""
    if not facts.software:
        return None
    bullets = software_lines(facts.software) + \
        other_lines(facts) + [SRC.FRED_PAGE]
    return Block(headline=JOBS_HEADLINE, bullets=bullets)


# --------------------------------------------------------------
def skills_part(data):
    """(Block, picture spec) from the JSON, or (None, None)."""
    part = data.get("skills")
    if not part:
        return None, None
    url = part.get("url")
    return block_from(part, SKILLS_HEADLINE, url), \
        part.get("image")


# --------------------------------------------------------------
def skills_picture(spec, frame):
    """The skills picture spec: the JSON's, else Dice's table.

    A region shot is taken in the frame's own shape.
    """
    if not spec:
        chart = SRC.dice_skills_chart()
        if not chart:
            log("  skills: Dice table not found; old picture "
                "stays")
            return None
        spec = {"shot": chart, "region": DICE_REGION,
                "width": DICE_WIDTH}
    if spec.get("region") and frame and "aspect" not in spec:
        spec = dict(spec, aspect=frame.rect.w / frame.rect.h)
    return spec


# --------------------------------------------------------------
def plan_for(blocks):
    """Build the planning function, given hosted pictures."""

    # --------------------------------------
    def make(urls):
        """Close over the hosted picture URLs."""

        # ----------------------------------
        def plan(deck):
            """Refresh both topics independently."""
            if not find_page(deck):
                log(f"{LABEL}: the slide is gone or parked. "
                    f"Nothing done.")
                return []
            batches = []
            for key in KEYS:
                report_frozen(deck, [T.shape_id(key, T.BOX)],
                              LABEL)
                batches.append((key, topic_requests(
                    deck, key, blocks.get(key), urls)))
            return batches

        return plan

    return make


# --------------------------------------------------------------
def topic_requests(deck, key, block, urls):
    """Words and picture for one topic, as far as allowed."""
    picture = T.shape_id(key, T.PICTURE)
    url = urls.get(picture)
    if block:
        return W.refresh_topic(deck, key, block, url)
    return W.refresh_picture(deck, picture, url)


# --------------------------------------------------------------
def gather(step, folder):
    """Topic blocks and picture specs for this run."""
    data = load_json(step.args.json) if step.args.json else {}
    facts = SRC.hiring_facts()
    jobs = jobs_block(facts)
    log(f"  job postings: {'read' if jobs else 'not read'}")
    skills, skills_spec = skills_part(data)
    if skills:
        deck = G.read_deck(step.slides, step.deck_id)
        skills_spec = skills_picture(skills_spec, deck.shape(
            T.shape_id(T.TOPIC_HIRING_SKILLS, T.PICTURE)))
    else:
        log("  skills: no --json, left as they are")
    blocks = {T.TOPIC_HIRING_JOBS: jobs,
              T.TOPIC_HIRING_SKILLS: skills}
    specs = {}
    chart = SRC.save_chart(folder)
    if chart:
        specs[T.shape_id(T.TOPIC_HIRING_JOBS, T.PICTURE)] = \
            {"file": chart}
    if skills_spec:
        specs[T.shape_id(T.TOPIC_HIRING_SKILLS, T.PICTURE)] = \
            skills_spec
    return blocks, specs


# --------------------------------------------------------------
def main():
    """Refresh the hiring slide of one deck."""
    step = start("Update the hiring slide")
    if not ensure_page(step):
        log(f"{LABEL}: could not draw the slide. Nothing done.")
        return
    folder = workdir()
    try:
        blocks, specs = gather(step, folder)
        sent = with_pictures(step, specs, plan_for(blocks),
                             LABEL)
    finally:
        remove_workdir(folder)
    if sent > 0:
        log(f"{LABEL}: updated ({sent} requests)")
    elif sent == 0:
        log(f"{LABEL}: nothing to change")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
