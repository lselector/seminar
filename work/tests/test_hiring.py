#!/usr/bin/env python3
"""
Tests for step 10, the Tech Hiring slide, and its sources.

Offline: CSV text and decks are built in the tests.

Usage:
    python3 -m tests.test_hiring

Created: 2026-09-27
Last updated: 2026-09-27
"""

import datetime as dt

import g1_new_deck as G1
import g10_update_hiring as HIRE
from gslides import deck as G
from gslides import ids as T
from layout import deck_layout as L
from sources import hiring as SRC
from tests.runner import run

FRED_TEXT = ("observation_date,IHLIDXUSTPSOFTDEVE,IHLIDXUS\n"
             "2025-09-18,64.87,102.73\n"
             "2026-02-01,80.00,.\n"
             "2026-09-18,77.32,103.45\n")
AI_TEXT = ("date,jobcountry,AI_share_postings\n"
           "2025-08-31,US,3.44\n2026-08-31,GB,5.0\n"
           "2026-08-31,US,6.74\n")


# --------------------------------------------------------------
def facts():
    """Facts parsed from the sample CSV text."""
    columns = SRC.fred_columns(FRED_TEXT)
    return SRC.Facts(software=columns[SRC.SOFTWARE],
                     all_jobs=columns[SRC.ALL_JOBS],
                     ai_share=SRC.ai_series(AI_TEXT))


# --------------------------------------------------------------
def test_fred_csv_skips_missing_values_per_column():
    """A "." drops that day from one series only."""
    columns = SRC.fred_columns(FRED_TEXT)
    assert len(columns[SRC.SOFTWARE].points) == 3
    assert len(columns[SRC.ALL_JOBS].points) == 2
    software = columns[SRC.SOFTWARE]
    assert software.last == (dt.date(2026, 9, 18), 77.32)
    assert software.year_ago == (dt.date(2025, 9, 18), 64.87)
    assert software.peak == (dt.date(2026, 2, 1), 80.0)


# --------------------------------------------------------------
def test_ai_share_keeps_only_us_rows():
    """Other countries in the file are ignored."""
    series = SRC.ai_series(AI_TEXT)
    assert [v for _, v in series.points] == [3.44, 6.74]


# --------------------------------------------------------------
def test_jobs_topic_states_the_numbers_and_ends_with_source():
    """Index, yearly change, peak, all jobs, AI, link."""
    block = HIRE.jobs_block(facts())
    text = "\n".join(block.bullets)
    assert "**77**" in text and "Sept 18" in text
    assert "==+19%== from a year ago (65)" in text
    assert "Peak was 80 in Feb 2026" in text
    assert "All US job postings: 103, +1%" in text
    assert "**6.7%** of all, up from 3.4%" in text
    assert block.bullets[-1] == SRC.FRED_PAGE


# --------------------------------------------------------------
def test_no_postings_numbers_means_no_jobs_topic():
    """A failed download leaves the box as it is."""
    assert HIRE.jobs_block(SRC.Facts()) is None


# --------------------------------------------------------------
def test_skills_json_gives_block_and_picture():
    """The source URL is the last line; no image by default."""
    block, spec = HIRE.skills_part({"skills": {
        "bullets": ["AI agents **+200%**"],
        "url": "https://x.io/report"}})
    assert block.headline == HIRE.SKILLS_HEADLINE
    assert block.bullets == ["AI agents **+200%**",
                             "https://x.io/report"]
    assert spec is None
    assert HIRE.skills_part({}) == (None, None)


# --------------------------------------------------------------
def test_a_region_shot_takes_the_frame_shape():
    """Aspect comes from the picture frame, not the JSON."""
    frame = G.Shape(id="p", slide="s", kind=G.KIND_IMAGE,
                    rect=L.Rect(0, 0, 3.0, 2.0))
    spec = HIRE.skills_picture({"shot": "u", "region": "body"},
                               frame)
    assert spec["aspect"] == 1.5
    plain = {"shot": "u"}
    assert HIRE.skills_picture(plain, frame) == plain


# --------------------------------------------------------------
def test_the_dice_table_is_found_by_its_title():
    """The embed id changes monthly; the title does not."""
    html = ('<iframe title="Top 50 Tech Job Titles" '
            'src="https://e.infogram.com/aaa-1?src=embed">'
            '<iframe style="x" title="Top 50 Skills" '
            'src="https://e.infogram.com/bbb-2?src=embed">')
    assert SRC.infogram_url(html, "Top 50 Skills") == \
        "https://e.infogram.com/bbb-2?src=embed"
    assert SRC.infogram_url("<p></p>", "Top 50 Skills") is None


# --------------------------------------------------------------
def test_a_carried_layoffs_page_lands_before_a_drawn_hiring():
    """Last week had layoffs but no hiring page yet."""
    kept = {"layoffs": "g_lay"}
    pages = G1.build_pages("AI News - Oct 9, 2026",
                           lambda path: "https://x/" + path,
                           skip=kept)
    order = G1.final_order([sid for sid, _ in pages], kept)
    assert order.index("g_lay") + 1 == \
        order.index(T.PAGE_HIRING)
    assert order.index(T.PAGE_HIRING) + 1 == \
        order.index(T.PAGE_ABOUT)


# --------------------------------------------------------------
def test_both_pages_carried_keep_their_order():
    """Layoffs, then hiring, then About."""
    kept = {"layoffs": "g_lay", "hiring": "g_hire"}
    pages = G1.build_pages("AI News - Oct 9, 2026",
                           lambda path: "https://x/" + path,
                           skip=kept)
    ids = [sid for sid, _ in pages]
    assert T.PAGE_HIRING not in ids
    order = G1.final_order(ids, kept)
    at = order.index("g_lay")
    assert order[at:at + 3] == ["g_lay", "g_hire",
                                T.PAGE_ABOUT]


# --------------------------------------------------------------
if __name__ == "__main__":
    run(globals())
