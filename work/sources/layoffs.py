#!/usr/bin/env python3
"""
Read current tech layoff totals from layoffs.fyi and TrueUp.

layoffs.fyi loads its yearly totals from a small JSON API,
so a plain download is enough. TrueUp states its totals in
one sentence at the top of its tracker ("So far in 2026,
there have been 602 layoffs at tech companies with 187,604
people impacted (727 people per day). In 2025, ..."), and
sits behind a Cloudflare check that stops headless Chrome,
so the page is read in a real Chrome window parked
off-screen.

Usage:
    python3 -m sources.layoffs
    from sources import layoffs
    fyi = layoffs.fyi_totals()        # {2026: 128873, ...}
    tru = layoffs.trueup_totals()     # {2026: ("187,604",
                                      #          "727"), ...}

Created: 2026-09-14
Last updated: 2026-09-14
"""

import json
import re
import urllib.request

from sources import element_shot as E
from sources import fetch_images as F

FYI_API = ("https://layoffsfyi-production.up.railway.app"
           "/api/annual-stats")
TRUEUP_URL = "https://trueup.io/layoffs"
TRUEUP_MARKER = "people impacted"
TRUEUP_SENTENCE = re.compile(
    r"in\s+(\d{4}),\s+there\s+(?:have\s+been|were)\s+[\d,]+"
    r"\s+layoffs\s+at\s+tech\s+companies\s+(?:with|w/)\s+"
    r"([\d,]+)\s+people\s+impacted\s+\(([\d,]+)\s+people\s+"
    r"per\s+day\)", re.IGNORECASE)


# --------------------------------------------------------------
def fyi_totals():
    """Year -> tech employees laid off, from layoffs.fyi."""
    request = urllib.request.Request(
        FYI_API, headers={"User-Agent": F.USER_AGENT})
    try:
        with urllib.request.urlopen(request,
                                    timeout=30) as reply:
            data = json.load(reply)
    except (OSError, ValueError) as exc:
        F.log_message(f"  layoffs.fyi totals not read: {exc}")
        return {}
    return {int(y["year"]): int(y["employees"])
            for y in data.get("years", [])
            if "year" in y and "employees" in y}


# --------------------------------------------------------------
def parse_trueup(text):
    """Year -> (people, per day) from TrueUp's page text."""
    return {int(year): (people, per_day)
            for year, people, per_day
            in TRUEUP_SENTENCE.findall(text or "")}


# --------------------------------------------------------------
def trueup_totals():
    """Year -> (people, per day), read from TrueUp's page."""
    text = E.page_text(TRUEUP_URL, TRUEUP_MARKER, visible=True)
    totals = parse_trueup(text)
    if not totals:
        F.log_message("  TrueUp totals not found on the page")
    return totals


# --------------------------------------------------------------
if __name__ == "__main__":
    print("layoffs.fyi", fyi_totals())
    print("TrueUp", trueup_totals())
