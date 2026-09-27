#!/usr/bin/env python3
"""
Read current US hiring numbers from Indeed Hiring Lab data.

Three public, keyless sources:

    FRED        Indeed's daily job postings indexes
                (Feb 1, 2020 = 100): software development
                (IHLIDXUSTPSOFTDEVE) and all postings
                (IHLIDXUS). FRED also draws both as one
                chart, which step 10 puts on the slide.
    ai-tracker  Indeed's share of US postings that mention
                AI, on GitHub (hiring-lab/ai-tracker),
                refreshed monthly.
    Dice        the monthly Tech Jobs Report. Its "Top 50
                Skills" table is an Infogram embed; step 10
                screenshots that embed on its own.

Usage:
    python3 -m sources.hiring
    from sources import hiring
    facts = hiring.hiring_facts()
    facts.software.last          # (date, 77.32)
    facts.software.year_ago      # (date, 68.9)
    path = hiring.save_chart("/tmp")
    url = hiring.dice_skills_chart()  # the embed, or None

Created: 2026-09-27
Last updated: 2026-09-27
"""

import datetime as dt
import os
import re
import urllib.request
from dataclasses import dataclass

from sources import fetch_images as F

SOFTWARE = "IHLIDXUSTPSOFTDEVE"
ALL_JOBS = "IHLIDXUS"
FRED_CSV = ("https://fred.stlouisfed.org/graph/fredgraph.csv"
            f"?id={SOFTWARE},{ALL_JOBS}")
FRED_CHART = ("https://fred.stlouisfed.org/graph/fredgraph.png"
              f"?id={SOFTWARE},{ALL_JOBS}&width=1200"
              "&height=700")
FRED_PAGE = f"https://fred.stlouisfed.org/series/{SOFTWARE}"
AI_CSV = ("https://raw.githubusercontent.com/hiring-lab/"
          "ai-tracker/master/AI_posting.csv")
COUNTRY = "US"
DICE_REPORT = ("https://www.dice.com/hiring/recruitment/"
               "reports/dice-tech-job-report")
DICE_SKILLS = "Top 50 Skills"
INFOGRAM = "https://e.infogram.com/{}?src=embed"
# FRED's bot filter stalls a browser User-Agent and some
# custom ones until the read times out, but lets Python's
# default one through (probed 2026-09-27), so no User-Agent
# is set here.
TIMEOUT = 60


@dataclass
class Series:
    """One daily series: its points, oldest first."""

    points: list

    # --------------------------------------
    @property
    def last(self):
        """The newest (date, value)."""
        return self.points[-1]

    # --------------------------------------
    @property
    def year_ago(self):
        """The (date, value) closest to a year before last."""
        target = self.last[0] - dt.timedelta(days=365)
        return min(self.points,
                   key=lambda p: abs((p[0] - target).days))

    # --------------------------------------
    @property
    def peak(self):
        """The highest (date, value)."""
        return max(self.points, key=lambda p: p[1])


@dataclass
class Facts:
    """Everything the hiring slide states as numbers."""

    software: Series = None
    all_jobs: Series = None
    ai_share: Series = None


# --------------------------------------------------------------
def parse_date(text):
    """A YYYY-MM-DD string as a date."""
    return dt.date.fromisoformat(text.strip())


# --------------------------------------------------------------
def fred_columns(text):
    """Column name -> Series, from a FRED CSV download.

    FRED writes "." for a missing value; those rows are
    skipped for that column only.
    """
    lines = [l for l in text.strip().splitlines() if l]
    names = lines[0].split(",")[1:]
    columns = {name: [] for name in names}
    for line in lines[1:]:
        cells = line.split(",")
        day = parse_date(cells[0])
        for name, cell in zip(names, cells[1:]):
            if cell.strip() not in ("", "."):
                columns[name].append((day, float(cell)))
    return {n: Series(p) for n, p in columns.items() if p}


# --------------------------------------------------------------
def ai_series(text, country=COUNTRY):
    """The AI share of postings for one country."""
    points = []
    for line in text.strip().splitlines()[1:]:
        cells = line.split(",")
        if len(cells) == 3 and cells[1] == country:
            points.append((parse_date(cells[0]),
                           float(cells[2])))
    return Series(sorted(points)) if points else None


# --------------------------------------------------------------
def fetch(url):
    """A URL's bytes, or None when it cannot be fetched."""
    try:
        with urllib.request.urlopen(url,
                                    timeout=TIMEOUT) as reply:
            return reply.read()
    except OSError as exc:
        F.log_message(f"  {url} not read: {exc}")
        return None


# --------------------------------------------------------------
def download(url):
    """A CSV as text, or "" when it cannot be fetched."""
    data = fetch(url)
    return data.decode("utf-8", "ignore") if data else ""


# --------------------------------------------------------------
def save_chart(folder):
    """FRED's chart of both indexes as a PNG; path or None."""
    data = fetch(FRED_CHART)
    if not data:
        return None
    path = os.path.join(folder, "fred-postings.png")
    with open(path, "wb") as handle:
        handle.write(data)
    return path


# --------------------------------------------------------------
def infogram_url(html, title):
    """The embed URL of the Infogram iframe with this title.

    Each monthly report embeds new charts, so the id is
    looked up by the iframe's title, not kept.
    """
    pattern = (r'title="' + re.escape(title) + r'"[^>]*?'
               r'src="https://e\.infogram\.com/([0-9a-f-]+)')
    found = re.search(pattern, html)
    return INFOGRAM.format(found.group(1)) if found else None


# --------------------------------------------------------------
def dice_skills_chart():
    """Dice's current "Top 50 Skills" embed URL, or None."""
    return infogram_url(download(DICE_REPORT), DICE_SKILLS)


# --------------------------------------------------------------
def hiring_facts():
    """Read both sources; a part that fails stays None."""
    facts = Facts()
    fred = download(FRED_CSV)
    if fred:
        columns = fred_columns(fred)
        facts.software = columns.get(SOFTWARE)
        facts.all_jobs = columns.get(ALL_JOBS)
    ai = download(AI_CSV)
    if ai:
        facts.ai_share = ai_series(ai)
    return facts


# --------------------------------------------------------------
if __name__ == "__main__":
    found = hiring_facts()
    for label in ("software", "all_jobs", "ai_share"):
        series = getattr(found, label)
        print(label, series and (series.last, series.year_ago,
                                 series.peak))
