#!/usr/bin/env python3
"""
Collect the pictures a deck Markdown file asks for.

The Google Slides workflow uses fetch_bytes, url_extension
and run_chrome through sources/pictures.py. The command
line below dates from the cancelled Markdown workflow and
reads a deck .md file.

Reads the deck file, finds every image reference and its
manifest comment, and puts a source picture into
"images_raw". The manifest lives next to the picture it
describes, so there is no second list to fall out of date:

    ![](images/gpt6-astra.jpg)
    <!-- src: https://openai.com/astra-hero.png -->
    downloaded with urllib

    ![](images/intelligence-index.jpg)
    <!-- shot: https://artificialanalysis.ai/... -->
    screenshotted with headless Chrome

    ![](images/lev-photo.jpg)
    (no comment) an original, checked but never fetched

Screenshots are how the pages that cannot be downloaded get
onto slides: the Intelligence Index chart and the layoff
trackers are live web pages, not image files.

Files already in "images_raw" are skipped, so the script is
safe to re-run. Use --force to fetch everything again.

Run sources/clean_images.py afterwards.

Usage:
    python3 -m sources.fetch_images 2026-09-18-AI-News.md
    python3 -m sources.fetch_images deck.md --force

Created: 2026-09-10
Last updated: 2026-09-12
"""

import glob
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

from layout.deck_parser import image_manifest, parse_deck_file

# work/data, whatever directory the command runs from.
DATA_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data")
RAW_DIR = os.path.join(DATA_DIR, "images_raw")
TIMEOUT = 30
RETRIES = 2
BACKOFF = 2
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"

CHROME = (
    "/Applications/Google Chrome.app/Contents/MacOS/"
    "Google Chrome"
)
SHOT_SIZE = "1600,1000"
SHOT_BUDGET = "9000"
SHOT_TIMEOUT = 90

IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.webp', '.gif')


# --------------------------------------------------------------
def log_message(message):
    """Print timestamped log message."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}")


# --------------------------------------------------------------
def raw_stem(deck_path):
    """Take the file stem a deck image path refers to."""
    name = os.path.basename(deck_path)
    return os.path.splitext(name)[0]


# --------------------------------------------------------------
def existing_raw(stem):
    """Return an existing raw file for a stem, or None."""
    matches = glob.glob(os.path.join(RAW_DIR, f"{stem}.*"))
    matches = [p for p in matches if not p.endswith('.tmp')]
    return matches[0] if matches else None


# --------------------------------------------------------------
def url_extension(url):
    """Pick a file extension from a URL, defaulting to jpg."""
    clean = url.split('?')[0].split('#')[0]
    ext = os.path.splitext(clean)[1].lower()
    return ext if ext in IMAGE_EXTS else '.jpg'


# --------------------------------------------------------------
def fetch_bytes(url):
    """Download a URL with retries, returning bytes."""
    request = urllib.request.Request(
        url, headers={'User-Agent': USER_AGENT}
    )
    for attempt in range(1, RETRIES + 1):
        try:
            with urllib.request.urlopen(
                request, timeout=TIMEOUT
            ) as response:
                return response.read()
        except (urllib.error.URLError, OSError) as exc:
            log_message(f"  attempt {attempt} failed: {exc}")
            if attempt < RETRIES:
                time.sleep(BACKOFF)
    return None


# --------------------------------------------------------------
def write_atomic(path, data):
    """Write bytes to a path through a temporary file."""
    temp = f"{path}.tmp"
    with open(temp, 'wb') as handle:
        handle.write(data)
    os.replace(temp, path)


# --------------------------------------------------------------
def download_image(stem, url):
    """Download one picture into the raw directory."""
    path = os.path.join(RAW_DIR, f"{stem}{url_extension(url)}")
    log_message(f"Downloading {stem} from {url}")
    data = fetch_bytes(url)
    if not data:
        log_message(f"  giving up on {stem}")
        return False
    write_atomic(path, data)
    log_message(f"  saved {path} ({len(data) // 1024} KB)")
    return True


# --------------------------------------------------------------
def shoot_page(stem, url):
    """Screenshot one web page with headless Chrome."""
    if not os.path.isfile(CHROME):
        log_message(f"  Chrome not found at {CHROME}")
        return False
    path = os.path.join(RAW_DIR, f"{stem}.png")
    # Chrome picks the image format from the extension and
    # refuses anything it does not know, so the temporary
    # name has to end in .png too.
    temp = os.path.join(RAW_DIR, f".tmp-{stem}.png")
    log_message(f"Shooting {stem} from {url}")
    command = [
        CHROME, '--headless', '--disable-gpu',
        '--hide-scrollbars', f'--window-size={SHOT_SIZE}',
        f'--virtual-time-budget={SHOT_BUDGET}',
        f'--screenshot={temp}', url,
    ]
    if not run_chrome(command, temp):
        return False
    os.replace(temp, path)
    size_kb = os.path.getsize(path) // 1024
    log_message(f"  saved {path} ({size_kb} KB)")
    return True


# --------------------------------------------------------------
def run_chrome(command, temp):
    """Run headless Chrome, returning True on success."""
    try:
        subprocess.run(
            command, capture_output=True,
            timeout=SHOT_TIMEOUT, check=False
        )
    except (subprocess.TimeoutExpired, OSError) as exc:
        log_message(f"  screenshot failed: {exc}")
        return False
    if not os.path.isfile(temp):
        log_message("  screenshot produced no file")
        return False
    return True


# --------------------------------------------------------------
def check_local(stem, deck_path):
    """Warn when a hand-placed original is missing."""
    if existing_raw(stem):
        return True
    if os.path.isfile(deck_path):
        return True
    log_message(
        f"WARNING: {deck_path} has no source comment and no "
        f"file. Put the original in {RAW_DIR}."
    )
    return False


# --------------------------------------------------------------
def fetch_one(deck_path, url, mode, force):
    """Fetch one manifest entry, returning its outcome."""
    stem = raw_stem(deck_path)

    if mode == "local":
        if check_local(stem, deck_path):
            return "local"
        return "error"

    if existing_raw(stem) and not force:
        log_message(f"Skipping {stem} - already fetched")
        return "skipped"

    if mode == "src":
        ok = download_image(stem, url)
    else:
        ok = shoot_page(stem, url)
    return "fetched" if ok else "error"


# --------------------------------------------------------------
def fetch_all(manifest, force):
    """Fetch every manifest entry, counting outcomes."""
    counts = {
        "fetched": 0, "skipped": 0,
        "local": 0, "error": 0,
    }
    for deck_path, url, mode in manifest:
        outcome = fetch_one(deck_path, url, mode, force)
        counts[outcome] += 1
    return counts


# --------------------------------------------------------------
def resolve_deck(argv):
    """Take the deck file path from the command line."""
    named = [a for a in argv[1:] if not a.startswith('-')]
    if not named:
        log_message(
            "Usage: python3 -m sources.fetch_images deck.md"
        )
        sys.exit(1)
    if not os.path.isfile(named[0]):
        log_message(f"ERROR: {named[0]} not found")
        sys.exit(1)
    return named[0]


# --------------------------------------------------------------
def print_summary(counts, total):
    """Print the collection summary."""
    log_message("=" * 50)
    log_message(f"Pictures declared in the deck: {total}")
    log_message(f"Fetched now: {counts['fetched']}")
    log_message(f"Already present: {counts['skipped']}")
    log_message(f"Local originals: {counts['local']}")
    log_message(f"Failed: {counts['error']}")
    log_message(f"Files in {RAW_DIR}: "
                f"{len(os.listdir(RAW_DIR))}")
    log_message("Next: python3 -m sources.clean_images")


# --------------------------------------------------------------
def main():
    """Collect every picture one deck file asks for."""
    source = resolve_deck(sys.argv)
    force = '--force' in sys.argv

    log_message(f"Reading {source}")
    os.makedirs(RAW_DIR, exist_ok=True)

    manifest = image_manifest(parse_deck_file(source))
    log_message(f"Found {len(manifest)} picture references")

    counts = fetch_all(manifest, force)
    print_summary(counts, len(manifest))


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
