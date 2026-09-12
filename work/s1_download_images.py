#!/usr/bin/env python3
"""
Collect the images used by GUIDE_to_M92.md.

Downloads one thumbnail for each YouTube video the guide
links to, into the local "images" directory.

The engine data plate photographs are not downloaded. They
are originals that live in "images" and are kept in the
repository, so this script only checks that they are still
there and warns if one has gone missing.

Files that are already present are skipped, so the script
is safe to re-run. Use --force to fetch the thumbnails
again.

Run s2_clean_images.py afterwards to bring the collected
files to the standard format.

Usage:
    python3 s1_download_images.py
    python3 s1_download_images.py --force

Created: 2026-09-10
Last updated: 2026-09-10
"""

import glob
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime

IMAGES_DIR = "images"
TIMEOUT = 30
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"

# Engine data plate photos. These are originals, not
# downloads: they are the evidence for the serial number
# decoded in section 1 of the guide. Checked, never fetched.
ENGINE_PHOTOS = [
    "engine-plate-closeup",
    "engine-plate-wide",
    "engine-plate-angle",
]

# Video thumbnails: (local name, YouTube video id).
# No M92 specific footage exists, so these cover the base
# 1000 Series engine and its injection pump.
THUMBNAILS = [
    ("video-perkins-timing-marks", "8O38u6g4p5M"),
    ("video-install-and-time-pump", "rbKTiMVtINc"),
    ("video-4-cylinder-pump-timing", "M_dtUUS81A0"),
    ("video-1004-4t-pump-renewal", "So6aybjCPsE"),
]

# Best first: YouTube serves 404 for sizes it does not have
THUMB_SIZES = ["maxresdefault", "sddefault", "hqdefault"]


# --------------------------------------------------------------
def log_message(message):
    """Print timestamped log message."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


# --------------------------------------------------------------
def existing_file(name):
    """Return existing image path for name, or None."""
    matches = glob.glob(os.path.join(IMAGES_DIR, f"{name}.*"))
    return matches[0] if matches else None


# --------------------------------------------------------------
def target_path(name, source):
    """Build local path for name using the source suffix."""
    ext = os.path.splitext(source)[1].lower()
    if ext not in ('.jpg', '.jpeg', '.png', '.webp'):
        ext = '.jpg'
    return os.path.join(IMAGES_DIR, f"{name}{ext}")


# --------------------------------------------------------------
def fetch_bytes(url):
    """Download URL and return its bytes, or None."""
    request = urllib.request.Request(
        url,
        headers={'User-Agent': USER_AGENT}
    )
    try:
        with urllib.request.urlopen(request,
                                    timeout=TIMEOUT) as resp:
            return resp.read()
    except (urllib.error.URLError, OSError) as exc:
        log_message(f"  failed: {exc}")
        return None


# --------------------------------------------------------------
def save_image(name, url, force):
    """Download one image unless it is already present."""
    present = existing_file(name)
    if present and not force:
        log_message(f"Skipping {name} - already have "
                    f"{os.path.basename(present)}")
        return False

    log_message(f"Downloading {name}")
    data = fetch_bytes(url)
    if not data:
        return False

    if present and force:
        os.remove(present)

    path = target_path(name, url)
    with open(path, 'wb') as handle:
        handle.write(data)

    size_kb = len(data) // 1024
    log_message(f"  saved {path} ({size_kb} KB)")
    return True


# --------------------------------------------------------------
def thumbnail_url(video_id):
    """Return the best available thumbnail URL, or None."""
    base = "https://img.youtube.com/vi"
    for size in THUMB_SIZES:
        url = f"{base}/{video_id}/{size}.jpg"
        request = urllib.request.Request(
            url,
            method='HEAD',
            headers={'User-Agent': USER_AGENT}
        )
        try:
            with urllib.request.urlopen(request,
                                        timeout=TIMEOUT):
                return url
        except (urllib.error.URLError, OSError):
            continue
    log_message(f"  no thumbnail found for {video_id}")
    return None


# --------------------------------------------------------------
def check_photos():
    """Warn about any missing engine plate photo."""
    missing = [n for n in ENGINE_PHOTOS if not existing_file(n)]
    for name in missing:
        log_message(f"WARNING: {name} is missing from "
                    f"{IMAGES_DIR}")
    if missing:
        log_message("These are originals and cannot be "
                    "re-downloaded. Restore them from a backup.")
    return len(ENGINE_PHOTOS) - len(missing)


# --------------------------------------------------------------
def download_thumbnails(force):
    """Download all video thumbnails."""
    count = 0
    for name, video_id in THUMBNAILS:
        if existing_file(name) and not force:
            log_message(f"Skipping {name} - already present")
            continue
        url = thumbnail_url(video_id)
        if url and save_image(name, url, force):
            count += 1
    return count


# --------------------------------------------------------------
def main():
    """Collect every image needed by GUIDE_to_M92.md."""
    force = '--force' in sys.argv

    log_message("Starting image collection")
    os.makedirs(IMAGES_DIR, exist_ok=True)

    photos = check_photos()
    thumbs = download_thumbnails(force)

    log_message("=" * 50)
    log_message(f"Engine photos present: {photos} of "
                f"{len(ENGINE_PHOTOS)}")
    log_message(f"Video thumbnails downloaded: {thumbs}")
    log_message(f"Files now in {IMAGES_DIR}: "
                f"{len(os.listdir(IMAGES_DIR))}")
    log_message("Next: python3 s2_clean_images.py")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
