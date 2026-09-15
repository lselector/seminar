#!/usr/bin/env python3
"""
Get one picture ready for a slide, on this machine.

Two stages per picture, in a temporary directory:

    fetch   download it ({"src": url}), screenshot a web page
            ({"shot": url}), screenshot only the block under
            a heading ({"shot": url, "element": heading},
            with optional "width" of the viewport,
            "visible": true for pages with a bot check, and
            "ready": a JavaScript test that the chart has
            its data),
            or use a local file ({"file": p})
    clean   trim, flatten and shrink it with ImageMagick to
            the pixels its box on the slide actually needs

Nothing here talks to Google. gslides/host.py uploads the
result so Slides can fetch it.

The download and screenshot code is the same as the Markdown
pipeline uses; it is imported from sources/fetch_images.py and
sources/clean_images.py rather than copied.

Usage:
    from sources.pictures import prepare, workdir
    path = prepare({"shot": url}, 3.6, workdir())

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os
import shutil
import tempfile

from sources import fetch_images as F
from sources import clean_images as C
from sources import element_shot as E

CHALLENGE_MARKS = ("just a moment", "challenge-platform",
                   "cf-chl", "performing security verification")


# --------------------------------------------------------------
def log(message):
    """Print one timestamped line."""
    F.log_message(message)


# --------------------------------------------------------------
def workdir():
    """A fresh temporary directory for one run."""
    return tempfile.mkdtemp(prefix="gslides-")


# --------------------------------------------------------------
def download(url, folder):
    """Download one picture, returning its path or None."""
    data = F.fetch_bytes(url)
    if not data:
        log(f"  could not download {url}")
        return None
    path = os.path.join(folder, "raw" + F.url_extension(url))
    with open(path, "wb") as handle:
        handle.write(data)
    return path


# --------------------------------------------------------------
def challenged(url):
    """Is the page behind a bot wall like Cloudflare's?

    A screenshot of such a page shows "Performing security
    verification" instead of the content, which is worse
    than keeping last week's picture.
    """
    import urllib.error
    import urllib.request
    request = urllib.request.Request(
        url, headers={"User-Agent": F.USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=20) as page:
            text = page.read(200000).decode("utf-8", "ignore")
    except urllib.error.HTTPError as exc:
        text = exc.read(200000).decode("utf-8", "ignore")
    except OSError:
        return False
    lowered = text.lower()
    return any(mark in lowered for mark in CHALLENGE_MARKS)


# --------------------------------------------------------------
def screenshot(url, folder):
    """Screenshot one web page, returning its path or None."""
    if challenged(url):
        log(f"  {url} is behind a bot check; not shooting it")
        return None
    if not os.path.isfile(F.CHROME):
        log(f"  Chrome not found at {F.CHROME}")
        return None
    path = os.path.join(folder, "shot.png")
    command = [
        F.CHROME, "--headless", "--disable-gpu",
        "--hide-scrollbars", f"--window-size={F.SHOT_SIZE}",
        f"--virtual-time-budget={F.SHOT_BUDGET}",
        f"--screenshot={path}", url,
    ]
    return path if F.run_chrome(command, path) else None


# --------------------------------------------------------------
def fetch(spec, folder):
    """Get the raw picture a spec describes."""
    if spec.get("file"):
        path = spec["file"]
        return path if os.path.isfile(path) else None
    if spec.get("src"):
        return download(spec["src"], folder)
    if spec.get("shot") and spec.get("element"):
        visible = spec.get("visible", False)
        if not visible and challenged(spec["shot"]):
            log(f"  {spec['shot']} is behind a bot check")
            return None
        return E.shoot_element(spec["shot"], spec["element"],
                               os.path.join(folder, "shot.png"),
                               width=spec.get("width", E.WIDTH),
                               visible=visible,
                               ready=spec.get("ready"))
    if spec.get("shot"):
        return screenshot(spec["shot"], folder)
    return None


# --------------------------------------------------------------
def clean(raw, width_in, folder):
    """Normalize a raw picture for a box width in inches."""
    if not C.MAGICK:
        log("  ImageMagick not found; using the raw picture")
        return raw
    out = os.path.join(folder, "clean.jpg")
    if C.run_convert(raw, out, C.Settings(), width_in):
        return out
    return None


# --------------------------------------------------------------
def prepare(spec, width_in, folder=None):
    """Fetch and clean one picture, returning a local path."""
    folder = tempfile.mkdtemp(dir=folder or workdir())
    raw = fetch(spec or {}, folder)
    if not raw:
        return None
    return clean(raw, width_in, folder)


# --------------------------------------------------------------
def remove_workdir(folder):
    """Delete a run's temporary directory."""
    shutil.rmtree(folder, ignore_errors=True)
