#!/usr/bin/env python3
"""
Connect to Google and read the workflow settings.

The one place that knows where the OAuth token lives and how
to turn it into Drive and Slides clients. Every g*_ script
starts here, so no script carries its own copy of the
sign-in code.

Settings live in config/gslides.json. They are
not secret: a folder id and a few public URLs.

    {"folder_id": "1XFo...", "youtube_url": "https://..."}

Usage:
    from gslides.client import services, settings, log
    drive, slides = services()
    folder = settings()["folder_id"]

Created: 2026-09-14
Last updated: 2026-09-14
"""

import json
import os
import sys
from datetime import datetime

# work/ is one level above this package.
ROOT = os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(ROOT, "config")
DATA_DIR = os.path.join(ROOT, "data")
TOKEN_FILE = os.path.join(ROOT, "credentials", "token.json")
SETTINGS_FILE = os.path.join(CONFIG_DIR, "gslides.json")


# --------------------------------------------------------------
def log(message):
    """Print one timestamped line."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}", flush=True)


# --------------------------------------------------------------
def credentials():
    """Load the token g_auth.py saved, or explain."""
    if not os.path.isfile(TOKEN_FILE):
        log(f"ERROR: {TOKEN_FILE} not found.")
        log("  Run: python3 g_auth.py <FOLDER_ID>")
        sys.exit(1)
    from google.oauth2.credentials import Credentials
    with open(TOKEN_FILE, encoding="utf-8") as handle:
        scopes = json.load(handle)["scopes"]
    return Credentials.from_authorized_user_file(
        TOKEN_FILE, scopes
    )


# --------------------------------------------------------------
def services():
    """Drive and Slides clients, signed in."""
    from googleapiclient.discovery import build
    creds = credentials()
    drive = build("drive", "v3", credentials=creds,
                  cache_discovery=False)
    slides = build("slides", "v1", credentials=creds,
                   cache_discovery=False)
    return drive, slides


# --------------------------------------------------------------
def settings():
    """The workflow settings, or exit with help."""
    if not os.path.isfile(SETTINGS_FILE):
        log(f"ERROR: {SETTINGS_FILE} not found.")
        sys.exit(1)
    with open(SETTINGS_FILE, encoding="utf-8") as handle:
        return json.load(handle)
