#!/usr/bin/env python3
"""
Sign in to Google once, then prove what the token can do.

Two jobs. It runs the OAuth flow and stores a refreshable
token, and it answers the question that decides whether the
narrow "drive.file" scope is enough for this project:

    can the script create a deck inside an existing folder
    of yours, given only that folder's id?

Under drive.file an app may only touch files it created
itself. Whether it may put a new file into a folder you
made by hand is not stated plainly in Google's docs, so
this measures it rather than assuming.

The test creates a throwaway presentation, edits it through
the Slides API, then deletes it, and prints a verdict.

Put the OAuth client here first:

    work/credentials/client_secret.json

That path is gitignored. Get the folder id from the address
bar with the folder open in Drive:

    .../drive/folders/THIS_PART_IS_THE_ID

Usage:
    python3 g_auth.py                 sign in only
    python3 g_auth.py <FOLDER_ID>     sign in and test

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os
import sys

CRED_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "credentials")
CLIENT_FILE = os.path.join(CRED_DIR, "client_secret.json")
TOKEN_FILE = os.path.join(CRED_DIR, "token.json")

SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/documents",
]

DECK_MIME = "application/vnd.google-apps.presentation"
TEST_NAME = "DELETE ME - seminar access test"


# --------------------------------------------------------------
def say(message):
    """Print one line of the report."""
    print(message)


# --------------------------------------------------------------
def load_libraries():
    """Import the Google client libraries, or explain."""
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        from googleapiclient.errors import HttpError
        return Request, Credentials, InstalledAppFlow, build, \
            HttpError
    except ImportError as exc:
        say(f"ERROR: missing library: {exc}")
        say("  pip install google-api-python-client "
            "google-auth-oauthlib")
        sys.exit(1)


# --------------------------------------------------------------
def check_client_file():
    """Stop early if the OAuth client is not in place."""
    if os.path.isfile(CLIENT_FILE):
        return
    say(f"ERROR: {CLIENT_FILE} not found.")
    say("")
    say("Create it in the Google Cloud console:")
    say("  project jarvis-lev -> APIs & Services")
    say("  -> Credentials -> Create credentials")
    say("  -> OAuth client ID -> Desktop app")
    say("  -> Download JSON, save it at the path above")
    sys.exit(1)


# --------------------------------------------------------------
def stored_token(Credentials, Request):
    """Load a saved token and refresh it if it is stale."""
    if not os.path.isfile(TOKEN_FILE):
        return None
    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE, SCOPES
    )
    if creds and creds.expired and creds.refresh_token:
        say("Refreshing the stored token")
        creds.refresh(Request())
    return creds if creds and creds.valid else None


# --------------------------------------------------------------
def authorize(Request, Credentials, InstalledAppFlow):
    """Return valid credentials, signing in if needed."""
    creds = stored_token(Credentials, Request)
    if creds:
        say(f"Using the token in {TOKEN_FILE}")
        return creds

    check_client_file()
    say("Opening a browser to sign in.")
    say("  Expect an 'unverified app' warning: click")
    say("  Advanced, then 'Go to ...'. That is normal for")
    say("  an app only you use.")
    flow = InstalledAppFlow.from_client_secrets_file(
        CLIENT_FILE, SCOPES
    )
    creds = flow.run_local_server(port=0)

    os.makedirs(CRED_DIR, exist_ok=True)
    with open(TOKEN_FILE, "w", encoding="utf-8") as handle:
        handle.write(creds.to_json())
    say(f"Saved {TOKEN_FILE}")
    return creds


# --------------------------------------------------------------
def describe_error(exc):
    """Turn an API error into one readable line."""
    status = getattr(exc, "status_code", None)
    if status is None:
        status = getattr(getattr(exc, "resp", None),
                         "status", "?")
    text = str(exc)
    if len(text) > 300:
        text = text[:300] + "..."
    return f"HTTP {status}: {text}"


# --------------------------------------------------------------
def make_deck(drive, HttpError, parent=None):
    """Create a presentation, optionally inside a folder."""
    body = {"name": TEST_NAME, "mimeType": DECK_MIME}
    if parent:
        body["parents"] = [parent]
    try:
        made = drive.files().create(
            body=body, fields="id, name, webViewLink"
        ).execute()
        return made, None
    except HttpError as exc:
        return None, describe_error(exc)


# --------------------------------------------------------------
def text_box_requests(page):
    """Build a text box and put some words in it.

    Slides rejects createShape without an explicit size and
    transform, so both carry units.
    """
    box = "accessTest01"
    return [
        {"createShape": {
            "objectId": box,
            "shapeType": "TEXT_BOX",
            "elementProperties": {
                "pageObjectId": page,
                "size": {
                    "width": {"magnitude": 3000000,
                              "unit": "EMU"},
                    "height": {"magnitude": 1000000,
                               "unit": "EMU"},
                },
                "transform": {
                    "scaleX": 1, "scaleY": 1,
                    "translateX": 500000,
                    "translateY": 500000, "unit": "EMU",
                },
            },
        }},
        {"insertText": {"objectId": box,
                        "text": "access test"}},
    ]


# --------------------------------------------------------------
def edit_deck(slides, HttpError, deck_id):
    """Write a line of text onto the deck's first slide."""
    try:
        deck = slides.presentations().get(
            presentationId=deck_id
        ).execute()
        page = deck["slides"][0]["objectId"]
        slides.presentations().batchUpdate(
            presentationId=deck_id,
            body={"requests": text_box_requests(page)},
        ).execute()
        return None
    except HttpError as exc:
        return describe_error(exc)


# --------------------------------------------------------------
def remove_deck(drive, HttpError, deck_id):
    """Delete the throwaway deck, reporting any failure."""
    try:
        drive.files().delete(fileId=deck_id).execute()
        return None
    except HttpError as exc:
        return describe_error(exc)


# --------------------------------------------------------------
def report_root(drive, HttpError):
    """Check that the token can create a deck at all."""
    say("")
    say("1. Create a deck with no parent folder")
    made, error = make_deck(drive, HttpError)
    if error:
        say(f"   FAILED  {error}")
        return None
    say(f"   ok      created {made['id']}")
    return made["id"]


# --------------------------------------------------------------
def report_folder(drive, HttpError, folder):
    """Check the question this script exists to answer."""
    say("")
    say("2. Create a deck inside your folder")
    made, error = make_deck(drive, HttpError, folder)
    if error:
        say(f"   FAILED  {error}")
        say("   => drive.file will NOT put decks in your")
        say("      folder. They would land in Drive root.")
        return None
    say(f"   ok      created {made['id']}")
    say(f"   => drive.file CAN create inside the folder")
    say(f"   {made.get('webViewLink', '')}")
    return made["id"]


# --------------------------------------------------------------
def report_edit(slides, HttpError, deck_id):
    """Check that the Slides API can edit what we made."""
    say("")
    say("3. Edit that deck through the Slides API")
    error = edit_deck(slides, HttpError, deck_id)
    if error:
        say(f"   FAILED  {error}")
        return
    say("   ok      added a text box")


# --------------------------------------------------------------
def cleanup(drive, HttpError, ids):
    """Delete every deck the test created."""
    say("")
    say("4. Delete the test decks")
    for deck_id in [i for i in ids if i]:
        error = remove_deck(drive, HttpError, deck_id)
        mark = f"FAILED  {error}" if error else "ok"
        say(f"   {mark}   {deck_id}")


# --------------------------------------------------------------
def verdict(folder, folder_id):
    """Say what the result means for the project."""
    say("")
    say("=" * 52)
    if not folder:
        say("Sign-in works. Pass a folder id to test the")
        say("rest:  python3 g_auth.py <FOLDER_ID>")
        return
    if folder_id:
        say("VERDICT: drive.file is enough. The script can")
        say("create decks straight into that folder, and")
        say("keeps access to them after you edit by hand.")
    else:
        say("VERDICT: drive.file cannot reach that folder.")
        say("Options: let decks land in Drive root and move")
        say("them yourself, or use the Picker instead.")


# --------------------------------------------------------------
def main():
    """Sign in, then measure what the token can do."""
    Request, Credentials, InstalledAppFlow, build, \
        HttpError = load_libraries()

    creds = authorize(Request, Credentials, InstalledAppFlow)
    drive = build("drive", "v3", credentials=creds)
    slides = build("slides", "v1", credentials=creds)

    folder = sys.argv[1] if len(sys.argv) > 1 else None
    if not folder:
        verdict(None, None)
        return

    root_id = report_root(drive, HttpError)
    folder_id = report_folder(drive, HttpError, folder)
    if folder_id:
        report_edit(slides, HttpError, folder_id)
    cleanup(drive, HttpError, [root_id, folder_id])
    verdict(folder, folder_id)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
