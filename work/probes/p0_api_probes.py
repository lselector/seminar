#!/usr/bin/env python3
"""
API probes behind ADD.md, run against a scratch deck.

Each probe answers one question the new workflow depends on,
prints PASS or FAIL with the evidence, and cleans up after
itself. Nothing here touches a real seminar deck.

    c  Drive-hosted picture, shared briefly, then unshared
    d  A write with a stale revision id is rejected
    e  Delete and recreate the same object id in one batch
    f  appProperties can be set and searched
    g  Speaker notes can be written and read back

Usage:
    python3 probes/p0_api_probes.py

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from layout import deck_layout as L
from gslides import api as S
from gslides.client import credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

FOLDER = "1XFo7T2X7OYR0T-qf_SFuyRDxZ9cLmFNN"
DECK_MIME = "application/vnd.google-apps.presentation"
PICTURE = os.path.join(os.path.dirname(HERE), "data", "images",
                       "meta-muse.jpg")


# --------------------------------------------------------------
def say(tag, ok, detail):
    """Print one probe verdict."""
    print(f"[{tag}] {'PASS' if ok else 'FAIL'}  {detail}")


# --------------------------------------------------------------
def scratch(drive, slides):
    """Make a scratch deck with one blank slide of ours."""
    deck = drive.files().create(body={
        "name": "DELETE ME - step 0 probes",
        "mimeType": DECK_MIME, "parents": [FOLDER],
        "appProperties": {"seminarDate": "1999-01-01",
                          "seminarProbe": "yes"},
    }, fields="id").execute()["id"]
    reqs = list(S.new_slide("s-probe"))
    reqs += S.textbox("t-probe-box-b", "s-probe",
                      L.Rect(1, 1, 4, 1))
    reqs += S.insert_text("t-probe-box-b", "first text")
    slides.presentations().batchUpdate(
        presentationId=deck, body={"requests": reqs}).execute()
    return deck


# --------------------------------------------------------------
def upload_shared(drive):
    """Upload the test picture and share it by link."""
    media = MediaFileUpload(PICTURE, mimetype="image/jpeg")
    pic = drive.files().create(
        body={"name": "DELETE ME probe.jpg",
              "parents": [FOLDER]},
        media_body=media, fields="id").execute()["id"]
    perm = drive.permissions().create(
        fileId=pic, body={"type": "anyone", "role": "reader"},
        fields="id").execute()["id"]
    return pic, perm


# --------------------------------------------------------------
def insert_first_accepted(slides, deck, pic):
    """Try Drive URL forms until Slides takes one."""
    urls = [
        f"https://drive.google.com/uc?export=download&id={pic}",
        f"https://lh3.googleusercontent.com/d/{pic}",
        f"https://drive.google.com/uc?id={pic}",
    ]
    for n, url in enumerate(urls):
        try:
            slides.presentations().batchUpdate(
                presentationId=deck, body={"requests":
                    S.picture(f"t-probe-pic{n}-p", "s-probe",
                              L.Rect(5, 1, 3, 2), url)}
            ).execute()
            return n
        except HttpError as exc:
            print(f"     url form {n} rejected: "
                  f"{str(exc)[-120:]}")
    return None


# --------------------------------------------------------------
def picture_survives(slides, deck, n):
    """Is the picture still there, and what does it show?"""
    import urllib.request
    got = slides.presentations().get(
        presentationId=deck).execute()
    page = next(s for s in got["slides"]
                if s["objectId"] == "s-probe")
    still = any(e["objectId"] == f"t-probe-pic{n}-p"
                for e in page["pageElements"])
    thumb = slides.presentations().pages().getThumbnail(
        presentationId=deck, pageObjectId="s-probe",
        thumbnailProperties_thumbnailSize="SMALL").execute()
    out = os.path.join("/tmp", "probe_c_thumb.png")
    with open(out, "wb") as handle:
        handle.write(urllib.request.urlopen(
            thumb["contentUrl"]).read())
    return still, out


# --------------------------------------------------------------
def probe_c(drive, slides, deck):
    """Insert a Drive-hosted picture, then unshare it."""
    pic, perm = upload_shared(drive)
    used = insert_first_accepted(slides, deck, pic)
    drive.permissions().delete(
        fileId=pic, permissionId=perm).execute()
    drive.files().delete(fileId=pic).execute()
    if used is None:
        say("c", False, "no Drive URL form was accepted")
        return
    time.sleep(2)
    still, out = picture_survives(slides, deck, used)
    say("c", still,
        f"url form {used} accepted; picture still in deck "
        f"after unsharing and deleting the Drive file: "
        f"{still}; thumbnail saved to {out}")


# --------------------------------------------------------------
def probe_d(slides, deck):
    """A stale revision id must be refused."""
    got = slides.presentations().get(
        presentationId=deck).execute()
    stale = got["revisionId"]
    slides.presentations().batchUpdate(
        presentationId=deck, body={"requests":
            S.insert_text("t-probe-box-b", "X")}).execute()
    try:
        slides.presentations().batchUpdate(
            presentationId=deck, body={
                "requests": S.insert_text("t-probe-box-b", "Y"),
                "writeControl": {"requiredRevisionId": stale},
            }).execute()
        say("d", False, "stale write was accepted")
    except HttpError as exc:
        status = getattr(exc.resp, "status", "?")
        say("d", True, f"stale write refused, HTTP {status}: "
            f"{str(exc)[-150:]}")


# --------------------------------------------------------------
def probe_e(slides, deck):
    """Recreate an object under its own id in one batch."""
    reqs = S.delete_object("t-probe-box-b")
    reqs += S.textbox("t-probe-box-b", "s-probe",
                      L.Rect(1, 3, 4, 1))
    reqs += S.insert_text("t-probe-box-b", "recreated")
    try:
        slides.presentations().batchUpdate(
            presentationId=deck, body={"requests": reqs}
        ).execute()
        say("e", True, "delete + create same id in one batch")
    except HttpError as exc:
        say("e", False, str(exc)[-200:])


# --------------------------------------------------------------
def probe_f(drive, deck):
    """Find the deck by its hidden tag."""
    q = ("appProperties has { key='seminarDate' and "
         "value='1999-01-01' } and trashed=false")
    found = drive.files().list(
        q=q, fields="files(id,name)").execute()["files"]
    say("f", [f["id"] for f in found] == [deck],
        f"search by appProperties found {len(found)} file(s)")


# --------------------------------------------------------------
def probe_g(slides, deck):
    """Write the speaker notes, then read them back."""
    got = slides.presentations().get(
        presentationId=deck).execute()
    notes = (got["slides"][0]["slideProperties"]["notesPage"]
             ["notesProperties"]["speakerNotesObjectId"])
    text = "ledger:\nt-one-1234\nt-two-5678"
    slides.presentations().batchUpdate(
        presentationId=deck, body={"requests":
            S.insert_text(notes, text)}).execute()
    got = slides.presentations().get(
        presentationId=deck).execute()
    page = got["slides"][0]["slideProperties"]["notesPage"]
    back = ""
    for el in page["pageElements"]:
        if el["objectId"] == notes:
            for run in el["shape"]["text"]["textElements"]:
                back += run.get("textRun", {}).get("content", "")
    say("g", back.strip() == text,
        f"notes id {notes!r}, read back {back.strip()!r}")


# --------------------------------------------------------------
def main():
    """Run every probe on one scratch deck."""
    creds = credentials()
    drive = build("drive", "v3", credentials=creds)
    slides = build("slides", "v1", credentials=creds)
    deck = scratch(drive, slides)
    try:
        probe_c(drive, slides, deck)
        probe_d(slides, deck)
        probe_e(slides, deck)
        probe_f(drive, deck)
        probe_g(slides, deck)
    finally:
        drive.files().delete(fileId=deck).execute()
        print("scratch deck removed")


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
