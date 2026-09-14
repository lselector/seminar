#!/usr/bin/env python3
"""
Give Slides a URL for a local picture, for a few seconds.

Slides accepts a picture only as a URL it can fetch itself.
Host uploads the file to the deck's Drive folder, shares it
"anyone with the link", and returns that URL. Slides copies
the picture into the presentation at insert time, so
Host.clean() then deletes every upload, which also ends the
sharing. probes/p0_api_probes.py (probe c) showed the
picture stays in the deck afterwards.

Usage:
    from gslides.host import Host
    host = Host(drive, folder_id)
    try:
        url = host.put("/tmp/clean.jpg")
        ... send the requests that use url ...
    finally:
        host.clean()

Created: 2026-09-14
Last updated: 2026-09-14
"""

import os

from gslides.client import log

UPLOAD_PREFIX = "_tmp-seminar-"


class Host:
    """Short-lived public copies of pictures on Drive."""

    # --------------------------------------
    def __init__(self, drive, folder_id):
        """Remember where uploads go."""
        self.drive = drive
        self.folder_id = folder_id
        self.uploads = []

    # --------------------------------------
    def put(self, path):
        """Upload one picture and return a fetchable URL."""
        from googleapiclient.http import MediaFileUpload
        name = UPLOAD_PREFIX + os.path.basename(path)
        media = MediaFileUpload(path, mimetype="image/jpeg")
        file_id = self.drive.files().create(
            body={"name": name, "parents": [self.folder_id]},
            media_body=media, fields="id").execute()["id"]
        self.uploads.append(file_id)
        self.drive.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
        return ("https://drive.google.com/uc?export=download"
                f"&id={file_id}")

    # --------------------------------------
    def clean(self):
        """Delete every upload, which also ends the sharing."""
        for file_id in self.uploads:
            try:
                self.drive.files().delete(
                    fileId=file_id).execute()
            except Exception as exc:
                log(f"  could not delete upload {file_id}: "
                    f"{exc}")
        self.uploads = []
