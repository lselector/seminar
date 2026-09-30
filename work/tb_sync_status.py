#!/usr/bin/env python3
"""
Show which Thunderbird IMAP folders are fully downloaded.

For every folder ticked for offline use, compares the number
of messages the server has (Thunderbird's folder cache) with
the number stored on this Mac (message separators in the
folder's mbox file), and prints a table:

    Account  Folder  On Gmail  On the Mac  Missing

By default only folders with missing messages are listed;
--all lists every ticked folder. "On the Mac" can be higher
than the server count: mail that was moved or deleted stays
in the file until File > Compact Folders, and a few messages
contain a "From - " line in their text. So "Missing" is a
lower bound.

--unticked lists the folders NOT ticked for offline use,
largest first, to help choose what to download next:

    Account  Folder  On Gmail  Est. size MB  On the Mac MB

Thunderbird does not know the size of a folder it has not
downloaded, so "Est. size" is the size of the same folder
in the old Postbox profile (downloaded before 2026-09-28).
It is an estimate: mail deleted since then (the daily Gmail
cleanup) makes the real size smaller. "?" means Postbox had
no copy of that folder.

Read-only; safe while Thunderbird runs.

Usage:
    python3 tb_sync_status.py
    python3 tb_sync_status.py --all
    python3 tb_sync_status.py --unticked
    python3 tb_sync_status.py --profile PROFILE_DIR

Created: 2026-09-30
Last updated: 2026-09-30
"""

import argparse
import configparser
import json
import os
import re

TB_ROOT = os.path.expanduser("~/Library/Thunderbird")
PB_ROOT = os.path.expanduser(
    "~/Library/Application Support/PostboxApp/Profiles")
MB = 1024 * 1024
OFFLINE_FLAG = 0x08000000
SEPARATOR = b"From - "
CHUNK = 8 * 1024 * 1024
RE_PREF = re.compile(
    r'user_pref\("mail\.server\.(server\d+)\.'
    r'(directory-rel|userName)", "([^"]*)"\);')


# --------------------------------------------------------------
def default_profile():
    """Path of the default Thunderbird profile."""
    ini = configparser.ConfigParser()
    ini.read(os.path.join(TB_ROOT, "profiles.ini"))
    for name in ini.sections():
        if name.startswith("Install"):
            rel = ini[name].get("Default")
            if rel:
                return os.path.join(TB_ROOT, rel)
    for name in ini.sections():
        if ini[name].get("Default") == "1":
            return os.path.join(TB_ROOT, ini[name]["Path"])
    raise SystemExit("No default Thunderbird profile found")


# --------------------------------------------------------------
def account_names(profile):
    """IMAP directory name -> account user name, from prefs."""
    servers = {}
    with open(os.path.join(profile, "prefs.js"),
              encoding="utf-8") as handle:
        for line in handle:
            match = RE_PREF.match(line.strip())
            if match:
                key, field, value = match.groups()
                servers.setdefault(key, {})[field] = value
    names = {}
    for fields in servers.values():
        rel = fields.get("directory-rel", "")
        if rel.startswith("[ProfD]ImapMail/"):
            user = fields.get("userName", "")
            names[rel.split("/")[-1]] = user
    return names


# --------------------------------------------------------------
def count_messages(path):
    """Number of message separators in an mbox file."""
    if not os.path.isfile(path):
        return 0
    count, tail = 0, b"\n"
    with open(path, "rb") as handle:
        while True:
            block = handle.read(CHUNK)
            if not block:
                return count
            data = tail + block
            count += data.count(b"\n" + SEPARATOR)
            tail = data[-len(SEPARATOR):]


# --------------------------------------------------------------
def cached_folders(profile, ticked):
    """(mbox path, server count) of IMAP folders, ticked or not.

    The server count is -1 when Thunderbird never opened the
    folder.
    """
    with open(os.path.join(profile, "folderCache.json"),
              encoding="utf-8") as handle:
        cache = json.load(handle)
    found = []
    for path, entry in cache.items():
        if "/ImapMail/" not in path:
            continue
        if bool(entry.get("flags", 0) & OFFLINE_FLAG) != ticked:
            continue
        mbox = path[:-4] if path.endswith(".msf") else path
        found.append((mbox, entry.get("totalMsgs", -1)))
    return found


# --------------------------------------------------------------
def split_path(mbox):
    """(IMAP directory, folder name) of a folder's mbox path."""
    rel = mbox.split("/ImapMail/", 1)[1]
    directory, _, folder = rel.partition("/")
    return directory, folder.replace(".sbd/", "/")


# --------------------------------------------------------------
def build_rows(profile):
    """(account, folder, server, local, missing) per folder."""
    names = account_names(profile)
    rows = []
    for mbox, server in cached_folders(profile, True):
        if server <= 0:
            continue
        directory, folder = split_path(mbox)
        local = count_messages(mbox)
        rows.append((names.get(directory, directory), folder,
                     server, local, max(0, server - local)))
    return sorted(rows, key=lambda r: (-r[4], r[0], r[1]))


# --------------------------------------------------------------
def postbox_profile():
    """Path of the old Postbox profile, or None."""
    if not os.path.isdir(PB_ROOT):
        return None
    for name in sorted(os.listdir(PB_ROOT)):
        path = os.path.join(PB_ROOT, name)
        if os.path.isfile(os.path.join(path, "prefs.js")):
            return path
    return None


# --------------------------------------------------------------
def postbox_size(postbox, pb_dirs, user, folder):
    """MB of the same folder in Postbox, or None.

    Postbox named some Gmail folders with a "-1" suffix, such
    as "[Gmail]/Sent Mail-1".
    """
    directory = pb_dirs.get(user)
    if not (postbox and directory):
        return None
    base = os.path.join(postbox, "ImapMail", directory,
                        folder.replace("/", ".sbd/"))
    for path in (base, base + "-1"):
        if os.path.isfile(path):
            return round(os.path.getsize(path) / MB)
    return None


# --------------------------------------------------------------
def build_unticked_rows(profile):
    """(account, folder, server, est. MB, Mac MB) per folder."""
    names = account_names(profile)
    postbox = postbox_profile()
    pb_dirs = {}
    if postbox:
        pb_dirs = {user: directory for directory, user
                   in account_names(postbox).items()}
    rows = []
    for mbox, server in cached_folders(profile, False):
        directory, folder = split_path(mbox)
        user = names.get(directory, directory)
        est = postbox_size(postbox, pb_dirs, user, folder)
        on_mac = os.path.getsize(mbox) / MB \
            if os.path.isfile(mbox) else 0
        if server > 0 or est:
            rows.append((user, folder,
                         server if server >= 0 else "?",
                         "?" if est is None else est,
                         round(on_mac)))
    return sorted(rows, key=lambda r: -(r[3] if r[3] != "?"
                                        else -1))


# --------------------------------------------------------------
def print_rows(head, rows):
    """Print rows under a header as an aligned text table."""
    count = len(head)
    width = [max([len(head[i])] + [len(str(r[i]))
                                   for r in rows])
             for i in range(count)]
    fmt = "  ".join(
        f"{{:{'<' if i < 2 else '>'}{width[i]}}}"
        for i in range(count))
    print(fmt.format(*head))
    print(fmt.format(*("-" * w for w in width)))
    for row in rows:
        print(fmt.format(*row))


# --------------------------------------------------------------
def print_unticked(rows):
    """Print the unticked-folder table and its total."""
    print_rows(("Account", "Folder", "On Gmail",
                "Est. size MB", "On the Mac MB"), rows)
    sized = [r for r in rows if r[3] != "?"]
    gmail = sum(r[3] for r in sized
                if r[1].startswith("[Gmail]"))
    labels = sum(r[3] for r in sized) - gmail
    print(f"\n{len(rows)} unticked folders. Postbox estimate: "
          f"{labels / 1024:.1f} GB in your own folders, plus "
          f"{gmail / 1024:.1f} GB in [Gmail] folders (All Mail "
          f"repeats every message)")


# --------------------------------------------------------------
def print_table(rows, show_all):
    """Print the download-status table and its summary."""
    shown = rows if show_all else [r for r in rows if r[4]]
    if shown:
        print_rows(("Account", "Folder", "On Gmail",
                    "On the Mac", "Missing"), shown)
    missing = sum(r[4] for r in rows)
    done = sum(1 for r in rows if not r[4])
    short = len(rows) - done
    print(f"\n{len(rows)} offline folders: {done} complete, "
          f"{short} short, ~{missing} messages missing")


# --------------------------------------------------------------
def main():
    """Parse arguments and print the sync status table."""
    parser = argparse.ArgumentParser(
        description="Thunderbird offline download status")
    parser.add_argument("--all", action="store_true",
                        help="also list complete folders")
    parser.add_argument("--unticked", action="store_true",
                        help="list folders not kept offline,"
                             " by estimated size")
    parser.add_argument("--profile", default=None,
                        help="Thunderbird profile directory")
    args = parser.parse_args()
    profile = os.path.expanduser(args.profile) \
        if args.profile else default_profile()
    if args.unticked:
        print_unticked(build_unticked_rows(profile))
    else:
        print_table(build_rows(profile), args.all)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
