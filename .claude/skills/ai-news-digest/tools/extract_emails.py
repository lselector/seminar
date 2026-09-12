"""
Extract plain text from recent messages in a Postbox
(Thunderbird-style) mbox mail folder.

Postbox stores each mail folder as a single mbox file.
This tool splits that file on the Postbox separator line
("From - <date>"), keeps messages from the last N days,
converts HTML bodies to plain text, and collects article
links together with their anchor text.

It writes a plain-text file (one block per message) and
prints a one-line JSON summary on stdout so the caller
can name downstream files from the real date range.

Usage:
  python3 extract_emails.py
  python3 extract_emails.py --days 7 --out extract.txt
  python3 extract_emails.py --mbox /path/to/mbox --days 14
  python3 extract_emails.py --days 7 --max-chars 8000

Notes:
  Reading is read-only and safe while Postbox runs, but a
  live folder may end mid-message; that block is skipped.
  Do NOT parse these files with mailbox.mbox: it splits on
  any "From " line and shreds newsletters into fragments.

Created: 2026-09-02
Last updated: 2026-09-02
"""

import argparse
import base64
import datetime as dt
import email
import email.header
import json
import re
import sys
from html.parser import HTMLParser
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import unquote

DEFAULT_MBOX = (
    "~/Library/Application Support/PostboxApp/Profiles/"
    "mfxvr37h.default/ImapMail/imap.gmail.com/"
    "my_gmail_folders.sbd/___Medium_Twitter"
)

# Postbox writes this exact separator before each message.
SEPARATOR = re.compile(rb"(?m)^From - .*\r?\n")

URL_RE = re.compile(r"https?://[^\s\"'<>)\]]+")

# Hosts that only ever serve images, CSS, or click
# tracking, never an article worth citing. Wrappers are
# checked only after unwrap_tracking() has had a try, so
# what remains here is genuinely opaque.
SKIP_HOSTS = (
    "cdn-cgi",
    "substackcdn.com",
    "media.beehiiv.com",
    "storage.googleapis.com",
    "sendgrid.net",
    "list-manage.com",
    "mailchimp.com",
    "google-analytics.com",
    "doubleclick.net",
    "link.mail.beehiiv.com",
    "app.alphasignal.ai",
    "tracking.tldrnewsletter.com",
    "link.sbstck.com",
    "substack.com/redirect",
    "hp.beehiiv.com",
    "beehiivstatus.com",
)

# Tracking wrappers that hide the real target inside the
# link. Both forms are recoverable, so decode before
# deciding a URL is noise.
SUBSTACK_RE = re.compile(
    r"https://substack\.com/redirect/2/([\w-]+)")

# The colon and the slashes are encoded independently by
# different senders, so allow every mix of the two.
EMBEDDED_RE = re.compile(
    r"https?(:|%3A)(//|%2F%2F)", re.I)

# Substrings marking opt-out, sharing, and account links.
SKIP_WORDS = (
    "unsubscribe",
    "manage-preferences",
    "/preferences",
    "email-settings",
    "opt_out",
    "optout",
    "/privacy",
    "/terms",
    "twitter.com/intent",
    "/share?",
    "/app-link/",
    "substack.com/signup",
    "redirect=app-store",
    "disable_email",
    "referrer_token",
    "action=restack",
)

ASSET_EXT = (
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg",
    ".css", ".js", ".ico",
)

# Query parameters that carry no meaning for a reader.
DROP_PARAMS = {
    "ref", "source", "r", "fbclid", "gclid", "mc_cid",
    "mc_eid", "_bhlid", "triedsigningin",
}

BREAK_TAGS = ("p", "br", "div", "tr", "li", "h1", "h2",
              "h3", "h4", "blockquote")


# --------------------------------------------------------------
class HtmlText(HTMLParser):
    """Collect visible text and <a href> links from HTML."""

    SKIP_TAGS = {"script", "style", "head", "title"}

    # --------------------------------------
    def __init__(self):
        """Set up empty text and link buffers."""
        super().__init__(convert_charrefs=True)
        self.chunks = []
        self.links = []
        self._skip = 0
        self._href = None
        self._anchor = []

    # --------------------------------------
    def handle_starttag(self, tag, attrs):
        """Track skipped tags, breaks, and open anchors."""
        if tag in self.SKIP_TAGS:
            self._skip += 1
        elif tag == "a":
            self._href = dict(attrs).get("href")
            self._anchor = []
        elif tag in BREAK_TAGS:
            self.chunks.append("\n")

    # --------------------------------------
    def handle_endtag(self, tag):
        """Close skipped tags and finish open anchors."""
        if tag in self.SKIP_TAGS and self._skip:
            self._skip -= 1
        elif tag == "a" and self._href:
            label = "".join(self._anchor)
            self.links.append((" ".join(label.split()),
                               self._href))
            self._href = None

    # --------------------------------------
    def handle_data(self, data):
        """Buffer visible text and anchor text."""
        if self._skip:
            return
        self.chunks.append(data)
        if self._href is not None:
            self._anchor.append(data)


# --------------------------------------------------------------
def clean_text(text):
    """Collapse runs of spaces and of blank lines."""
    out = []
    for line in text.splitlines():
        line = re.sub(r"[ \t\u00a0\u200b\u200c\ufeff]+",
                      " ", line)
        line = line.strip()
        if line or (out and out[-1]):
            out.append(line)
    return "\n".join(out).strip()


# --------------------------------------------------------------
def html_to_text(html):
    """Return (plain_text, links) parsed from HTML."""
    parser = HtmlText()
    try:
        parser.feed(html)
        parser.close()
    except Exception:
        pass
    return clean_text("".join(parser.chunks)), parser.links


# --------------------------------------------------------------
def split_mbox(raw):
    """Split raw mbox bytes on the Postbox separator."""
    return [p for p in SEPARATOR.split(raw) if p.strip()]


# --------------------------------------------------------------
def decode_field(value):
    """Decode a MIME header into a plain string."""
    if not value:
        return ""
    out = []
    for chunk, charset in email.header.decode_header(value):
        if isinstance(chunk, bytes):
            chunk = chunk.decode(charset or "utf-8", "replace")
        out.append(chunk)
    return " ".join("".join(out).split())


# --------------------------------------------------------------
def msg_date(msg):
    """Return the message date in UTC, or None."""
    try:
        when = parsedate_to_datetime(msg.get("Date"))
    except Exception:
        return None
    if when is None:
        return None
    if when.tzinfo is None:
        when = when.replace(tzinfo=dt.timezone.utc)
    return when.astimezone(dt.timezone.utc)


# --------------------------------------------------------------
def pick_part(msg, subtype):
    """Return decoded text of the first matching part."""
    for part in msg.walk():
        if part.get_content_subtype() != subtype:
            continue
        if "attachment" in str(part.get(
                "Content-Disposition")):
            continue
        raw = part.get_payload(decode=True)
        if not raw:
            continue
        charset = part.get_content_charset() or "utf-8"
        return raw.decode(charset, "replace")
    return ""


# --------------------------------------------------------------
def body_and_links(msg):
    """Return (plain_text, links) for one message."""
    html = pick_part(msg, "html")
    if html:
        return html_to_text(html)
    plain = pick_part(msg, "plain")
    links = [("", u) for u in URL_RE.findall(plain)]
    return clean_text(plain), links


# --------------------------------------------------------------
def decode_substack(url):
    """Return the target inside a Substack redirect link."""
    match = SUBSTACK_RE.match(url)
    if not match:
        return ""
    blob = match.group(1)
    blob += "=" * (-len(blob) % 4)
    try:
        raw = base64.urlsafe_b64decode(blob)
        return json.loads(raw).get("e", "")
    except Exception:
        return ""


# --------------------------------------------------------------
def unwrap_tracking(url):
    """Recover the destination behind a tracker wrapper."""
    inner = decode_substack(url)
    if inner.startswith("http"):
        return inner
    # TLDR and similar wrappers keep the real target in
    # the path, plain or percent-encoded.
    tail = url[8:]
    match = EMBEDDED_RE.search(tail)
    if not match:
        return url
    return unquote(tail[match.start():])


# --------------------------------------------------------------
def strip_tracking(url):
    """Drop utm_* and similar query parameters."""
    base, _, query = url.partition("?")
    if not query:
        return url
    keep = []
    for pair in query.split("&"):
        name = pair.split("=")[0].lower()
        if name.startswith("utm_") or name in DROP_PARAMS:
            continue
        keep.append(pair)
    return base + ("?" + "&".join(keep) if keep else "")


# --------------------------------------------------------------
def is_noise(url):
    """True if a URL is an asset, tracker, or opt-out."""
    low = url.lower()
    if not low.startswith("http"):
        return True
    if any(host in low for host in SKIP_HOSTS):
        return True
    if any(word in low for word in SKIP_WORDS):
        return True
    return low.split("?")[0].endswith(ASSET_EXT)


# --------------------------------------------------------------
def clean_links(links, limit):
    """Filter, de-duplicate, and cap a list of links."""
    seen = set()
    out = []
    for label, url in links:
        url = strip_tracking(unwrap_tracking(url.strip()))
        if is_noise(url) or url in seen:
            continue
        seen.add(url)
        out.append((label[:90], url))
        if len(out) >= limit:
            break
    return out


# --------------------------------------------------------------
def build_record(msg, when, max_chars, max_links):
    """Return a dict describing one extracted message."""
    text, links = body_and_links(msg)
    if len(text) > max_chars:
        text = text[:max_chars] + "\n[...truncated...]"
    return {
        "date": when,
        "sender": decode_field(msg.get("From")),
        "subject": decode_field(msg.get("Subject")),
        "text": text,
        "links": clean_links(links, max_links),
    }


# --------------------------------------------------------------
def read_dated(path):
    """Return [(date, message)] for a whole mbox file."""
    raw = Path(path).expanduser().read_bytes()
    dated = []
    for block in split_mbox(raw):
        msg = email.message_from_bytes(block)
        when = msg_date(msg)
        if when is not None:
            dated.append((when, msg))
    return dated


# --------------------------------------------------------------
def collect(path, days, max_chars, max_links):
    """Extract records for messages from the last N days."""
    dated = read_dated(path)
    if not dated:
        return []
    # Window is relative to the newest message, so a
    # stale mailbox still yields its most recent week.
    cutoff = max(w for w, _ in dated) - dt.timedelta(
        days=days)
    return [
        build_record(msg, when, max_chars, max_links)
        for when, msg in sorted(dated, key=lambda r: r[0])
        if when >= cutoff
    ]


# --------------------------------------------------------------
def render(records):
    """Render records as one plain-text document."""
    out = []
    for num, rec in enumerate(records, 1):
        out.append(f"=== MESSAGE {num} ===")
        out.append("DATE: " + rec["date"].strftime(
            "%Y-%m-%d %H:%M UTC"))
        out.append(f"FROM: {rec['sender']}")
        out.append(f"SUBJECT: {rec['subject']}")
        out.append("--- TEXT ---")
        out.append(rec["text"])
        if rec["links"]:
            out.append("--- LINKS ---")
            for label, url in rec["links"]:
                out.append(f"- {label or '(no text)'} | {url}")
        out.append("")
    return "\n".join(out)


# --------------------------------------------------------------
def parse_args():
    """Define and parse command-line arguments."""
    ap = argparse.ArgumentParser(
        description="Extract recent mbox messages as text.")
    ap.add_argument("--mbox", default=DEFAULT_MBOX,
                    help="Path to the mbox file.")
    ap.add_argument("--days", type=int, default=7,
                    help="Days back from the newest message.")
    ap.add_argument("--out", default="email_extract.txt",
                    help="Where to write the plain text.")
    ap.add_argument("--max-chars", type=int, default=5000,
                    help="Body characters kept per message.")
    ap.add_argument("--max-links", type=int, default=12,
                    help="Links kept per message.")
    return ap.parse_args()


# --------------------------------------------------------------
def main():
    """Extract messages and print a JSON summary."""
    args = parse_args()
    try:
        records = collect(args.mbox, args.days,
                          args.max_chars, args.max_links)
    except OSError as err:
        print(json.dumps({"count": 0, "error": str(err)}))
        return 1
    if not records:
        print(json.dumps({"count": 0,
                          "error": "no dated messages"}))
        return 1
    out_path = Path(args.out)
    out_path.write_text(render(records), encoding="utf-8")
    print(json.dumps({
        "count": len(records),
        "first_date": records[0]["date"].strftime("%Y-%m-%d"),
        "last_date": records[-1]["date"].strftime("%Y-%m-%d"),
        "out": str(out_path),
        "chars": out_path.stat().st_size,
    }))
    return 0


# --------------------------------------------------------------
if __name__ == "__main__":
    sys.exit(main())
