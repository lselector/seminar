#!/usr/bin/env python3
"""
Read a live Google Slides deck and apply the ownership rules.

Turns the large JSON that presentations.get returns into a
few small objects (Deck, Slide, Shape) and answers the only
questions the g*_ scripts ask of it:

    is this shape the script's?          owned
    has a human frozen it with a fill?   frozen
    may a script change it right now?    editable
    which slides are in the talk, and which are parked?

The rules, from ADD.md:

    A script may change a shape only when its id starts with
    s- or t-, the box has no background fill, and it sits
    before the separator slide. A picture follows its text
    box: t-...-p is frozen when t-...-b is.

"No fill" is read from shapeBackgroundFill.propertyState.
The colour is never compared: an unfilled box still reports
white, so the colour alone cannot tell the two apart.

Everything here is pure except read_deck and find_deck, so
the rules can be tested on saved JSON with no network.

Usage:
    from gslides.deck import read_deck, find_deck, seminar_date
    deck = read_deck(slides, deck_id)
    for shape in deck.slide("s-bench").shapes:
        print(shape.id, deck.editable(shape))

Created: 2026-09-14
Last updated: 2026-09-18
"""

import datetime as dt
from dataclasses import dataclass, field
from zoneinfo import ZoneInfo

from layout import deck_layout as L
from gslides import ids as T

EMU_PER_INCH = 914400
EMU_PER_PT = 12700
DATE_KEY = "seminarDate"
DECK_MIME = "application/vnd.google-apps.presentation"

FRIDAY = 4
# The seminar ends on Friday afternoon. From this hour on a
# Friday, US Eastern time, the steps work on next week's deck.
EASTERN = ZoneInfo("America/New_York")
SWITCH_HOUR = 15
MONTHS = [
    "Jan", "Feb", "March", "April", "May", "June",
    "July", "Aug", "Sept", "Oct", "Nov", "Dec",
]

# A human's own title box starts above this line.
TITLE_BAND = L.BAND_TOP - 0.06

KIND_TEXT = "text"
KIND_IMAGE = "image"
KIND_TABLE = "table"
KIND_OTHER = "other"

# Placeholder hints written by g1_new_deck.py. A box whose
# whole text is one bracketed hint is still empty.
HINT_NEWS = "[news goes here]"
HINT_TOC = "[table of contents]"
HINT_EPIGRAPH = "[epigraph: one short line about the week]"


# --------------------------------------------------------------
def is_placeholder(text):
    """Is this text just a bracketed placeholder hint?"""
    text = (text or "").strip()
    return (not text) or (text.startswith("[")
                          and text.endswith("]")
                          and "\n" not in text)


@dataclass
class Shape:
    """One element on a slide, reduced to what matters."""

    id: str
    slide: str
    kind: str
    text: str = ""
    filled: bool = False
    rect: L.Rect = None
    size_emu: tuple = (0, 0)
    transform: dict = field(default_factory=dict)
    size_pt: float = 0.0
    cells: list = field(default_factory=list)
    styles: list = field(default_factory=list)
    description: str = ""

    # --------------------------------------
    def paragraphs(self):
        """Non-empty lines of the shape's text."""
        return [p.strip() for p in self.text.split("\n")
                if p.strip()]

    # --------------------------------------
    def first_line(self):
        """The first non-empty line, or an empty string."""
        lines = self.paragraphs()
        return lines[0] if lines else ""


@dataclass
class Slide:
    """One slide: its id, place in the deck, and shapes."""

    id: str
    index: int
    shapes: list = field(default_factory=list)
    notes_id: str = ""
    notes: str = ""

    # --------------------------------------
    def text_shapes(self):
        """Shapes holding text, top to bottom."""
        found = [s for s in self.shapes if s.kind == KIND_TEXT]
        return sorted(found, key=lambda s: (s.rect.y, s.rect.x))


@dataclass
class Deck:
    """A whole presentation, read at one revision."""

    id: str
    title: str
    revision: str
    slides: list = field(default_factory=list)

    # --------------------------------------
    def slide(self, slide_id):
        """The slide with this id, or None."""
        for slide in self.slides:
            if slide.id == slide_id:
                return slide
        return None

    # --------------------------------------
    def shape(self, object_id):
        """The shape with this id anywhere, or None."""
        for slide in self.slides:
            for shape in slide.shapes:
                if shape.id == object_id:
                    return shape
        return None

    # --------------------------------------
    def ids(self):
        """Every slide and shape id in the deck."""
        out = {s.id for s in self.slides}
        for slide in self.slides:
            out.update(shape.id for shape in slide.shapes)
        return out

    # --------------------------------------
    def separator(self):
        """Index of the separator slide, or the deck end."""
        found = self.slide(T.PAGE_SEPARATOR)
        return found.index if found else len(self.slides)

    # --------------------------------------
    def main(self):
        """Slides that are part of the talk."""
        return self.slides[:self.separator()]

    # --------------------------------------
    def parked(self):
        """Slides after the separator, never written."""
        return self.slides[self.separator() + 1:]

    # --------------------------------------
    def room_below(self, shape, bottom=L.BAND_BOT):
        """How tall a shape may grow without covering another.

        Anything on the same slide that starts lower down
        and overlaps it left to right is in the way.
        """
        limit = bottom
        home = self.slide(shape.slide)
        box = shape.rect
        for other in (home.shapes if home else []):
            near = other.rect
            if other.id == shape.id or near.y <= box.y + 0.01:
                continue
            if near.x < box.x + box.w and \
                    box.x < near.x + near.w:
                limit = min(limit, near.y - 0.06)
        return max(box.h, limit - box.y)

    # --------------------------------------
    def frozen(self, shape):
        """Has a human claimed this shape with a fill?"""
        if shape.kind == KIND_IMAGE:
            box = self.shape(T.box_of(shape.id) or "")
            return bool(box and box.filled)
        return shape.filled

    # --------------------------------------
    def editable(self, shape):
        """May a script change this shape right now?"""
        if not T.owned(shape.id) or self.frozen(shape):
            return False
        home = self.slide(shape.slide)
        return bool(home and home.index < self.separator())


# --------------------------------------------------------------
def marked_picture(page, mark):
    """Id of the picture whose alt text carries mark, or "".

    A picture you added yourself has no mark, so a step
    that refreshes a chart on a hand-made slide can never
    mistake it for the chart.
    """
    for shape in page.shapes:
        if shape.kind == KIND_IMAGE and \
                mark.lower() in shape.description.lower():
            return shape.id
    return ""


# --------------------------------------------------------------
def is_title(shape):
    """Is this box a slide title rather than content?

    Script titles end in -title. A box a human made counts
    as a title when it starts in the band above the content.
    """
    if shape.id.endswith(f"-{T.TITLE}"):
        return True
    return not T.owned(shape.id) and shape.rect.y < TITLE_BAND


# --------------------------------------------------------------
def titled_page(deck, slide_id, title):
    """A slide in the talk, by its fixed id or by its title."""
    page = deck.slide(slide_id)
    if page and page.index < deck.separator():
        return page
    want = title.lower()
    for slide in deck.main():
        for shape in slide.text_shapes():
            if is_title(shape) and \
                    shape.first_line().lower() == want:
                return slide
    return None


# --------------------------------------------------------------
def to_inches(dimension):
    """One API dimension in inches."""
    if not dimension:
        return 0.0
    value = dimension.get("magnitude", 0.0)
    if dimension.get("unit") == "PT":
        return value * EMU_PER_PT / EMU_PER_INCH
    return value / EMU_PER_INCH


# --------------------------------------------------------------
def element_rect(element):
    """Where an element sits on the page, in inches."""
    size = element.get("size", {})
    move = element.get("transform", {})
    unit = move.get("unit", "EMU")
    scale = EMU_PER_PT if unit == "PT" else 1
    width = to_inches(size.get("width"))
    height = to_inches(size.get("height"))
    return L.Rect(
        move.get("translateX", 0.0) * scale / EMU_PER_INCH,
        move.get("translateY", 0.0) * scale / EMU_PER_INCH,
        width * move.get("scaleX", 1.0),
        height * move.get("scaleY", 1.0),
    )


# --------------------------------------------------------------
def text_of(shape_json):
    """The plain text of a shape, without the final newline."""
    parts = []
    for piece in shape_json.get("text", {}).get(
            "textElements", []):
        run = piece.get("textRun") or piece.get("autoText")
        if run:
            parts.append(run.get("content", ""))
    text = "".join(parts)
    return text[:-1] if text.endswith("\n") else text


# --------------------------------------------------------------
def main_font_size(shape_json):
    """The font size most of a shape's text is set in."""
    weight = {}
    for piece in shape_json.get("text", {}).get(
            "textElements", []):
        run = piece.get("textRun")
        if not run:
            continue
        size = run.get("style", {}).get("fontSize", {}).get(
            "magnitude")
        if size:
            chars = len(run.get("content", ""))
            weight[size] = weight.get(size, 0) + chars
    if not weight:
        return 0.0
    return max(weight, key=weight.get)


# --------------------------------------------------------------
def paragraph_styles(shape_json):
    """(size, bold) most of each paragraph is set in."""
    styles, weight = [], {}
    for piece in shape_json.get("text", {}).get(
            "textElements", []):
        run = piece.get("textRun")
        if not run:
            continue
        style = run.get("style", {})
        key = (style.get("fontSize", {}).get("magnitude") or 12,
               bool(style.get("bold")))
        parts = run.get("content", "").split("\n")
        for index, part in enumerate(parts):
            if index:
                styles.append(max(weight, key=weight.get)
                              if weight else key)
                weight = {}
            if part:
                weight[key] = weight.get(key, 0) + len(part)
    if weight:
        styles.append(max(weight, key=weight.get))
    return styles


# --------------------------------------------------------------
def is_filled(shape_json):
    """Does a shape have a background fill of its own?"""
    props = shape_json.get("shapeProperties", {})
    fill = props.get("shapeBackgroundFill")
    if not fill:
        return False
    return fill.get("propertyState", "RENDERED") == "RENDERED"


# --------------------------------------------------------------
def parse_element(element, slide_id):
    """Shapes found in one page element, groups opened."""
    group = element.get("elementGroup")
    if group:
        found = []
        for child in group.get("children", []):
            found += parse_element(child, slide_id)
        return found
    size = element.get("size", {})
    shape = Shape(
        id=element["objectId"], slide=slide_id,
        description=element.get("description", ""),
        kind=KIND_OTHER, rect=element_rect(element),
        size_emu=(to_inches(size.get("width")) * EMU_PER_INCH,
                  to_inches(size.get("height")) * EMU_PER_INCH),
        transform=dict(element.get("transform", {})),
    )
    if "shape" in element:
        shape.kind = KIND_TEXT
        shape.text = text_of(element["shape"])
        shape.filled = is_filled(element["shape"])
        shape.size_pt = main_font_size(element["shape"])
        shape.styles = paragraph_styles(element["shape"])
    elif "image" in element:
        shape.kind = KIND_IMAGE
    elif "table" in element:
        shape.kind = KIND_TABLE
        shape.cells = [
            [text_of(cell) for cell in row.get("tableCells", [])]
            for row in element["table"].get("tableRows", [])
        ]
    return [shape]


# --------------------------------------------------------------
def parse_notes(slide_json):
    """The speaker notes shape id and its text."""
    page = slide_json.get("slideProperties", {}).get(
        "notesPage", {})
    notes_id = page.get("notesProperties", {}).get(
        "speakerNotesObjectId", "")
    for element in page.get("pageElements", []):
        if element.get("objectId") == notes_id:
            return notes_id, text_of(element.get("shape", {}))
    return notes_id, ""


# --------------------------------------------------------------
def parse_deck(data):
    """Build a Deck from a presentations.get response."""
    deck = Deck(id=data.get("presentationId", ""),
                title=data.get("title", ""),
                revision=data.get("revisionId", ""))
    for index, slide_json in enumerate(data.get("slides", [])):
        slide = Slide(id=slide_json["objectId"], index=index)
        for element in slide_json.get("pageElements", []):
            slide.shapes += parse_element(element, slide.id)
        slide.notes_id, slide.notes = parse_notes(slide_json)
        deck.slides.append(slide)
    return deck


# --------------------------------------------------------------
def read_deck(slides, deck_id):
    """Fetch a deck from Google and parse it."""
    data = slides.presentations().get(
        presentationId=deck_id).execute()
    return parse_deck(data)


# --------------------------------------------------------------
def eastern_now(now=None):
    """now as a US Eastern datetime; a bare date is midnight."""
    if now is None:
        return dt.datetime.now(EASTERN)
    if not isinstance(now, dt.datetime):
        return dt.datetime.combine(now, dt.time(),
                                   tzinfo=EASTERN)
    if now.tzinfo is None:
        return now.replace(tzinfo=EASTERN)
    return now.astimezone(EASTERN)


# --------------------------------------------------------------
def next_friday(now=None):
    """The Friday the steps work on.

    Today on a Friday before 3 pm US Eastern (the seminar
    has not ended); from 3 pm on, the Friday after. Any
    other day, the coming Friday.
    """
    now = eastern_now(now)
    days = (FRIDAY - now.weekday()) % 7
    if days == 0 and now.hour >= SWITCH_HOUR:
        days = 7
    return now.date() + dt.timedelta(days=days)


# --------------------------------------------------------------
def seminar_date(text=None, now=None):
    """The date given as YYYY-MM-DD, else next_friday."""
    if text:
        return dt.date.fromisoformat(text)
    return next_friday(now)


# --------------------------------------------------------------
def deck_title(date):
    """The deck title in the style of the archive."""
    month = MONTHS[date.month - 1]
    return f"AI News - {month} {date.day}, {date.year}"


# --------------------------------------------------------------
def deck_name(date):
    """The Drive file name for one seminar."""
    return f"{date.isoformat()}-AI-News"


# --------------------------------------------------------------
def find_decks(drive, date):
    """Every deck the app made for one seminar date."""
    query = (f"appProperties has {{ key='{DATE_KEY}' and "
             f"value='{date.isoformat()}' }} and "
             f"mimeType='{DECK_MIME}' and trashed=false")
    got = drive.files().list(
        q=query, orderBy="createdTime desc",
        fields="files(id,name,webViewLink,createdTime)",
    ).execute()
    return got.get("files", [])


# --------------------------------------------------------------
def latest_before(files, date):
    """Of dated deck files, the newest seminar before date."""
    dated = []
    for item in files:
        text = item.get("appProperties", {}).get(DATE_KEY, "")
        try:
            day = dt.date.fromisoformat(text)
        except ValueError:
            continue
        if day < date:
            dated.append((day, item.get("createdTime", ""),
                          item))
    return max(dated, key=lambda d: d[:2])[2] if dated else None


# --------------------------------------------------------------
def find_previous_deck(drive, date):
    """The deck of the latest seminar before date, or None."""
    got = drive.files().list(
        q=f"mimeType='{DECK_MIME}' and trashed=false",
        pageSize=1000,
        fields="files(id,name,appProperties,createdTime)",
    ).execute()
    return latest_before(got.get("files", []), date)
