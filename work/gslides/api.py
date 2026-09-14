#!/usr/bin/env python3
"""
Build Google Slides API requests for styled text and art.

The low level of the Slides renderer: it knows how to put a
coloured box, a run of styled text, or a picture on a slide,
and nothing about decks or seminars.

Google Slides pages are 10 x 5.625 inches, so every
rectangle deck_layout computes is used as it is, converted
to EMU (914400 per inch).

Two things do not carry over. Slides has no equivalent of
PowerPoint's per-paragraph bullet glyph with a colour of its
own, so bullets are written as a literal coloured character
at the start of the line. And text is addressed by index
range rather than by run, so a box is built by inserting all
of its text at once and then styling spans of it.

Usage:
    from gslides.api import textbox, insert_text, style_span
    reqs = textbox("box1", "slide1", rect)
    reqs += insert_text("box1", "hello")
    reqs += style_span("box1", 0, 5, size=12, bold=True)

Created: 2026-09-14
Last updated: 2026-09-14
"""

EMU_PER_INCH = 914400

BLANK = "BLANK"


# --------------------------------------------------------------
def emu(inches):
    """Convert inches to the unit Slides measures in."""
    return int(round(inches * EMU_PER_INCH))


# --------------------------------------------------------------
def rgb(colour):
    """Turn an (r, g, b) byte triple into a Slides colour."""
    red, green, blue = colour
    return {"rgbColor": {
        "red": red / 255.0,
        "green": green / 255.0,
        "blue": blue / 255.0,
    }}


# --------------------------------------------------------------
def new_slide(slide_id):
    """Add one blank slide with a known object id."""
    return [{"createSlide": {
        "objectId": slide_id,
        "slideLayoutReference": {"predefinedLayout": BLANK},
    }}]


# --------------------------------------------------------------
def placement(slide_id, rect):
    """Where and how big one element sits on a slide."""
    return {
        "pageObjectId": slide_id,
        "size": {
            "width": {"magnitude": emu(rect.w),
                      "unit": "EMU"},
            "height": {"magnitude": emu(rect.h),
                       "unit": "EMU"},
        },
        "transform": {
            "scaleX": 1, "scaleY": 1,
            "translateX": emu(rect.x),
            "translateY": emu(rect.y),
            "unit": "EMU",
        },
    }


# --------------------------------------------------------------
def textbox(box_id, slide_id, rect):
    """Create an empty text box at a rectangle."""
    return [{"createShape": {
        "objectId": box_id,
        "shapeType": "TEXT_BOX",
        "elementProperties": placement(slide_id, rect),
    }}]


# --------------------------------------------------------------
def paint_box(box_id, fill=None, border=None, weight=0.75):
    """Give a box a background colour and an outline."""
    props = {}
    fields = []
    if fill:
        props["shapeBackgroundFill"] = {
            "solidFill": {"color": rgb(fill)}
        }
        fields.append("shapeBackgroundFill.solidFill.color")
    if border:
        props["outline"] = {
            "outlineFill": {
                "solidFill": {"color": rgb(border)}
            },
            "weight": {"magnitude": weight, "unit": "PT"},
            "dashStyle": "SOLID",
        }
        fields.append("outline")
    if not fields:
        return []
    return [{"updateShapeProperties": {
        "objectId": box_id,
        "shapeProperties": props,
        "fields": ",".join(fields),
    }}]


# --------------------------------------------------------------
def insert_text(box_id, text):
    """Put all of a box's text in, in one go."""
    if not text:
        return []
    return [{"insertText": {
        "objectId": box_id,
        "insertionIndex": 0,
        "text": text,
    }}]


# --------------------------------------------------------------
def style_span(box_id, start, end, size=None, bold=False,
               colour=None, link=None, font=None,
               italic=False, background=None):
    """Style one span of a box's text, by index range."""
    if end <= start:
        return []
    style = {"bold": bold, "italic": italic}
    fields = ["bold", "italic"]
    if size:
        style["fontSize"] = {"magnitude": size,
                             "unit": "PT"}
        fields.append("fontSize")
    if colour:
        style["foregroundColor"] = {
            "opaqueColor": rgb(colour)
        }
        fields.append("foregroundColor")
    if font:
        style["fontFamily"] = font
        fields.append("fontFamily")
    if link:
        style["link"] = {"url": link}
        fields.append("link")
    if background:
        style["backgroundColor"] = {
            "opaqueColor": rgb(background)
        }
        fields.append("backgroundColor")
    return [{"updateTextStyle": {
        "objectId": box_id,
        "textRange": {"type": "FIXED_RANGE",
                      "startIndex": start,
                      "endIndex": end},
        "style": style,
        "fields": ",".join(fields),
    }}]


# --------------------------------------------------------------
def hanging_indent(box_id, start, end, indent=0.1875):
    """Line up a wrapped bullet under its first word."""
    if end <= start:
        return []
    return [{"updateParagraphStyle": {
        "objectId": box_id,
        "textRange": {"type": "FIXED_RANGE",
                      "startIndex": start,
                      "endIndex": end},
        "style": {
            "indentStart": {"magnitude": emu(indent),
                            "unit": "EMU"},
            "indentFirstLine": {"magnitude": 0,
                                "unit": "EMU"},
            "spaceAbove": {"magnitude": 0, "unit": "PT"},
            "spaceBelow": {"magnitude": 0, "unit": "PT"},
        },
        "fields": ("indentStart,indentFirstLine,"
                   "spaceAbove,spaceBelow"),
    }}]


# --------------------------------------------------------------
def tighten(box_id, start, end):
    """Remove the space a paragraph adds around itself."""
    if end <= start:
        return []
    return [{"updateParagraphStyle": {
        "objectId": box_id,
        "textRange": {"type": "FIXED_RANGE",
                      "startIndex": start,
                      "endIndex": end},
        "style": {
            "spaceAbove": {"magnitude": 0, "unit": "PT"},
            "spaceBelow": {"magnitude": 0, "unit": "PT"},
        },
        "fields": "spaceAbove,spaceBelow",
    }}]


# --------------------------------------------------------------
def centre(box_id, start, end):
    """Centre a span of paragraphs in their box."""
    if end <= start:
        return []
    return [{"updateParagraphStyle": {
        "objectId": box_id,
        "textRange": {"type": "FIXED_RANGE",
                      "startIndex": start,
                      "endIndex": end},
        "style": {"alignment": "CENTER"},
        "fields": "alignment",
    }}]


# --------------------------------------------------------------
def picture(image_id, slide_id, rect, url):
    """Place a picture Google can fetch from a URL."""
    return [{"createImage": {
        "objectId": image_id,
        "url": url,
        "elementProperties": placement(slide_id, rect),
    }}]


# --------------------------------------------------------------
def outline_picture(image_id, border, weight=0.75):
    """Draw a thin border around a placed picture."""
    return [{"updateImageProperties": {
        "objectId": image_id,
        "imageProperties": {"outline": {
            "outlineFill": {
                "solidFill": {"color": rgb(border)}
            },
            "weight": {"magnitude": weight, "unit": "PT"},
            "dashStyle": "SOLID",
        }},
        "fields": "outline",
    }}]


# --------------------------------------------------------------
def set_height(object_id, size_emu, transform, height_in):
    """Change an element's height, keeping everything else.

    Slides stores a size that never changes plus a scale, so
    height is set through scaleY. Position, width, rotation
    and shear are copied from the element as it is.
    """
    base = size_emu[1]
    if base <= 0:
        return []
    move = dict(transform)
    move["unit"] = move.get("unit", "EMU")
    if move["unit"] == "PT":
        base = base / 12700.0
        move["scaleY"] = height_in * 72.0 / base
    else:
        move["scaleY"] = emu(height_in) / base
    move.setdefault("scaleX", 1)
    return [{"updatePageElementTransform": {
        "objectId": object_id,
        "transform": move,
        "applyMode": "ABSOLUTE",
    }}]


# --------------------------------------------------------------
def delete_object(object_id):
    """Remove one object, such as the starting slide."""
    return [{"deleteObject": {"objectId": object_id}}]
