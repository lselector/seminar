#!/usr/bin/env python3
"""
Draw styled text into PowerPoint shapes.

The low level of the renderer: the deck's palette, its
font, and the handful of calls that put a styled run of
text inside a box. Both the page builders in
s3_make_pptx.py and the benchmark page in bench_page.py sit
on top of this, which is what keeps one definition of the
pale yellow fill and the red border.

Knows nothing about decks, slides, or layout. It takes a
shape and some text and styles them.

Usage:
    from pptx_text import add_textbox, style_run
    frame = add_textbox(slide, rect, boxed=True, fit=True)
    run = frame.paragraphs[0].add_run()
    run.text = "hello"
    style_run(run, 12, bold=True)

Created: 2026-09-12
Last updated: 2026-09-12
"""

import re

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

FONT = "Calibri"
COLOR_TEXT = RGBColor(0x00, 0x00, 0x00)
COLOR_LINK = RGBColor(0x05, 0x63, 0xC1)
COLOR_MUTED = RGBColor(0x59, 0x59, 0x59)
COLOR_HILITE = RGBColor(0xFF, 0xF2, 0xCC)

# Every content box wears the pale yellow fill and thin red
# border used throughout the hand-made decks. The slide
# title is left plain, as it is there.
COLOR_FILL = RGBColor(0xFF, 0xF2, 0xCC)
COLOR_BORDER = RGBColor(0xFF, 0x00, 0x00)
BORDER_WIDTH = Pt(0.75)
BOX_PAD = 0.06

# Body bullets, matching the hand-made decks: a filled dot
# with a hanging indent so wrapped lines line up under the
# first word rather than under the dot.
BULLET_CHAR = "\u25cf"
BULLET_MAR_L = 171450
BULLET_INDENT = -133350
BULLET_INSET = BULLET_MAR_L / 914400.0

# Left and right text margin inside every box.
LEFT_MARGIN_EMU = 45720

RE_MARKUP = re.compile(r"(\*\*.+?\*\*|==.+?==)")


# --------------------------------------------------------------
# --------------------------------------------------------------
def style_box(shape):
    """Paint the pale yellow fill and the red border."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_FILL
    shape.line.color.rgb = COLOR_BORDER
    shape.line.width = BORDER_WIDTH


# --------------------------------------------------------------
def add_textbox(slide, rect, boxed=False, fit=False):
    """Add a word-wrapping text box at a Rect.

    boxed paints the fill and border. fit asks PowerPoint
    to settle the height against the real text, which
    corrects any small error in our own estimate.
    """
    shape = slide.shapes.add_textbox(
        Inches(rect.x), Inches(rect.y),
        Inches(rect.w), Inches(rect.h)
    )
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = Emu(LEFT_MARGIN_EMU)
    frame.margin_right = Emu(LEFT_MARGIN_EMU)
    frame.margin_top = Emu(18288)
    frame.margin_bottom = Emu(18288)
    if boxed:
        style_box(shape)
    if fit:
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    return frame


# --------------------------------------------------------------
def style_run(run, size, bold=False, color=COLOR_TEXT):
    """Apply the deck font to one run."""
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


# --------------------------------------------------------------
def split_markup(text):
    """Split a bullet into (text, bold, hilite) pieces."""
    pieces = []
    for part in RE_MARKUP.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            pieces.append((part[2:-2], True, False))
        elif part.startswith("==") and part.endswith("=="):
            pieces.append((part[2:-2], False, True))
        else:
            pieces.append((part, False, False))
    return pieces


# --------------------------------------------------------------
def set_bullet(para, size):
    """Give a paragraph a hanging dot bullet."""
    props = para._p.get_or_add_pPr()
    props.set("marL", str(BULLET_MAR_L))
    props.set("indent", str(BULLET_INDENT))
    for tag in ("a:buClr", "a:buSzPts", "a:buFont",
                "a:buChar", "a:buNone"):
        for old in props.findall(qn(tag)):
            props.remove(old)
    colour = props.makeelement(qn("a:buClr"), {})
    colour.append(props.makeelement(
        qn("a:srgbClr"), {"val": str(COLOR_BORDER)}
    ))
    nodes = [
        colour,
        props.makeelement(
            qn("a:buSzPts"), {"val": str(int(size * 100))}
        ),
        props.makeelement(
            qn("a:buFont"), {"typeface": FONT}
        ),
        props.makeelement(
            qn("a:buChar"), {"char": BULLET_CHAR}
        ),
    ]
    for node in nodes:
        props.insert_element_before(
            node, "a:tabLst", "a:defRPr", "a:extLst"
        )


# --------------------------------------------------------------
def first_or_new(frame, used):
    """Reuse the empty first paragraph, then add more."""
    if not used:
        return frame.paragraphs[0]
    return frame.add_paragraph()


# --------------------------------------------------------------
def add_link_para(frame, text, used, size):
    """Add a paragraph holding one clickable URL."""
    para = first_or_new(frame, used)
    run = para.add_run()
    run.text = text
    style_run(run, size, color=COLOR_LINK)
    run.hyperlink.address = text


# --------------------------------------------------------------
def set_highlight(run, rgb):
    """Paint a highlight behind one run, via DrawingML."""
    props = run.font._rPr
    for old in props.findall(qn("a:highlight")):
        props.remove(old)
    node = props.makeelement(qn("a:highlight"), {})
    node.append(props.makeelement(
        qn("a:srgbClr"), {"val": str(rgb)}
    ))
    props.insert_element_before(
        node, "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill",
        "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick",
        "a:hlinkMouseOver", "a:rtl", "a:extLst"
    )


# --------------------------------------------------------------
def add_rich_para(frame, text, size, used):
    """Add a paragraph with bold and highlight markup."""
    para = first_or_new(frame, used)
    for piece, bold, hilite in split_markup(text):
        run = para.add_run()
        run.text = piece
        style_run(run, size, bold=bold)
        if hilite:
            set_highlight(run, COLOR_HILITE)
    return para


# --------------------------------------------------------------
def add_headline_para(frame, text, size, used):
    """Add the bold headline paragraph of a block."""
    para = first_or_new(frame, used)
    run = para.add_run()
    run.text = text
    style_run(run, size, bold=True)
    return para
