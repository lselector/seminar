#!/usr/bin/env python3
"""
Render a Markdown file to PDF.

Converts GUIDE_to_M92.md to GUIDE_to_M92.pdf using the
Python markdown library for the HTML and WeasyPrint for the
page layout. Styling lives in styles.css, not in here, so
the look can be changed without editing Python. The page is
US Letter with 0.7 inch margins, set in the stylesheet.

Heading anchors are slugified the same way GitHub does it,
so the links in the table of contents stay clickable in the
finished PDF.

Images are resolved relative to the Markdown file, so run
s1_download_images.py and s2_clean_images.py first.

Usage:
    python3 s3_make_pdf.py
    python3 s3_make_pdf.py README.md
    python3 s3_make_pdf.py README.md out/readme.pdf

Created: 2026-09-04
Last updated: 2026-09-10
"""

import os
import re
import sys
from datetime import datetime

DEFAULT_INPUT = "GUIDE_to_M92.md"
STYLESHEET = "styles.css"
EXTENSIONS = [
    'tables', 'fenced_code', 'sane_lists', 'attr_list', 'toc'
]


# --------------------------------------------------------------
def log_message(message):
    """Print timestamped log message."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")


# --------------------------------------------------------------
def github_slug(text, separator="-"):
    """Slugify a heading the way GitHub anchors do."""
    slug = text.strip().lower()
    slug = re.sub(r"[^\w\s-]", "", slug)
    return re.sub(r"\s+", separator, slug)


# --------------------------------------------------------------
def import_dependencies():
    """Import markdown and weasyprint, or exit with help."""
    try:
        import markdown
        from weasyprint import CSS, HTML
        return markdown, HTML, CSS
    except ImportError as exc:
        log_message(f"ERROR: missing dependency: {exc}")
        log_message("Install: pip install markdown weasyprint")
        sys.exit(1)


# --------------------------------------------------------------
def markdown_to_html(markdown, text):
    """Convert Markdown text to an HTML fragment."""
    converter = markdown.Markdown(
        extensions=EXTENSIONS,
        extension_configs={
            'toc': {'slugify': github_slug, 'permalink': False}
        }
    )
    return converter.convert(text)


# --------------------------------------------------------------
def wrap_html(body, title):
    """Wrap an HTML fragment in a minimal document."""
    return (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n<head>\n'
        '<meta charset="utf-8">\n'
        f"<title>{title}</title>\n"
        "</head>\n<body>\n"
        f"{body}\n"
        "</body>\n</html>\n"
    )


# --------------------------------------------------------------
def document_title(text, fallback):
    """Take the title from the first level-one heading."""
    for line in text.splitlines():
        if line.startswith("# "):
            return line[2:].strip()
    return fallback


# --------------------------------------------------------------
def resolve_paths(argv):
    """Work out the input and output file paths."""
    source = argv[1] if len(argv) > 1 else DEFAULT_INPUT

    if not os.path.isfile(source):
        log_message(f"ERROR: {source} not found")
        sys.exit(1)

    if len(argv) > 2:
        target = argv[2]
    else:
        target = os.path.splitext(source)[0] + ".pdf"

    return source, target


# --------------------------------------------------------------
def build_document(html_text, base_url, HTML, CSS):
    """Lay out the HTML, returning a rendered document."""
    stylesheets = []
    if os.path.isfile(STYLESHEET):
        stylesheets.append(CSS(filename=STYLESHEET))
    else:
        log_message(
            f"WARNING: {STYLESHEET} not found, "
            f"using browser defaults"
        )

    document = HTML(string=html_text, base_url=base_url)
    return document.render(stylesheets=stylesheets)


# --------------------------------------------------------------
def report_result(target, pages):
    """Log the finished file, its size and page count."""
    size_kb = os.path.getsize(target) // 1024
    log_message(f"Wrote {target}")
    log_message(f"Pages: {pages}, size: {size_kb} KB")


# --------------------------------------------------------------
def main():
    """Convert a Markdown file to a styled PDF."""
    markdown, HTML, CSS = import_dependencies()
    source, target = resolve_paths(sys.argv)

    log_message(f"Reading {source}")
    with open(source, encoding='utf-8') as handle:
        text = handle.read()

    title = document_title(text, os.path.basename(source))
    body = markdown_to_html(markdown, text)
    html_text = wrap_html(body, title)

    base_url = os.path.abspath(source)
    log_message(f"Rendering with {STYLESHEET}")
    document = build_document(html_text, base_url, HTML, CSS)

    document.write_pdf(target)
    report_result(target, len(document.pages))


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
