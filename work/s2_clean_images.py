#!/usr/bin/env python3
"""
Normalize downloaded pictures for use on slides.

Reads every picture in "images_raw" and writes a slide-ready
copy into "images". The originals are never modified, so
this script needs no backup of its own: "images" can be
deleted and rebuilt at any time, and "images_raw" is the
permanent record of what the network returned.

Keeping the deck small is this script's job. A PPTX is
almost entirely its pictures, so shrinking them here is
what PowerPoint's "Compress Pictures" used to do by hand,
except it happens once, automatically, and the deck never
grows large in the first place.

Each picture is scaled to fit the largest box the layout
can ever give it, at a chosen resolution:

    box    5.10 x 5.07 inches, from deck_layout
    at 150 ppi that is 765 x 760 pixels

Storing more pixels than that is wasted bytes: the slide
cannot show them. Aspect ratio is preserved and pictures
are never enlarged, so a small source stays small.

    --ppi 96     PowerPoint's "email", smallest files
    --ppi 150    the default, good on a 1080p projector
    --ppi 220    PowerPoint's "print", visibly sharper

The standard applied to each file:

    JPEG, sRGB, quality 85
    white background, alpha flattened
    stray border whitespace trimmed (2% fuzz)
    scaled to fit the slide's picture box
    aspect ratio preserved, never padded

Aspect ratio matters. An earlier version padded everything
onto a fixed canvas, which is right for Markdown because
Markdown cannot resize a picture. On a slide it is wrong:
the white bars get baked into the file and then the
renderer adds its own. s3_make_pptx.py fits pictures into
their box at placement time.

Files already carrying the marker for the current settings
are skipped, so re-runs are cheap. Changing --ppi or
--quality changes the marker, so everything is rebuilt.

Usage:
    python3 s2_clean_images.py
    python3 s2_clean_images.py --ppi 96
    python3 s2_clean_images.py --force
    python3 s2_clean_images.py images_raw/gpt6-astra.png

Created: 2026-09-04
Last updated: 2026-09-12
"""

import glob
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime

import deck_layout as L
from deck_parser import DeckError, parse_deck_file

DECK_GLOB = "*-AI-News.md"

RAW_DIR = "images_raw"
OUT_DIR = "images"

SUPPORTED_EXTENSIONS = [
    '.avif', '.heic', '.jpeg', '.jpg',
    '.png', '.webp', '.gif', '.bmp', '.tif', '.tiff'
]

DEFAULT_PPI = 150
DEFAULT_QUALITY = 85
BG_COLOR = "white"
TRIM_FUZZ = "2%"

CONVERT_LOG = "/tmp/convert.log"
MAGICK = shutil.which('magick') or shutil.which('convert')
IDENTIFY = shutil.which('identify')


@dataclass
class Settings:
    """How pictures should be normalized."""

    ppi: int = DEFAULT_PPI
    quality: int = DEFAULT_QUALITY

    # --------------------------------------
    def box(self, width_in=None):
        """Pixel box one picture must fit inside."""
        wide = width_in if width_in else L.MAX_IMG_W
        return (int(wide * self.ppi),
                int(L.BAND_H * self.ppi))

    # --------------------------------------
    def geometry(self, width_in=None):
        """ImageMagick resize argument for that box."""
        width, height = self.box(width_in)
        return f"{width}x{height}>"

    # --------------------------------------
    def stamp(self, width_in=None):
        """Marker recording the settings used.

        The box is part of it, so moving a picture to a
        wider slot rebuilds it without --force.
        """
        width, height = self.box(width_in)
        return (
            f"seminar-{width}x{height}-q{self.quality}"
        )


# --------------------------------------------------------------
def log_message(message):
    """Print timestamped log message."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {message}")


# --------------------------------------------------------------
def check_imagemagick():
    """Check that ImageMagick tools are available."""
    if MAGICK and IDENTIFY:
        log_message(f"ImageMagick found: {MAGICK}")
        return True
    log_message(
        "ERROR: ImageMagick not found. "
        "Install: brew install imagemagick"
    )
    return False


# --------------------------------------------------------------
def find_images_for_ext(directory, ext):
    """Find images for a single extension."""
    files = []
    for pattern in [f"*{ext}", f"*{ext.upper()}"]:
        files.extend(glob.glob(
            os.path.join(directory, pattern)
        ))
    return files


# --------------------------------------------------------------
def find_raw_files(directory):
    """Find every source picture in the raw directory."""
    found = []
    for ext in SUPPORTED_EXTENSIONS:
        found.extend(find_images_for_ext(directory, ext))
    return sorted(set(found))


# --------------------------------------------------------------
def output_path(raw_path):
    """Build the cleaned path for one raw picture."""
    stem = os.path.splitext(os.path.basename(raw_path))[0]
    return os.path.join(OUT_DIR, f"{stem}.jpg")


# --------------------------------------------------------------
def read_comment(path):
    """Read the marker comment of an image, or ''."""
    try:
        result = subprocess.run(
            [IDENTIFY, '-format', '%c', path],
            capture_output=True, text=True, check=True
        )
        return result.stdout.strip()
    except (subprocess.CalledProcessError, OSError):
        return ""


# --------------------------------------------------------------
def needs_processing(raw_path, out_path, force,
                     settings, wide):
    """Decide whether one picture must be rebuilt."""
    if force or not os.path.isfile(out_path):
        return True
    if read_comment(out_path) != settings.stamp(wide):
        log_message(
            f"Rebuilding {out_path} - different settings"
        )
        return True
    if os.path.getmtime(out_path) < os.path.getmtime(raw_path):
        log_message(f"Rebuilding {out_path} - source newer")
        return True
    return False


# --------------------------------------------------------------
def build_convert_cmd(src, dst, settings, wide):
    """Build the ImageMagick command for one picture."""
    return [
        MAGICK, f"{src}[0]",
        '-auto-orient',
        '-background', BG_COLOR,
        '-alpha', 'remove', '-alpha', 'off',
        '-colorspace', 'sRGB',
        '-fuzz', TRIM_FUZZ, '-trim', '+repage',
        '-resize', settings.geometry(wide),
        '-strip',
        '-set', 'comment', settings.stamp(wide),
        '-density', str(settings.ppi),
        '-units', 'PixelsPerInch',
        '-quality', str(settings.quality),
        '-interlace', 'none',
        '-sampling-factor', '4:2:0',
        dst
    ]


# --------------------------------------------------------------
def run_convert(src, dst, settings, wide):
    """Run ImageMagick, returning True on success."""
    try:
        with open(CONVERT_LOG, 'a') as log_file:
            subprocess.run(
                build_convert_cmd(src, dst, settings, wide),
                stdout=log_file, stderr=log_file, check=True
            )
        return True
    except (subprocess.CalledProcessError, OSError) as exc:
        log_message(f"ERROR: convert failed for {src}: {exc}")
        return False


# --------------------------------------------------------------
def normalize_image(raw_path, out_path, settings, wide):
    """Write one normalized copy, atomically."""
    temp_path = f"{out_path}.tmp.jpg"
    log_message(
        f"Cleaning {os.path.basename(raw_path)} -> "
        f"{os.path.basename(out_path)}"
    )
    if not run_convert(raw_path, temp_path, settings, wide):
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False
    os.replace(temp_path, out_path)
    return True


# --------------------------------------------------------------
def process_all(raw_files, force, settings, widths):
    """Normalize every raw picture, counting outcomes."""
    done = skipped = errors = 0
    for raw_path in raw_files:
        out_path = output_path(raw_path)
        stem = os.path.splitext(
            os.path.basename(raw_path)
        )[0]
        wide = widths.get(stem, L.MAX_IMG_W)
        if not needs_processing(
            raw_path, out_path, force, settings, wide
        ):
            skipped += 1
            continue
        if normalize_image(
            raw_path, out_path, settings, wide
        ):
            done += 1
        else:
            errors += 1
    return done, skipped, errors


# --------------------------------------------------------------
def report_collisions(raw_files):
    """Warn when two raw files map to the same output."""
    seen = {}
    for raw_path in raw_files:
        seen.setdefault(output_path(raw_path), []).append(
            raw_path
        )
    clashes = 0
    for out_path, sources in seen.items():
        if len(sources) > 1:
            clashes += 1
            log_message(
                f"WARNING: {len(sources)} raw files map to "
                f"{out_path}: {', '.join(sources)}"
            )
    return clashes


# --------------------------------------------------------------
def report_orphans(raw_files):
    """List cleaned pictures with no source any more."""
    expected = {output_path(p) for p in raw_files}
    orphans = [
        p for p in find_images_for_ext(OUT_DIR, '.jpg')
        if p not in expected
    ]
    for path in sorted(orphans):
        log_message(f"Orphan (no raw source): {path}")
    if orphans:
        log_message(
            "Orphans are reported, never deleted. Remove "
            "them by hand once you are sure."
        )
    return len(orphans)


# --------------------------------------------------------------
def directory_size(directory, pattern="*"):
    """Total bytes of the files matching a pattern."""
    return sum(
        os.path.getsize(p)
        for p in glob.glob(os.path.join(directory, pattern))
        if os.path.isfile(p)
    )


# --------------------------------------------------------------
def init_conversion_log():
    """Initialize the ImageMagick output log file."""
    with open(CONVERT_LOG, 'w') as handle:
        handle.write(f"Conversion log - {datetime.now()}\n")
        handle.write("=" * 50 + "\n")


# --------------------------------------------------------------
def print_summary(counts, orphans, settings):
    """Print the processing summary."""
    done, skipped, errors = counts
    width, height = settings.box()
    raw_kb = directory_size(RAW_DIR) / 1024
    out_kb = directory_size(OUT_DIR, "*.jpg") / 1024

    log_message("=" * 50)
    log_message("Image cleaning completed")
    log_message(
        f"Standard: JPEG q{settings.quality}, at most "
        f"{width}x{height} px at {settings.ppi} ppi"
    )
    log_message(f"Cleaned: {done}")
    log_message(f"Skipped (already current): {skipped}")
    log_message(f"Errors: {errors}")
    log_message(f"Orphans in {OUT_DIR}: {orphans}")
    log_message(
        f"Size: {raw_kb:.0f} KB raw -> {out_kb:.0f} KB "
        f"for the slides"
    )
    if errors:
        log_message(
            f"WARNING: {errors} files failed. "
            f"Check {CONVERT_LOG}."
        )


# --------------------------------------------------------------
def read_option(argv, name, fallback):
    """Read an integer --option value from the command."""
    if name not in argv:
        return fallback
    index = argv.index(name)
    if index + 1 >= len(argv):
        log_message(f"ERROR: {name} needs a number")
        sys.exit(1)
    try:
        return int(argv[index + 1])
    except ValueError:
        log_message(f"ERROR: {name} needs a number")
        sys.exit(1)


# --------------------------------------------------------------
def read_settings(argv):
    """Build the Settings from the command line."""
    settings = Settings(
        ppi=read_option(argv, "--ppi", DEFAULT_PPI),
        quality=read_option(
            argv, "--quality", DEFAULT_QUALITY
        ),
    )
    if settings.ppi < 48 or settings.ppi > 600:
        log_message("ERROR: --ppi must be between 48 and 600")
        sys.exit(1)
    return settings


# --------------------------------------------------------------
def collect_targets(argv):
    """Return the raw picture paths to process."""
    skip = {"--ppi", "--quality"}
    named = []
    for index, value in enumerate(argv[1:], start=1):
        if value.startswith('-') or argv[index - 1] in skip:
            continue
        named.append(value)

    if named:
        return [p for p in named if os.path.isfile(p)]
    if not os.path.isdir(RAW_DIR):
        log_message(
            f"ERROR: {RAW_DIR} not found. "
            f"Run s1_fetch_images.py first."
        )
        sys.exit(1)
    return find_raw_files(RAW_DIR)


# --------------------------------------------------------------
def read_widths(argv):
    """Learn each picture's widest box from the deck.

    Falls back to the widest box any layout uses, which is
    safe but stores more pixels than most slots can show.
    """
    named = [
        a for a in argv[1:]
        if a.endswith(".md") and os.path.isfile(a)
    ]
    source = named[0] if named else None
    if source is None:
        found = sorted(glob.glob(DECK_GLOB))
        source = found[-1] if found else None
    if source is None:
        return {}
    try:
        deck = parse_deck_file(source)
    except DeckError as exc:
        log_message(f"WARNING: cannot read {source}: {exc}")
        return {}
    log_message(f"Sizing pictures against {source}")
    return L.image_widths(deck.sections)


# --------------------------------------------------------------
def main():
    """Normalize every picture from raw into images."""
    log_message("Starting image cleaning")

    if not check_imagemagick():
        sys.exit(1)

    settings = read_settings(sys.argv)
    force = '--force' in sys.argv
    raw_files = collect_targets(sys.argv)
    if not raw_files:
        log_message(f"No pictures found in {RAW_DIR}")
        return

    os.makedirs(OUT_DIR, exist_ok=True)
    init_conversion_log()
    report_collisions(raw_files)

    widths = read_widths(sys.argv)
    counts = process_all(raw_files, force, settings, widths)
    orphans = report_orphans(find_raw_files(RAW_DIR))
    print_summary(counts, orphans, settings)


# --------------------------------------------------------------
if __name__ == "__main__":
    main()
