"""Constants for St. Luke worship script to PowerPoint converter."""

from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor

# ── Colors ──────────────────────────────────────────────────────────────────
GOLD = RGBColor(0xDB, 0xB2, 0x41)          # C: prefix color
RED = RGBColor(0xFF, 0x00, 0x00)            # Stage directions
WHITE = RGBColor(0xFF, 0xFF, 0xFF)          # Default text
HYMN_NUM_GOLD = RGBColor(0xFF, 0xC0, 0x00)  # Hymn number highlight
PEACH = RGBColor(0xFA, 0xBD, 0xA7)          # Title slide pastor/music

# ── Font sizes (EMU values used by python-pptx) ────────────────────────────
TITLE_SIZE = Pt(60)           # Section header titles (~762000 EMU)
SUBTITLE_SIZE = Pt(50)        # Sub-titles (~635000 EMU)
BODY_SIZE = Pt(38)            # Standard body text (~482600 EMU)
BODY_LARGE = Pt(40)           # Slightly larger body (~508000 EMU)
BODY_SMALL = Pt(35)           # Smaller body text (~444500 EMU)
READING_SIZE = Pt(38)         # Scripture reading text
HYMN_LYRIC_SIZE = Pt(44)      # Hymn verse lyrics (~558800 EMU)
VERSE_NUM_SIZE = Pt(50)       # Verse number in text box (~635000 EMU)
PAGE_REF_SIZE = Pt(45)        # Page references (~571500 EMU)
SMALL_TEXT = Pt(30)           # Smaller notes
LICENSE_SIZE = Pt(18)         # License text

# ── Layout name → index mapping ────────────────────────────────────────────
# These map layout names (as they appear in the template) to their index
LAYOUTS = {
    'Title Slide': 0,
    'Title and Content': 1,
    '5_Section Header': 2,
    '1_Section Header': 3,
    '4_Section Header': 4,
    '1_Title and Content': 5,
    '6_Section Header': 8,
    '5_Title and Content': 12,
    '14_Section Header': 31,
    '23_Title and Content': 42,
    '23_Section Header': 54,
    '24_Section Header': 11,
}

# ── Slide dimension constants ──────────────────────────────────────────────
SLIDE_WIDTH = Emu(9144000)    # 10 inches
SLIDE_HEIGHT = Emu(6858000)   # 7.5 inches

# ── Template and file paths ────────────────────────────────────────────────
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'template', 'template.pptx')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# ── Text splitting thresholds (approximate chars per slide) ─────────────────
MAX_CHARS_BODY = 450          # Max characters for body placeholder
MAX_CHARS_READING = 500       # Max characters for reading slides
MAX_LINES_BODY = 8            # Max lines for body placeholder
MAX_LINES_READING = 10        # Max lines for reading slides

# ── Speaker prefixes ───────────────────────────────────────────────────────
SPEAKER_PREFIXES = ['P:', 'PM:', 'AM:', 'C:', 'L:']
CONGREGATION_PREFIXES = ['C:']
MINISTER_PREFIXES = ['P:', 'PM:', 'AM:', 'L:']
