"""Text splitting and formatting helpers for slide generation."""

import re
from config import (
    MAX_CHARS_BODY, MAX_CHARS_READING,
    MAX_LINES_BODY, MAX_LINES_READING,
    SPEAKER_PREFIXES, CONGREGATION_PREFIXES
)


def is_speaker_line(line):
    """Check if a line starts with a speaker prefix (P:, PM:, AM:, C:, L:)."""
    stripped = line.strip()
    for prefix in SPEAKER_PREFIXES:
        if stripped.startswith(prefix):
            return True
    return False


def is_congregation_line(line):
    """Check if a line starts with C: (congregation)."""
    stripped = line.strip()
    for prefix in CONGREGATION_PREFIXES:
        if stripped.startswith(prefix):
            return True
    return False


def get_speaker_prefix(line):
    """Extract the speaker prefix from a line, or None."""
    stripped = line.strip()
    for prefix in SPEAKER_PREFIXES:
        if stripped.startswith(prefix):
            return prefix
    return None


def strip_speaker_prefix(line):
    """Remove the speaker prefix from a line."""
    stripped = line.strip()
    for prefix in SPEAKER_PREFIXES:
        if stripped.startswith(prefix):
            return stripped[len(prefix):].strip()
    return stripped


def split_text_into_slides(text, max_chars=MAX_CHARS_BODY, max_lines=MAX_LINES_BODY):
    """Split a block of text into slide-sized chunks.

    Tries to split at natural boundaries:
    1. Speaker changes (P:/C: transitions)
    2. Blank lines / paragraph breaks
    3. Sentence boundaries
    """
    if not text or not text.strip():
        return []

    lines = text.strip().split('\n')

    # If it fits on one slide, return as-is
    if len(lines) <= max_lines and len(text) <= max_chars:
        return [text.strip()]

    chunks = []
    current_lines = []
    current_chars = 0

    for line in lines:
        line_len = len(line) + 1  # +1 for newline
        would_exceed_chars = (current_chars + line_len) > max_chars
        would_exceed_lines = len(current_lines) >= max_lines

        # Check for speaker change as a natural split point
        is_new_speaker = is_speaker_line(line) and current_lines

        if (would_exceed_chars or would_exceed_lines) and current_lines:
            chunks.append('\n'.join(current_lines))
            current_lines = []
            current_chars = 0
        elif is_new_speaker and current_chars > max_chars * 0.5:
            # Split at speaker change if we're past halfway
            chunks.append('\n'.join(current_lines))
            current_lines = []
            current_chars = 0

        current_lines.append(line)
        current_chars += line_len

    if current_lines:
        chunks.append('\n'.join(current_lines))

    return chunks


def split_dialogue_into_slides(lines, max_chars=MAX_CHARS_BODY, max_lines=MAX_LINES_BODY):
    """Split dialogue (P:/C: exchanges) into slide-sized chunks.

    Keeps P:/C: exchanges together when possible, splitting at speaker changes.
    """
    if not lines:
        return []

    chunks = []
    current_chunk = []
    current_chars = 0

    for line in lines:
        line_len = len(line) + 1
        would_overflow = (current_chars + line_len > max_chars or
                          len(current_chunk) >= max_lines)

        if would_overflow and current_chunk:
            chunks.append(current_chunk)
            current_chunk = []
            current_chars = 0

        current_chunk.append(line)
        current_chars += line_len

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


def split_reading_text(text, max_chars=MAX_CHARS_READING, max_lines=MAX_LINES_READING):
    """Split scripture reading text into slide-sized chunks.

    Tries to split at verse boundaries (lines starting with verse numbers).
    """
    if not text.strip():
        return []

    lines = text.strip().split('\n')

    if len(lines) <= max_lines and len(text) <= max_chars:
        return [text.strip()]

    chunks = []
    current_lines = []
    current_chars = 0

    for line in lines:
        line_len = len(line) + 1
        would_overflow = (current_chars + line_len > max_chars or
                          len(current_lines) >= max_lines)

        if would_overflow and current_lines:
            chunks.append('\n'.join(current_lines))
            current_lines = []
            current_chars = 0

        current_lines.append(line)
        current_chars += line_len

    if current_lines:
        chunks.append('\n'.join(current_lines))

    return chunks


def split_hymn_into_verses(text):
    """Split hymn text into individual verses.

    Verses are typically separated by blank lines.
    Returns list of (verse_number, verse_text) tuples.
    """
    if not text.strip():
        return []

    # Split on double newlines or verse number patterns
    blocks = re.split(r'\n\s*\n', text.strip())
    verses = []

    for i, block in enumerate(blocks):
        block = block.strip()
        if not block:
            continue

        # Try to extract verse number from block
        # Patterns: "1.", "1", "Verse 1", standalone number at end
        verse_num = i + 1

        # Check if last line is just a number (verse number indicator)
        block_lines = block.split('\n')
        last_line = block_lines[-1].strip()
        if last_line.isdigit():
            verse_num = int(last_line)
            block = '\n'.join(block_lines[:-1]).strip()

        verses.append((verse_num, block))

    return verses


def extract_hymn_info(text):
    """Extract hymn title and number from header text.

    Looks for patterns like:
    - "A Mighty Fortress #504"
    - "O Lord, throughout These Forty Days #319"

    Returns (title, hymn_number) tuple.
    """
    # Look for #NNN pattern
    match = re.search(r'#(\d+)', text)
    hymn_number = match.group(0) if match else None

    # Title is everything before the hymn number
    if match:
        title = text[:match.start()].strip()
        # Clean up title - remove trailing newlines and dashes
        title = title.rstrip('- \n')
    else:
        title = text.strip()

    return title, hymn_number


def extract_page_reference(text):
    """Extract page reference like 'p. 105' or 'p.147' from text."""
    match = re.search(r'p\.?\s*\d+', text)
    return match.group(0) if match else None


def extract_scripture_reference(text):
    """Extract scripture reference and pew bible page from text.

    Example: "Genesis 2:15-17,3:1-7    Pew Bibles O.T. p.2"
    Returns (reference, pew_bible_page) tuple.
    """
    # Look for "Pew Bibles" pattern
    pew_match = re.search(r'Pew Bibles?\s+[A-Z.]+\s+p\.?\s*\d+', text)
    pew_bible = pew_match.group(0) if pew_match else None

    # Reference is everything before the pew bible ref (or the whole text)
    if pew_match:
        reference = text[:pew_match.start()].strip()
    else:
        reference = text.strip()

    return reference, pew_bible


def clean_text(text):
    """Clean up text by normalizing whitespace and removing artifacts."""
    if not text:
        return ''
    # Normalize tabs to spaces
    text = text.replace('\t', ' ')
    # Collapse multiple spaces (but not newlines)
    text = re.sub(r' +', ' ', text)
    # Strip each line
    lines = [line.strip() for line in text.split('\n')]
    return '\n'.join(lines)
