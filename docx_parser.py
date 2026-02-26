"""Parse .docx worship scripts into structured sections.

Walks through every paragraph in a Word document and identifies sections by:
- Bold text starting with '+' (e.g., '+ Confession and Forgiveness')
- Known section names (Prelude, First Reading, Psalm, Message, etc.)
- Speaker prefixes: P: (Pastor), PM: (Presiding Minister), AM: (Assisting Minister), C: (Congregation)
"""

import re
from docx import Document
from docx.shared import RGBColor as DocxRGBColor


# ── Section detection patterns ──────────────────────────────────────────────

SECTION_MARKERS = [
    # (pattern, section_type, display_name)
    (r'^\+?\s*Confession\s+and\s+Forgiveness', 'confession', '+ Confession and Forgiveness'),
    (r'^\+?\s*Gathering\s+(Song|Hymn)', 'gathering_hymn', '+ Gathering Song'),
    (r'^\+?\s*Greeting', 'greeting', '+ Greeting'),
    (r'^\+?\s*Kyrie', 'kyrie', '+ Kyrie'),
    (r'^\+?\s*Canticle\s+of\s+Praise', 'canticle_of_praise', '+ Canticle of Praise'),
    (r'^\+?\s*Prayer\s+of\s+the\s+Day', 'prayer_of_day', '+ Prayer of the Day'),
    (r'^\+?\s*Children\'?s?\s+Message', 'childrens_message', "Children's Message"),
    (r'^\+?\s*First\s+Reading', 'first_reading', 'First Reading'),
    (r'^\+?\s*Psalm', 'psalm', 'Psalm'),
    (r'^\+?\s*Second\s+Reading', 'second_reading', 'Second Reading'),
    (r'^\+?\s*Gospel\s+Acclamation', 'gospel_acclamation', '+ Gospel Acclamation'),
    (r'^\+?\s*Gospel\s+Announcement', 'gospel_announcement', '+ Gospel Announcement'),
    (r'^\+?\s*Gospel\s+Reading', 'gospel_reading', '+ Gospel Reading'),
    (r'^\+?\s*Gospel', 'gospel_reading', '+ Gospel Reading'),  # fallback
    (r'^\+?\s*Message\b', 'message', 'Message'),
    (r'^\+?\s*Sermon\b', 'message', 'Message'),
    (r'^\+?\s*Hymn\s+of\s+the\s+Day', 'hymn_of_day', '+ Hymn of the Day'),
    (r'^\+?\s*Apostles\'?\s+Creed', 'creed', "+ Apostles' Creed"),
    (r'^\+?\s*Nicene\s+Creed', 'creed', '+ Nicene Creed'),
    (r'^\+?\s*Prayers?\s+of\s+(Intercession|the\s+People)', 'prayers', '+ Prayers of Intercession'),
    (r'^\+?\s*Sharing\s+of\s+the\s+Peace', 'peace', 'Sharing of the Peace'),
    (r'^\+?\s*Offering\b(?!\s+Prayer|ory)', 'offering', 'Offering'),
    (r'^\+?\s*Offertory\s+Prayer', 'offertory_prayer', '+ Offertory Prayer'),
    (r'^\+?\s*Offertory\s+(Hymn|Song)', 'offertory_hymn', '+ Offertory'),
    (r'^\+?\s*Offertory\b', 'offertory_hymn', '+ Offertory'),
    (r'^\+?\s*(The\s+)?Great\s+Thanksgiving', 'great_thanksgiving', '+ The Great Thanksgiving'),
    (r'^\+?\s*Words?\s+of\s+Institution', 'words_of_institution', '+ Words of Institution'),
    (r'^\+?\s*(The\s+)?Lord\'?s?\s+Prayer', 'lords_prayer', "+ The Lord's Prayer"),
    (r'^\+?\s*Invitation\s+to\s+(Holy\s+)?Communion', 'communion_invitation', '+ Invitation to Holy Communion'),
    (r'^\+?\s*Communion\s+(Hymn|Song)', 'communion_hymn', 'Communion Hymn'),
    (r'^\+?\s*Lamb\s+of\s+God', 'lamb_of_god', 'Lamb of God'),
    (r'^\+?\s*Post\s*[-\s]?Communion\s+Blessing', 'post_communion_blessing', '+ Post Communion Blessing'),
    (r'^\+?\s*Post\s*[-\s]?Communion\s+Canticle', 'post_communion_canticle', '+ Post Communion Canticle'),
    (r'^\+?\s*Post\s*[-\s]?Communion\s+Prayer', 'post_communion_prayer', '+ Post Communion Prayer'),
    (r'^\+?\s*Blessing\b', 'blessing', '+ Blessing'),
    (r'^\+?\s*Sending\s+(Hymn|Song)', 'sending_hymn', '+ Sending Hymn'),
    (r'^\+?\s*Dismissal', 'dismissal', '+ Dismissal'),
    (r'^\+?\s*Postlude', 'postlude', '+Postlude'),
    (r'^\+?\s*Announcements?\b', 'announcements', 'Announcements'),
    (r'^\+?\s*Prelude\b', 'prelude', 'Prelude'),
    (r'^\+?\s*Holy,?\s*Holy,?\s*Holy', 'holy_holy_holy', 'Holy, Holy, Holy'),
]

HYMN_SECTIONS = {
    'gathering_hymn', 'hymn_of_day', 'sending_hymn',
    'offertory_hymn', 'post_communion_canticle', 'communion_hymn',
}

READING_SECTIONS = {
    'first_reading', 'second_reading', 'psalm', 'gospel_reading',
}


def _get_paragraph_text(para):
    """Get full text of a paragraph."""
    return para.text.strip()


def _is_bold_paragraph(para):
    """Check if the paragraph is predominantly bold."""
    if not para.runs:
        return False
    bold_chars = sum(len(r.text) for r in para.runs if r.bold)
    total_chars = sum(len(r.text) for r in para.runs)
    if total_chars == 0:
        return False
    return bold_chars / total_chars > 0.5


def _has_cross_marker(para):
    """Check if paragraph starts with '+' or '☩' (section marker)."""
    text = para.text.strip()
    return text.startswith('+') or text.startswith('☩')


def _detect_section(text):
    """Try to match text against known section markers.

    Returns (section_type, display_name) or (None, None).
    """
    clean = text.strip().lstrip('+').strip()
    full = text.strip()

    for pattern, sec_type, display in SECTION_MARKERS:
        if re.match(pattern, full, re.IGNORECASE) or re.match(pattern, clean, re.IGNORECASE):
            return sec_type, display
    return None, None


def _is_stage_direction(para):
    """Check if a paragraph is a stage direction (italic, often in red)."""
    if not para.runs:
        return False
    # Check if all runs are italic
    all_italic = all(r.italic for r in para.runs if r.text.strip())
    if not all_italic:
        return False
    # Check for red color
    for r in para.runs:
        if r.font.color and r.font.color.rgb:
            try:
                if r.font.color.rgb == DocxRGBColor(0xFF, 0x00, 0x00):
                    return True
            except:
                pass
    # Some stage directions are just italic without red
    text = para.text.strip().lower()
    stage_keywords = ['silence', 'stand', 'sit', 'kneel', 'congregation',
                      'assembly', 'please', 'all may']
    return any(kw in text for kw in stage_keywords)


def _extract_runs_with_formatting(para):
    """Extract runs preserving bold/italic/color information."""
    runs = []
    for run in para.runs:
        color = None
        if run.font.color and run.font.color.rgb:
            try:
                color = str(run.font.color.rgb)
            except:
                pass
        runs.append({
            'text': run.text,
            'bold': run.bold,
            'italic': run.italic,
            'color': color,
        })
    return runs


def parse_worship_script(docx_path):
    """Parse a .docx worship script into structured sections.

    Returns a list of section dicts:
    {
        'type': str,            # Section type identifier
        'title': str,           # Display title for the section
        'content': list[str],   # Lines of content text
        'raw_paragraphs': list, # Raw paragraph data with formatting
        'metadata': dict,       # Additional info (hymn numbers, page refs, etc.)
    }
    """
    doc = Document(docx_path)
    sections = []
    current_section = None
    title_block = None

    # First pass: detect if the document starts with a title block
    # (service name, date, etc. before the first section marker)
    paragraphs = list(doc.paragraphs)

    i = 0
    # Look for title block at the beginning
    title_lines = []
    while i < len(paragraphs):
        text = _get_paragraph_text(paragraphs[i])
        if not text:
            i += 1
            continue

        sec_type, sec_display = _detect_section(text)
        if sec_type:
            break  # Found first real section

        # Check if it's a section-like bold header
        if _has_cross_marker(paragraphs[i]) and _is_bold_paragraph(paragraphs[i]):
            break

        title_lines.append(text)
        i += 1

    if title_lines:
        title_block = _parse_title_block(title_lines)
        sections.append({
            'type': 'title',
            'title': 'Title Slide',
            'content': title_lines,
            'raw_paragraphs': [],
            'metadata': title_block,
        })

    # Main parsing loop
    while i < len(paragraphs):
        para = paragraphs[i]
        text = _get_paragraph_text(para)

        if not text:
            # Blank line - could be verse separator in hymns
            if current_section and current_section['type'] in HYMN_SECTIONS:
                current_section['content'].append('')  # Preserve blank line as verse separator
            i += 1
            continue

        # Check if this starts a new section
        # Only treat as a section header if the paragraph is bold OR starts with '+'
        # This prevents content lines like "Psalm 32:1-11" or "Lamb of God" from
        # being mistakenly detected as new section headers.
        sec_type, sec_display = None, None
        is_header_candidate = (_is_bold_paragraph(para) or _has_cross_marker(para))

        if is_header_candidate:
            sec_type, sec_display = _detect_section(text)

            # Also detect bold lines with + as section headers
            if not sec_type and _has_cross_marker(para):
                sec_type = 'generic_section'
                sec_display = text.strip()

        if sec_type:
            # Start a new section
            current_section = {
                'type': sec_type,
                'title': sec_display,
                'content': [],
                'raw_paragraphs': [],
                'metadata': {},
            }
            sections.append(current_section)

            # Check if the same line has more content after the section name
            # e.g., "+ Gathering Song  O Lord, throughout These Forty Days #319"
            remaining = _extract_remaining_after_section(text, sec_type)
            if remaining:
                current_section['content'].append(remaining)
                current_section['raw_paragraphs'].append(
                    _extract_runs_with_formatting(para))

            i += 1
            continue

        # Not a new section - add to current section
        if current_section is not None:
            current_section['content'].append(text)
            current_section['raw_paragraphs'].append(
                _extract_runs_with_formatting(para))

            # Check for stage directions
            if _is_stage_direction(para):
                if 'stage_directions' not in current_section['metadata']:
                    current_section['metadata']['stage_directions'] = []
                current_section['metadata']['stage_directions'].append(text)
        else:
            # Content before any section - add to title block or create misc
            if sections and sections[0]['type'] == 'title':
                sections[0]['content'].append(text)
            else:
                current_section = {
                    'type': 'misc',
                    'title': '',
                    'content': [text],
                    'raw_paragraphs': [_extract_runs_with_formatting(para)],
                    'metadata': {},
                }
                sections.append(current_section)

        i += 1

    # Post-processing: extract metadata from section contents
    for section in sections:
        _enrich_section_metadata(section)

    return sections


def _parse_title_block(lines):
    """Parse the title block lines into metadata.

    Looks for: service name, date, Sunday/season name, pastor, music director.
    """
    metadata = {
        'service_name': '',
        'date': '',
        'sunday_name': '',
        'pastor': '',
        'music_director': '',
        'tagline': '',
    }

    for line in lines:
        lower = line.lower()
        if 'worship' in lower or 'communion' in lower or 'service' in lower:
            metadata['service_name'] = line
        elif 'pastor' in lower or 'dr.' in lower or 'rev.' in lower:
            metadata['pastor'] = line
        elif 'music' in lower or 'director' in lower or 'organist' in lower:
            metadata['music_director'] = line
        elif re.search(r'(sunday|lent|advent|easter|epiphany|pentecost|christmas|ordinary)', lower):
            metadata['sunday_name'] = line
        elif re.search(r'\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b', line):
            metadata['date'] = line
        elif 'seek god' in lower or 'share life' in lower:
            metadata['tagline'] = line
        elif not metadata['sunday_name'] and len(line) < 50:
            # Short unidentified line - likely Sunday name or subtitle
            metadata['sunday_name'] = line

    return metadata


def _extract_remaining_after_section(text, sec_type):
    """Extract any content on the same line after the section name.

    For example: "+ Gathering Song  O Lord, throughout These Forty Days"
    Returns the remaining text or None.
    """
    # For most sections, we just check if there's a multi-line section header
    # with content embedded. The docx typically has these as separate paragraphs.
    return None


def _enrich_section_metadata(section):
    """Extract metadata like hymn numbers, page refs, scripture refs from content."""
    sec_type = section['type']
    content = section['content']

    if sec_type in HYMN_SECTIONS:
        _extract_hymn_metadata(section)
    elif sec_type in READING_SECTIONS:
        _extract_reading_metadata(section)
    elif sec_type in ('peace', 'dismissal', 'creed', 'lords_prayer',
                       'great_thanksgiving', 'offering', 'blessing'):
        _extract_page_ref(section)


def _extract_hymn_metadata(section):
    """Extract hymn title, number, and verses from a hymn section."""
    content = section['content']
    if not content:
        return

    # Look for hymn number pattern (#NNN)
    hymn_title = None
    hymn_number = None
    verse_start = 0

    for idx, line in enumerate(content):
        match = re.search(r'#(\d+)', line)
        if match:
            hymn_number = '#' + match.group(1)
            # Title is this line without the number
            title_part = line[:match.start()].strip().rstrip('-').strip()
            if title_part:
                hymn_title = title_part
            elif idx > 0:
                hymn_title = content[0]
            verse_start = idx + 1
            break
        elif idx == 0 and not re.match(r'^[A-Z][a-z]', line):
            continue
        elif idx == 0:
            hymn_title = line
            verse_start = 1

    if hymn_title is None and content:
        hymn_title = content[0]
        verse_start = 1

    # Parse verses (separated by blank lines)
    verses = []
    current_verse_lines = []
    verse_num = 1

    for line in content[verse_start:]:
        if line.strip() == '':
            if current_verse_lines:
                verses.append((verse_num, '\n'.join(current_verse_lines)))
                verse_num += 1
                current_verse_lines = []
        else:
            current_verse_lines.append(line)

    if current_verse_lines:
        verses.append((verse_num, '\n'.join(current_verse_lines)))

    section['metadata']['hymn_title'] = hymn_title
    section['metadata']['hymn_number'] = hymn_number
    section['metadata']['verses'] = verses


def _extract_reading_metadata(section):
    """Extract scripture reference and pew bible page from a reading section."""
    content = section['content']
    if not content:
        return

    # First content line often has the reference
    first_line = content[0] if content else ''

    # Look for scripture reference pattern (Book Chapter:Verse)
    ref_match = re.match(
        r'^((?:\d\s+)?[A-Z][a-z]+(?:\s+[a-z]+)*)\s+(\d+[:\d,\-.\s]+\d*)',
        first_line
    )
    if ref_match:
        section['metadata']['reference'] = ref_match.group(0).strip()
    else:
        section['metadata']['reference'] = first_line.split('Pew')[0].strip()

    # Look for pew bible reference
    pew_match = re.search(r'Pew\s+Bibles?\s+[A-Z.]+\s+p\.?\s*\d+', first_line)
    if pew_match:
        section['metadata']['pew_bible'] = pew_match.group(0)

    # Reading text starts after the reference line
    reading_lines = content[1:] if len(content) > 1 else []
    section['metadata']['reading_text'] = '\n'.join(reading_lines)


def _extract_page_ref(section):
    """Extract page reference (p. NNN) from section content."""
    for line in section['content']:
        match = re.search(r'p\.?\s*\d+', line)
        if match:
            section['metadata']['page_ref'] = match.group(0)
            return


# ── Standalone test ─────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys
    import json

    if len(sys.argv) < 2:
        print("Usage: python docx_parser.py <path_to_docx>")
        sys.exit(1)

    sections = parse_worship_script(sys.argv[1])

    for s in sections:
        print(f"\n{'='*60}")
        print(f"Section: {s['type']} - {s['title']}")
        print(f"Metadata: {s['metadata']}")
        print(f"Content lines: {len(s['content'])}")
        for line in s['content'][:5]:
            print(f"  {line[:80]}")
        if len(s['content']) > 5:
            print(f"  ... ({len(s['content']) - 5} more lines)")
