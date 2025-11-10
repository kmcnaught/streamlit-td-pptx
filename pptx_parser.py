"""
PowerPoint speaker notes extraction and splitting logic
"""

import re
from pptx import Presentation
from typing import List, Tuple


def extract_slides(pptx_file) -> List[dict]:
    """
    Extract slide information from PowerPoint file.

    Returns list of dicts with:
    - slide_num: int
    - title: str
    - notes: str (full speaker notes text)
    """
    prs = Presentation(pptx_file)
    slides = []

    for i, slide in enumerate(prs.slides, 1):
        # Get slide title
        title = "Slide " + str(i)
        if slide.shapes.title and slide.shapes.title.text.strip():
            title = slide.shapes.title.text.strip()

        # Get speaker notes
        notes = ""
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame:
                notes = notes_slide.notes_text_frame.text.strip()

        slides.append({
            'slide_num': i,
            'title': title,
            'notes': notes
        })

    return slides


def split_notes(notes: str, level: int) -> List[str]:
    """
    Split speaker notes into chunks based on split level.

    Args:
        notes: Full speaker notes text
        level: 1-4 split granularity
            1: Whole note (no split)
            2: Double line breaks (paragraphs)
            3: Single line breaks
            4: Sentences (period + space)

    Returns:
        List of text chunks
    """
    if not notes or not notes.strip():
        return []

    notes = notes.strip()

    if level == 1:
        # Whole note
        return [notes]

    elif level == 2:
        # Split by double line breaks (paragraphs)
        chunks = re.split(r'\n\s*\n', notes)
        chunks = [c.strip() for c in chunks if c.strip()]
        return chunks if chunks else [notes]

    elif level == 3:
        # Split by single line breaks
        chunks = notes.split('\n')
        chunks = [c.strip() for c in chunks if c.strip()]
        return chunks if chunks else [notes]

    elif level == 4:
        # Split by sentences (period + space/newline)
        chunks = re.split(r'\.(?:\s+|\n+)', notes)
        chunks = [c.strip() + '.' if c.strip() and not c.strip().endswith('.') else c.strip()
                  for c in chunks if c.strip()]
        return chunks if chunks else [notes]

    else:
        # Default to whole note for any other value
        return [notes]


def create_button_label(title: str, chunk_index: int, total_chunks: int, max_length: int = 30) -> str:
    """
    Create button label from slide title.

    Args:
        title: Slide title
        chunk_index: Current chunk number (0-based)
        total_chunks: Total number of chunks for this slide
        max_length: Maximum label length before adding ellipsis

    Returns:
        Label like "Slide Title... (1)" or "Introduction (2)"
    """
    # Truncate title if too long
    if len(title) > max_length:
        title = title[:max_length-3] + "..."

    # Add chunk number if multiple chunks
    if total_chunks > 1:
        label = f"{title} ({chunk_index + 1})"
    else:
        label = title

    return label


def parse_pptx_to_buttons(pptx_file, split_levels: dict = None, default_level: int = 2) -> List[Tuple[str, str, int]]:
    """
    Parse PowerPoint file into button data.

    Args:
        pptx_file: PowerPoint file object
        split_levels: Dict mapping slide_num to split level (overrides default)
        default_level: Default split level for all slides

    Returns:
        List of tuples: (label, message, slide_num)
    """
    slides = extract_slides(pptx_file)
    buttons = []

    for slide in slides:
        slide_num = slide['slide_num']
        title = slide['title']
        notes = slide['notes']

        # Skip slides with no notes
        if not notes:
            continue

        # Determine split level for this slide
        level = split_levels.get(slide_num, default_level) if split_levels else default_level

        # Split notes
        chunks = split_notes(notes, level)

        # Create buttons
        for i, chunk in enumerate(chunks):
            label = create_button_label(title, i, len(chunks))
            buttons.append((label, chunk, slide_num))

    return buttons
