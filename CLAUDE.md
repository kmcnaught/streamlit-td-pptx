# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Streamlit web application that converts PowerPoint speaker notes into TD Snap communication pagesets (`.spb` files). The app extracts speaker notes from slides, splits them into button-sized chunks using configurable split levels, and generates a TD Snap pageset with buttons arranged on a grid. Each button speaks its associated message when pressed.

**Key workflow:** PowerPoint (.pptx) → Extract speaker notes → Split into chunks → Generate TD Snap database (.spb) with colored buttons

## Development Commands

```bash
# Install dependencies
uv sync

# Run the Streamlit app
uv run streamlit run sl_pptx_main.py

# Run test script (processes PowerPoint without UI)
uv run python test_pptx_processing.py
```

## Architecture

### Data Flow
1. **Upload & Extract:** User uploads `.pptx` → `extract_slides()` parses slides and speaker notes
2. **Configure & Preview:** User adjusts split levels (1-4) and label format → `split_notes()` previews chunks
3. **Process:** User uploads blank `.spb` → App creates buttons in TD Snap database
4. **Database Operations:**
   - Copy home button from reference template (`static/home_button_ref.spb`)
   - Find available grid positions (10 pages, excluding reserved corners)
   - Insert buttons with slide-based colors (alternating orange shades)
   - Update timestamps and title
5. **Download:** User receives modified `.spb` file

### Core Modules

**`sl_pptx_main.py`** - Streamlit UI and orchestration
- Session state management for preview configuration
- Progress feedback during processing
- Preview with per-slide split controls
- Handles file uploads and download

**`pptx_parser.py`** - PowerPoint extraction and splitting
- `extract_slides()`: Reads `.pptx`, extracts slide titles and speaker notes
- `split_notes()`: Splits text by level (1=whole, 2=paragraphs, 3=lines, 4=sentences)
- `create_button_label()`: Generates button labels with 5 format types (title_part, slide_content, content_only, num_title, num_part_content)
- `parse_pptx_to_buttons()`: End-to-end conversion to button data tuples

**`td_utils_simple.py`** - TD Snap database utilities
- `add_buttons_from_pptx()`: Main button insertion logic - adds Button/ElementReference once, then ElementPlacement for each layout
- `find_available_positions()`: Finds empty grid cells across 10 pages
- `add_button()`, `add_element_reference()`, `add_button_placement()`: Low-level SQLite operations
- `add_home_button()`: Copies home button from reference database
- `update_timestamps()`: Updates Windows FILETIME timestamps for TD Snap compatibility
- Grid layout: Excludes bottom-right (home position) and top-right (page navigation) of each page
- **Important:** Buttons are added once to Button/ElementReference tables, but ElementPlacement entries must be duplicated for each PageLayout

**`colour_simple.py`** - Color management
- Alternates between `DARK_ORANGE` (4294951115) and `LIGHT_ORANGE` (4294934323)
- Odd slide numbers → dark orange, even → light orange
- Colors stored as 32-bit RGBA integers for TD Snap database

### TD Snap Database Structure

TD Snap `.spb` files are SQLite databases with key tables:
- **Button**: Button properties (Id, Label, Message, ImageOwnership, BorderColor, etc.) - one row per button
- **ElementReference**: Display properties (ForegroundColor, BackgroundColor, PageId) - one row per button
- **ElementPlacement**: Grid positioning (GridPosition as "col,row", PageLayoutId) - **multiple rows per button** (one for each PageLayout)
- **CommandSequence**: Button actions (serialized commands for "speak message") - one row per button
- **Page/PageLayout**: Page structure (multiple layouts per page, each with ncols×nrows grid)
- **PageSetProperties**: Pageset metadata (FriendlyName, TimeStamp)

**Critical:** A single button appears once in Button/ElementReference/CommandSequence tables, but has multiple ElementPlacement entries (one per PageLayout). This allows the same button to appear in different positions across different device layouts.

Important: Message text must be single-line (newlines converted to spaces) for TD Snap compatibility.

## File Structure

- `sl_pptx_main.py` - Main Streamlit application
- `pptx_parser.py` - PowerPoint extraction and splitting logic
- `td_utils_simple.py` - TD Snap database utilities (SQLite operations)
- `colour_simple.py` - Color palette for slide-based coloring
- `test_pptx_processing.py` - Non-interactive test script
- `static/home_button_ref.spb` - Reference template for home button
- `main.py` - Minimal entry point (not used by app)

## Split Levels

The app offers 4 levels of text splitting granularity:
1. **Level 1 (Whole)**: One button per slide (entire speaker notes)
2. **Level 2 (Paragraphs)**: Split on double line breaks (`\n\n`) - **recommended default**
3. **Level 3 (Lines)**: Split on single line breaks (`\n`)
4. **Level 4 (Sentences)**: Split on period + space/newline (`. ` or `.\n`)

Users can set a default level for all slides and override per-slide using the preview controls.

## Technical Details

### Button Label Formats
- **title_part**: "Title (Part N)" - e.g., "Introduction (2)"
- **slide_content**: "Slide N: Content..." - e.g., "Slide 3: Welcome to our..."
- **content_only**: "Content..." - e.g., "Welcome to our program..."
- **num_title**: "N - Title" - e.g., "3 - Introduction"
- **num_part_content**: "N.P: Content..." - e.g., "3.2: Welcome to our..."

### Grid Layout Logic
- 10 pages per pageset, each with configurable grid (e.g., 7×4)
- Reserved positions: Bottom-right of each page (home button), top-right of pages 2-10 (page navigation)
- `find_available_positions()` generates all valid positions and filters out occupied cells
- Buttons fill available positions sequentially across pages

### ID Management
SQLite `sqlite_sequence` table tracks next available IDs for auto-increment columns. `get_next_id()` queries this table to get the next ButtonId, ElementReferenceId, etc.

### Windows FILETIME
TD Snap uses Windows FILETIME format (100-nanosecond intervals since 1/1/1 AD). `dt_to_filetime()` converts Python datetime to this format for timestamp updates.

## Python Version

Python 3.11+ (specified in `.python-version` and `pyproject.toml`)

## Package Management

Uses `uv` for dependency management. The `pyproject.toml` includes Poetry configuration (`package-mode = false`) but primary package management is via `uv`.
