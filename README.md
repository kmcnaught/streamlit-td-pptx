# PowerPoint to TD Snap Pageset Converter

Convert PowerPoint speaker notes into TD Snap communication pagesets.

## Features

- Extract speaker notes from PowerPoint presentations
- Split notes into button-sized chunks (4 split levels)
- Interactive preview with per-slide split controls
- Adjustable button label length (10-60 characters)
- Real-time progress feedback during processing
- Automatic color coding by slide number
- No symbol matching required (simplified for speaker notes)

## Running the App

```bash
# Install dependencies
uv sync

# Run the Streamlit app
uv run streamlit run sl_pptx_main.py
```

## How to Use

1. **Upload PowerPoint**: Select a .pptx file with speaker notes
2. **Configure Splitting**: Adjust how notes are split into buttons
   - Use split level buttons (1-4) for default behavior across all slides
   - Use ➕/➖ buttons to adjust maximum label length
   - Fine-tune individual slides with per-slide ➕/➖ controls
   - Preview shows exactly what each button will contain
3. **Upload Blank Pageset**: Select a blank TD Snap .spb file
4. **Create**: Click to generate your pageset and download
   - Real-time progress feedback shows each step

## Split Levels

- **Level 1**: Whole note (one button per slide)
- **Level 2**: Paragraphs (double line breaks) - recommended default
- **Level 3**: Lines (single line breaks)
- **Level 4**: Sentences (period + space)

## Example

If a slide titled "Introduction" has these speaker notes:
```
This is paragraph one.

This is paragraph two.
```

With split level 2 (paragraphs), you'll get:
- Button 1: "Introduction (1)" � "This is paragraph one."
- Button 2: "Introduction (2)" � "This is paragraph two."

## File Structure

- `sl_pptx_main.py` - Main Streamlit application
- `pptx_parser.py` - PowerPoint extraction and splitting logic
- `td_utils_simple.py` - TD Snap database utilities
- `colour_simple.py` - Color palette for slide-based coloring
- `static/home_button_ref.spb` - Reference template for home button

## Requirements

- Python 3.11+
- streamlit
- python-pptx
