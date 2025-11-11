#!/usr/bin/env python3
"""
Test script that replicates the exact processing flow from sl_pptx_main.py
without the Streamlit UI.

Uses the same code paths and functions as the main app.
"""

import os
from pptx_parser import parse_pptx_to_buttons
from td_utils_simple import (
    create_temp_file,
    add_home_button,
    get_static_path,
    update_timestamps,
    update_page_title,
    add_buttons_from_pptx,
)

# Input files
PPTX_FILE = "Alternative access to block-based coding tools-v2.pptx"
BLANK_SPB = "working-examples/empty page.spb"
OUTPUT_FILE = "test-output.spb"

def main():
    print("="*80)
    print("Testing PPTX to TD Snap conversion")
    print("="*80)

    # Step 1: Parse PowerPoint
    print("\n📄 Parsing PowerPoint file...")
    with open(PPTX_FILE, 'rb') as f:
        pptx_data = f.read()

    # Create a file-like object (similar to st.file_uploader)
    from io import BytesIO
    pptx_file = BytesIO(pptx_data)

    buttons_data = parse_pptx_to_buttons(
        pptx_file,
        split_levels={},  # Empty dict - no per-slide overrides
        default_level=2   # Default split level
    )

    print(f"✓ Found {len(buttons_data)} buttons to create")
    if buttons_data:
        print(f"  First button: {buttons_data[0]}")

    # Step 2: Create temp copy of blank SPB
    print("\n📋 Creating temporary pageset copy...")
    with open(BLANK_SPB, 'rb') as f:
        spb_data = f.read()

    spb_file = BytesIO(spb_data)
    temp_db_path = create_temp_file(spb_file)
    print(f"✓ Created temp file: {temp_db_path}")

    # Step 3: Add home button
    print("\n🏠 Adding home button...")
    reference_db = get_static_path('home_button_ref.spb')
    add_home_button(temp_db_path, reference_db)
    print("✓ Added home button")

    # Step 4: Add buttons from PowerPoint
    print(f"\n➕ Adding {len(buttons_data)} buttons to pageset...")
    num_added = add_buttons_from_pptx(temp_db_path, buttons_data)
    print(f"✅ Successfully added {num_added} buttons")

    # Step 5: Update title
    print("\n📝 Updating pageset title...")
    pageset_title = "Test Pageset"
    update_page_title(temp_db_path, pageset_title)
    print(f"✓ Updated title to: {pageset_title}")

    # Step 6: Update timestamps
    print("\n🕒 Updating timestamps...")
    update_timestamps(temp_db_path)
    print("✓ Updated timestamps")

    # Step 7: Copy to output file
    print(f"\n💾 Saving to {OUTPUT_FILE}...")
    with open(temp_db_path, 'rb') as f:
        output_data = f.read()

    with open(OUTPUT_FILE, 'wb') as f:
        f.write(output_data)

    print(f"✅ Saved to {OUTPUT_FILE}")

    # Clean up temp file
    try:
        os.remove(temp_db_path)
        print(f"🧹 Cleaned up temp file")
    except:
        pass

    print("\n" + "="*80)
    print(f"DONE! Output file: {OUTPUT_FILE}")
    print("="*80)
    print(f"\nCreated pageset with {num_added} buttons")
    print(f"Try importing '{OUTPUT_FILE}' into TD Snap")
    print("="*80)

if __name__ == "__main__":
    main()
