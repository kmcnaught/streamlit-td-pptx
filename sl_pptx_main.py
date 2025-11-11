"""
PowerPoint Speaker Notes to TD Snap Pageset Converter
Main Streamlit application
"""

import streamlit as st
import os
from pptx_parser import extract_slides, split_notes, create_button_label, parse_pptx_to_buttons
from td_utils_simple import (
    create_temp_file,
    add_home_button,
    get_static_path,
    update_timestamps,
    update_page_title,
    add_buttons_from_pptx,
    check_existing_buttons
)

# Initialize session state
if 'slides_data' not in st.session_state:
    st.session_state.slides_data = None
if 'slides_with_notes_list' not in st.session_state:
    st.session_state.slides_with_notes_list = None
if 'split_levels' not in st.session_state:
    st.session_state.split_levels = {}
if 'default_split_level' not in st.session_state:
    st.session_state.default_split_level = 2
if 'max_label_length' not in st.session_state:
    st.session_state.max_label_length = 30


st.title('PowerPoint to TD Snap Pageset Converter')

st.markdown("""
This app converts PowerPoint speaker notes into a TD Snap communication pageset.
Each slide's notes will be split into button cells based on your preferences.
""")

# Step 1: Upload PowerPoint file
st.header("Step 1: Upload PowerPoint File")
pptx_file = st.file_uploader("Choose a PowerPoint file", type=['pptx'])

if pptx_file is not None:
    # Extract slides when file is uploaded
    if st.session_state.slides_data is None:
        # Create progress tracking elements
        progress_text = st.empty()
        progress_bar = st.progress(0)

        progress_text.text("📄 Extracting slides and speaker notes...")
        progress_bar.progress(50)

        st.session_state.slides_data = extract_slides(pptx_file)
        pptx_file.seek(0)  # Reset file pointer

        progress_text.text("✅ Extraction complete!")
        progress_bar.progress(100)

        # Clean up progress indicators immediately
        progress_text.empty()
        progress_bar.empty()

        # Cache the filtered slides list for performance
        slides_data = st.session_state.slides_data
        st.session_state.slides_with_notes_list = [s for s in slides_data if s['notes']]

    # Step 2: Configure splitting - show this ASAP
    st.header("Step 2: Configure Content Splitting")

    slides_data = st.session_state.slides_data

    # Show summary (calculate only once)
    total_slides = len(slides_data)
    slides_with_notes = len(st.session_state.slides_with_notes_list)

    st.success(f"Found {total_slides} slides, {slides_with_notes} with speaker notes")

    st.markdown("""
    **Split levels:**
    - **1**: Whole note (one button per slide)
    - **2**: Paragraphs (double line breaks)
    - **3**: Lines (single line breaks)
    - **4**: Sentences (period + space)
    """)

    # Global split level buttons
    st.subheader("Default split level for all slides")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("1: Whole", use_container_width=True,
                     type="primary" if st.session_state.default_split_level == 1 else "secondary"):
            st.session_state.default_split_level = 1
            st.rerun()

    with col2:
        if st.button("2: Paragraphs", use_container_width=True,
                     type="primary" if st.session_state.default_split_level == 2 else "secondary"):
            st.session_state.default_split_level = 2
            st.rerun()

    with col3:
        if st.button("3: Lines", use_container_width=True,
                     type="primary" if st.session_state.default_split_level == 3 else "secondary"):
            st.session_state.default_split_level = 3
            st.rerun()

    with col4:
        if st.button("4: Sentences", use_container_width=True,
                     type="primary" if st.session_state.default_split_level == 4 else "secondary"):
            st.session_state.default_split_level = 4
            st.rerun()

    default_level = st.session_state.default_split_level

    # Maximum label length control
    st.subheader("Maximum button label length")
    col1, col2, col3 = st.columns([1, 2, 1])

    with col1:
        if st.button("➖ Shorter", key="label_shorter", use_container_width=True):
            st.session_state.max_label_length = max(10, st.session_state.max_label_length - 5)
            st.rerun()

    with col2:
        st.markdown(f"<center><h3>{st.session_state.max_label_length} characters</h3></center>",
                   unsafe_allow_html=True)

    with col3:
        if st.button("➕ Longer", key="label_longer", use_container_width=True):
            st.session_state.max_label_length = min(60, st.session_state.max_label_length + 5)
            st.rerun()

    max_label_length = st.session_state.max_label_length

    # Preview with per-slide controls
    st.subheader("Preview and Adjust")

    # Use cached filtered slides list
    slides_with_notes_list = st.session_state.slides_with_notes_list

    if not slides_with_notes_list:
        st.warning("No slides with speaker notes found!")
    else:
        for slide in slides_with_notes_list:
            slide_num = slide['slide_num']
            title = slide['title']
            notes = slide['notes']

            # Use per-slide override or default
            current_level = st.session_state.split_levels.get(slide_num, default_level)

            # Split notes according to current level
            chunks = split_notes(notes, current_level)

            # Display slide preview
            with st.expander(f"📊 Slide {slide_num}: {title} ({len(chunks)} buttons)", expanded=False):
                # Per-slide controls
                col1, col2, col3 = st.columns([1, 1, 3])

                with col1:
                    if st.button("Split More ➕", key=f"more_{slide_num}", use_container_width=True):
                        new_level = min(4, current_level + 1)
                        st.session_state.split_levels[slide_num] = new_level
                        st.rerun()

                with col2:
                    if st.button("Split Less ➖", key=f"less_{slide_num}", use_container_width=True):
                        new_level = max(1, current_level - 1)
                        st.session_state.split_levels[slide_num] = new_level
                        st.rerun()

                with col3:
                    st.caption(f"Current split level: {current_level}")

                # Show preview of resulting buttons
                st.markdown("**Resulting buttons:**")
                for i, chunk in enumerate(chunks):
                    label = create_button_label(title, i, len(chunks), max_length=max_label_length)
                    preview_text = chunk[:100] + "..." if len(chunk) > 100 else chunk

                    st.markdown(f"""
                    **Cell {i+1}:** `{label}`
                    > {preview_text}
                    """)

    # Step 3: Upload blank pageset and process
    st.header("Step 3: Create TD Snap Pageset")

    db_file = st.file_uploader("Choose a blank TD Snap pageset (.spb)", type=['spb'])

    # Create a log expander for processing visibility
    log_expander = st.expander("Show Processing Logs", expanded=False)

    # Optional: Update pageset title
    st.write("Pageset name (optional)")
    update_title = st.checkbox('Update pageset title', value=True)

    pageset_title = None
    if update_title and pptx_file is not None:
        file_name, _ = os.path.splitext(pptx_file.name)
        pageset_title = st.text_input("Pageset name:", value=file_name)

    # Check for existing buttons and show warning if needed
    proceed_with_existing = True
    button_count = 0
    button_samples = []

    if db_file is not None:
        # Create temp file to check for existing buttons
        temp_check_path = create_temp_file(db_file)
        button_count, button_samples = check_existing_buttons(temp_check_path)

        # Clean up temp file
        try:
            os.remove(temp_check_path)
        except:
            pass

        if button_count > 0:
            st.warning(f"⚠️ Found {button_count} existing button(s) in this pageset")

            if button_samples:
                st.write("**Example buttons:**")
                for i, label in enumerate(button_samples, 1):
                    st.write(f"{i}. {label}")

            proceed_with_existing = st.checkbox(
                "I understand this will add buttons to existing content, using only empty cells",
                value=False
            )

    # Process button (only enabled if no existing buttons or user confirmed)
    button_disabled = button_count > 0 and not proceed_with_existing

    if db_file is not None and not button_disabled and st.button("Create Pageset", type="primary"):
        try:
            # Create progress placeholder
            progress_text = st.empty()
            progress_bar = st.progress(0)

            # Step 1: Parse PowerPoint
            progress_text.text("📄 Parsing PowerPoint file...")
            progress_bar.progress(10)
            pptx_file.seek(0)

            buttons_data = parse_pptx_to_buttons(
                pptx_file,
                split_levels=st.session_state.split_levels,
                default_level=default_level
            )

            if not buttons_data:
                progress_text.empty()
                progress_bar.empty()
                st.error("No buttons to create! Make sure slides have speaker notes.")
            else:
                progress_text.text(f"✓ Found {len(buttons_data)} buttons to create")
                progress_bar.progress(20)

                log_expander.write(f"📊 Parsed {len(buttons_data)} buttons from PowerPoint")

                # Step 2: Create temp copy
                progress_text.text("📋 Creating temporary pageset copy...")
                progress_bar.progress(30)
                temp_db_path = create_temp_file(db_file)

                # Check existing buttons in temp file
                existing_count, existing_samples = check_existing_buttons(temp_db_path)
                if existing_count > 0:
                    log_expander.write(f"📋 Found {existing_count} existing buttons in pageset")
                    if existing_samples:
                        log_expander.write(f"   Examples: {', '.join(existing_samples[:3])}")
                else:
                    log_expander.write("📋 Pageset is empty - starting fresh")

                # Step 3: Add home button
                progress_text.text("🏠 Adding home button...")
                progress_bar.progress(40)
                reference_db = get_static_path('home_button_ref.spb')
                add_home_button(temp_db_path, reference_db)

                if existing_count == 0:
                    log_expander.write("🏠 Added home button")

                # Step 4: Add buttons from PowerPoint
                progress_text.text(f"➕ Adding {len(buttons_data)} buttons to pageset...")
                progress_bar.progress(50)
                log_expander.write(f"➕ Adding {len(buttons_data)} buttons to available positions...")
                num_added = add_buttons_from_pptx(temp_db_path, buttons_data)
                progress_bar.progress(70)
                log_expander.write(f"✅ Successfully added {num_added} buttons")

                # Step 5: Update title
                if update_title and pageset_title:
                    progress_text.text("📝 Updating pageset title...")
                    progress_bar.progress(80)
                    update_page_title(temp_db_path, pageset_title)

                # Step 6: Update timestamps
                progress_text.text("🕒 Updating timestamps...")
                progress_bar.progress(90)
                update_timestamps(temp_db_path)

                # Step 7: Read for download
                progress_text.text("💾 Preparing download...")
                progress_bar.progress(95)
                with open(temp_db_path, 'rb') as f:
                    modified_db = f.read()

                progress_bar.progress(100)
                progress_text.text("✅ Complete!")

                # Clean up progress indicators
                import time
                time.sleep(0.5)
                progress_text.empty()
                progress_bar.empty()

                # Offer download
                download_name = f"{pageset_title}.spb" if pageset_title else "modified_pageset.spb"

                st.success(f"✅ Created pageset with {num_added} buttons!")

                st.download_button(
                    label="Download TD Snap Pageset",
                    data=modified_db,
                    file_name=download_name,
                    mime="application/octet-stream"
                )

                # Clean up temp file
                try:
                    os.remove(temp_db_path)
                except:
                    pass

        except Exception as e:
            st.error(f"Error creating pageset: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("👆 Upload a PowerPoint file to get started")

# Sidebar with instructions
with st.sidebar:
    st.markdown("""
    ## How to Use

    1. **Upload PowerPoint**: Select a .pptx file with speaker notes
    2. **Configure Splitting**: Choose split level and label length
       - Use buttons to set default split level
       - Use ➕/➖ to adjust label length
       - Fine-tune individual slides with per-slide controls
    3. **Upload Blank Pageset**: Select a blank TD Snap .spb file
    4. **Create**: Click to generate your pageset

    ## Features

    - Preview all button content before creating
    - Per-slide split control
    - Adjustable button label length
    - Automatic color coding by slide
    - Real-time progress feedback
    - Speaker notes become button messages

    ## Split Levels

    - **Level 1**: Whole note (one button per slide)
    - **Level 2**: Paragraphs (recommended)
    - **Level 3**: Lines
    - **Level 4**: Sentences
    """)
