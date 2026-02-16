"""
Simplified TD Snap database utilities for PowerPoint app
Adapted from streamlit-tdpages/TDutils.py
"""

import sqlite3
import uuid
import os
import datetime
import tempfile
import shutil
from colour_simple import get_color_for_slide, rgb_to_int, int_to_rgb


def get_static_path(fname):
    """Get path to file in static directory."""
    fname = os.path.join(os.path.dirname(__file__), 'static/' + fname)
    fname = os.path.abspath(fname)
    return fname


def create_temp_file(file_obj, extension='.spb'):
    """Create a temporary file from uploaded file object."""
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    temp_filename = f"pageset_{timestamp}{extension}"
    temp_path = os.path.join(temp_dir, temp_filename)

    with open(temp_path, 'wb') as f:
        f.write(file_obj.getvalue())

    return temp_path


def get_page_layout_details(db_filename):
    """
    Get page ID and layout details from TD Snap database.

    Returns:
        tuple: (pageId, layouts)
            pageId: int
            layouts: list of tuples (pageLayoutId, num_columns, num_rows)
    """
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()

    # Select rows from Page table excluding ignored titles
    rows_to_ignore = ['Dashboard', 'Message Bar']
    query = "SELECT Id, Title FROM Page WHERE Title NOT IN ({})".format(','.join('?' for _ in rows_to_ignore))
    cursor.execute(query, rows_to_ignore)
    rows = cursor.fetchall()

    if len(rows) == 0:
        conn.close()
        raise ValueError("Error: Couldn't find Page row")
    if len(rows) != 1:
        error_message = 'Error: Found multiple Page IDs in file: ' + ', '.join(row[1] for row in rows)
        conn.close()
        raise ValueError(error_message)

    pageId = rows[0][0]

    # Retrieve PageLayoutSetting for the unique Id
    cursor.execute("SELECT Id, PageLayoutSetting FROM PageLayout WHERE PageId = ?", (pageId,))
    settings = cursor.fetchall()

    # Organize data into a list of (pageLayoutId, num_columns, num_rows)
    layouts = [(setting[0], *map(int, setting[1].split(',')[:2])) for setting in settings]

    conn.close()
    return pageId, layouts


def find_available_positions(db_filename, pageLayoutId, ncols, nrows):
    """
    Find available grid positions for buttons.

    Args:
        db_filename: Path to TD Snap database
        pageLayoutId: Page layout ID
        ncols: Number of columns
        nrows: Number of rows

    Returns:
        list: Available (col, row) positions
    """
    # Generate all possible positions across 10 pages
    npages = 10
    all_positions = [
        (c, r) for r in range(nrows * npages) for c in range(ncols)
        if not (
            # Exclude bottom right of any page
            (c == ncols - 1 and r % nrows == nrows - 1) or
            # Exclude top right of any page except the first one
            (c == ncols - 1 and r % nrows == 0 and r != 0)
        )
    ]

    # Connect to DB and fetch occupied positions
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()
    cursor.execute("SELECT GridPosition FROM ElementPlacement WHERE PageLayoutId = ?", (pageLayoutId,))
    occupied_positions_raw = cursor.fetchall()

    # Parse 'c, r' format and convert to list of tuples
    occupied_positions = []
    for pos_raw in occupied_positions_raw:
        c, r = map(int, pos_raw[0].split(','))
        occupied_positions.append((c, r))

    # Filter out occupied positions
    available_positions = [pos for pos in all_positions if pos not in occupied_positions]

    conn.close()
    return available_positions


def add_button(cursor, buttonId, refId, label, message, symbol=None):
    """
    Add a button to the database.

    Args:
        cursor: Database cursor
        buttonId: Button ID
        refId: Element reference ID
        label: Button label text
        message: Message to speak when button pressed
        symbol: Library symbol ID (optional, None for no symbol)
    """
    label_ownership = 0 if label is None else 3
    image_ownership = 0 if symbol is None else 3

    # TEMP FIX: Strip newlines from message - TD Snap may not support multi-line messages
    if message:
        message = message.replace('\n', ' ').replace('\r', ' ')

    new_uuid = str(uuid.uuid1())
    cursor.execute("""
        INSERT INTO Button (Id, Label, Message, ImageOwnership, BorderColor, BorderThickness, LabelOwnership, CommandFlags, ContentType, UniqueId, ElementReferenceId, ActiveContentType, LibrarySymbolId, PageSetImageId, SymbolColorDataId, MessageRecordingId)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (buttonId, label, message, image_ownership, '-132102', 0.0, label_ownership, 8, 6, new_uuid, refId, 0, symbol, 0, 0, 0))


def add_command_speak_message(cursor, buttonId):
    """Add 'speak message' command to button."""
    content = '\'{"$type":"1","$values":[{"$type":"3","MessageAction":0}]}\''
    cursor.execute("""
        INSERT INTO CommandSequence (SerializedCommands, ButtonId)
        VALUES ({}, "{}")
    """.format(content, buttonId))


def add_element_reference(cursor, refId, pageId, color):
    """
    Add element reference with specified color.

    Args:
        cursor: Database cursor
        refId: Element reference ID
        pageId: Page ID
        color: Background color (32-bit integer)
    """
    foregroundColor = '-14934754'  # Standard foreground color

    cursor.execute("""
        INSERT INTO ElementReference
        (Id, ElementType, ForegroundColor, BackgroundColor, AudioCueRecordingId, PageId)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (refId, 0, foregroundColor, color, 0, pageId))


def add_button_placement(cursor, pageLayoutId, elementRefId, position):
    """
    Place button on grid.

    Args:
        cursor: Database cursor
        pageLayoutId: Page layout ID
        elementRefId: Element reference ID
        position: (col, row) tuple
    """
    c, r = position
    cursor.execute(f"""
        INSERT INTO ElementPlacement
        (GridPosition, GridSpan, Visible, ElementReferenceId, PageLayoutId)
        VALUES ('{c},{r}', '1,1', '1', '{elementRefId}', '{pageLayoutId}')
    """)


def get_next_id(cursor, table_name):
    """Get next available ID from sqlite_sequence table."""
    cursor.execute("SELECT seq FROM sqlite_sequence WHERE name=?", (table_name,))
    result = cursor.fetchone()

    next_id = 1
    if result is not None:
        next_id = result[0] + 1

    return next_id


def dt_to_filetime(dt):
    """Convert Python datetime to Windows FILETIME."""
    delta = dt - datetime.datetime(1, 1, 1)
    filetime = int(delta.total_seconds() * 10**7)
    return filetime


def get_timestamp():
    """Get current timestamp in Windows FILETIME format."""
    return dt_to_filetime(datetime.datetime.now())


def update_timestamps(filename):
    """Update all synchronization timestamps in pageset to current time."""
    try:
        conn = sqlite3.connect(filename)
        cursor = conn.cursor()

        new_timestamp = get_timestamp()
        cursor.execute("UPDATE Page SET TimeStamp = ?", (new_timestamp,))
        cursor.execute("UPDATE Synchronization SET PageSetTimestamp = ?", (new_timestamp,))
        cursor.execute("UPDATE PageSetProperties SET TimeStamp = ?", (new_timestamp,))

        conn.commit()
    finally:
        if conn:
            conn.close()


def update_page_title(db_path, new_name, page_set_id=1):
    """Update the pageset title."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Update the FriendlyName in PageSetProperties
        cursor.execute("UPDATE PageSetProperties SET FriendlyName = ? WHERE ID = ?",
                      (new_name, page_set_id))

        # Find page where Title is not 'Dashboard' or 'Message Bar'
        cursor.execute("SELECT Id FROM Page WHERE Title NOT IN ('Dashboard', 'Message Bar')")
        page_row = cursor.fetchone()

        if page_row:
            page_id = page_row[0]
            cursor.execute("UPDATE Page SET Title = ? WHERE Id = ?", (new_name, page_id))

        conn.commit()
    finally:
        conn.close()


def update_page_grid_dimension(db_path, grid_dimension=None):
    """
    Update Page.GridDimension field.

    Args:
        db_path: Path to TD Snap database
        grid_dimension: String like "3,3" or None/NULL to match PageSet setting
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Find the main page (excluding Dashboard/Message Bar)
        cursor.execute("SELECT Id FROM Page WHERE Title NOT IN ('Dashboard', 'Message Bar')")
        page_row = cursor.fetchone()

        if page_row:
            page_id = page_row[0]
            cursor.execute("UPDATE Page SET GridDimension = ? WHERE Id = ?", (grid_dimension, page_id))

        conn.commit()
    finally:
        conn.close()


def check_existing_buttons(db_filename):
    """
    Check if buttons already exist in database.

    Returns:
        tuple: (button_count, button_samples)
            button_count: int - number of existing buttons
            button_samples: list - first 3 button labels
    """
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT COUNT(*) FROM Button")
        button_count = cursor.fetchone()[0]

        button_samples = []
        if button_count > 0:
            cursor.execute("SELECT Label FROM Button LIMIT 3")
            button_samples = [row[0] for row in cursor.fetchall() if row[0]]

        return button_count, button_samples
    finally:
        conn.close()


def get_grid_capacity(db_path):
    """
    Calculate grid capacity information for the pageset.

    Args:
        db_path: Path to TD Snap database

    Returns:
        dict: Grid capacity information
            - ncols: Number of columns
            - nrows: Number of rows
            - total_pages: Total number of pages (hardcoded to 10)
            - reserved_cells: Number of cells reserved for navigation (19)
            - occupied_cells: Number of cells already occupied by buttons (max across all layouts)
            - available_cells: Number of cells available for new buttons (minimum across all layouts)
            - cells_per_page: Grid cells per page (ncols × nrows)
    """
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Get page and layout info
        pageId, layouts = get_page_layout_details(db_path)

        # Use first layout to get grid dimensions (assume all layouts have same dimensions)
        first_layout = layouts[0]
        layoutId, ncols, nrows = first_layout

        # Constants
        total_pages = 10  # Hardcoded in find_available_positions
        reserved_cells = 19  # 10 bottom-right (home) + 9 top-right (navigation on pages 2-10)

        # Check ALL layouts and find the minimum available cells
        # (different layouts may have different numbers of occupied cells)
        min_available = float('inf')
        max_occupied = 0
        limiting_layout_index = 0
        layout_info = []

        for idx, (layoutId, layout_ncols, layout_nrows) in enumerate(layouts):
            available_positions = find_available_positions(db_path, layoutId, layout_ncols, layout_nrows)
            num_available = len(available_positions)

            # Count occupied for this layout
            cursor.execute(
                "SELECT COUNT(*) FROM ElementPlacement WHERE PageLayoutId = ?",
                (layoutId,)
            )
            occupied = cursor.fetchone()[0]

            # Track minimum available
            if num_available < min_available:
                min_available = num_available
                limiting_layout_index = idx

            if occupied > max_occupied:
                max_occupied = occupied

            # Store layout info
            layout_info.append({
                'id': layoutId,
                'ncols': layout_ncols,
                'nrows': layout_nrows,
                'cells_per_page': layout_ncols * layout_nrows,
                'available_cells': num_available,
                'occupied_cells': occupied
            })

        # Calculate cells per page for first layout (for backward compatibility)
        cells_per_page = ncols * nrows

        return {
            'ncols': ncols,
            'nrows': nrows,
            'total_pages': total_pages,
            'reserved_cells': reserved_cells,
            'occupied_cells': max_occupied,
            'available_cells': min_available,
            'cells_per_page': cells_per_page,
            'layouts': layout_info,
            'limiting_layout_index': limiting_layout_index
        }

    finally:
        conn.close()


def add_home_button(pageset_db_filename, reference_db_filename):
    """
    Copy home button from reference database to pageset.

    Note: Only adds home button if no buttons exist yet.
    Use add_buttons_from_pptx() to add content buttons - it handles existing buttons properly.
    """
    pageId, layouts = get_page_layout_details(pageset_db_filename)

    conn_pageset = sqlite3.connect(pageset_db_filename)

    try:
        cursor_pageset = conn_pageset.cursor()
        cursor_pageset.execute("SELECT COUNT(*) FROM Button")
        if cursor_pageset.fetchone()[0] > 0:
            # Buttons already exist, skip adding home button
            # (it may already be there or user wants to preserve existing layout)
            conn_pageset.close()
            return

        # Attach the reference database
        conn_pageset.execute(f"ATTACH DATABASE '{reference_db_filename}' AS ref_db")

        # Copy entries from reference database
        conn_pageset.execute("INSERT INTO Button SELECT * FROM ref_db.Button")
        conn_pageset.execute("INSERT INTO ElementReference SELECT * FROM ref_db.ElementReference")
        conn_pageset.execute("INSERT INTO ElementPlacement SELECT * FROM ref_db.ElementPlacement")
        conn_pageset.execute("INSERT INTO CommandSequence SELECT * FROM ref_db.CommandSequence")

        # Update page IDs - need to properly duplicate ElementPlacement for each layout
        cursor = conn_pageset.cursor()

        # Update ElementReference with correct PageId
        cursor.execute("UPDATE ElementReference SET PageId = ?", (pageId,))

        # For ElementPlacement, we need to duplicate the home button row for each layout
        # Get the home button's ElementReferenceId
        cursor.execute("SELECT ElementReferenceId FROM Button WHERE Label = 'Home'")
        home_ref_id = cursor.fetchone()
        if home_ref_id:
            home_ref_id = home_ref_id[0]

            # Get the existing ElementPlacement row for the home button
            cursor.execute("SELECT * FROM ElementPlacement WHERE ElementReferenceId = ?", (home_ref_id,))
            existing_row = cursor.fetchone()

            if existing_row:
                # Get column info
                cursor.execute("PRAGMA table_info(ElementPlacement)")
                columns = cursor.fetchall()
                column_names = [col[1] for col in columns]

                # Get the ID of the existing row so we can delete it later
                existing_row_id = existing_row[column_names.index('Id')]

                # Duplicate the row for each layout
                for layoutId, _, _ in layouts:
                    new_row = list(existing_row)
                    new_row[column_names.index('PageLayoutId')] = layoutId
                    new_row[column_names.index('Id')] = None  # Let SQLite auto-generate

                    placeholders = ', '.join(['?' for _ in column_names])
                    cursor.execute(f"INSERT INTO ElementPlacement ({', '.join(column_names)}) VALUES ({placeholders})", new_row)

                # Delete the original row
                cursor.execute("DELETE FROM ElementPlacement WHERE Id = ?", (existing_row_id,))

        conn_pageset.commit()
        conn_pageset.execute("DETACH DATABASE ref_db")

    finally:
        conn_pageset.close()


def add_buttons_from_pptx(db_path, buttons_data, selected_layout_ids=None):
    """
    Add buttons to TD Snap database from PowerPoint data.

    Args:
        db_path: Path to TD Snap database
        buttons_data: List of tuples (label, message, slide_num)
        selected_layout_ids: List of layout IDs to populate (default None means all layouts)

    Returns:
        Number of buttons added
    """
    if not buttons_data:
        return 0

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Get page and layout info
        pageId, layouts = get_page_layout_details(db_path)

        # Filter to selected layouts (or all if none specified)
        layouts_to_use = layouts if selected_layout_ids is None else [
            (lid, nc, nr) for lid, nc, nr in layouts if lid in selected_layout_ids
        ]

        # Validate space availability for SELECTED layouts only
        min_available = float('inf')
        limiting_layout = None

        for layoutId, ncols, nrows in layouts_to_use:
            available_positions = find_available_positions(db_path, layoutId, ncols, nrows)
            if len(available_positions) < min_available:
                min_available = len(available_positions)
                limiting_layout = (layoutId, ncols, nrows)

        if min_available < len(buttons_data):
            shortage = len(buttons_data) - min_available
            raise ValueError(
                f"Not enough grid space!\n\n"
                f"Required: {len(buttons_data)} cells\n"
                f"Available: {min_available} cells (limited by layout {limiting_layout[0]})\n"
                f"Shortage: {shortage} cells\n\n"
                f"Solutions:\n"
                f"- Reduce split level (use fewer buttons per slide)\n"
                f"- Use a blank pageset with a larger grid"
            )

        # Get starting IDs
        buttonId = get_next_id(cursor, 'Button')
        refId = get_next_id(cursor, 'ElementReference')

        # Add buttons, commands, and element references ONCE
        for i, (label, message, slide_num) in enumerate(buttons_data):
            print(f"Adding button: {label} (Slide {slide_num})")
            current_buttonId = buttonId + i
            current_refId = refId + i

            # Get color for this slide
            color = get_color_for_slide(slide_num)

            # Add button (no symbol)
            add_button(cursor, current_buttonId, current_refId, label, message, symbol=None)

            # Add speak message command
            add_command_speak_message(cursor, current_buttonId)

            # Add element reference with slide color
            add_element_reference(cursor, current_refId, pageId, color)

        # Add button placements for SELECTED layouts only
        for layoutId, ncols, nrows in layouts_to_use:
            # Find available positions for this layout
            available_positions = find_available_positions(db_path, layoutId, ncols, nrows)

            # Add placement for each button
            for i, (label, message, slide_num) in enumerate(buttons_data):
                current_refId = refId + i
                position = available_positions[i]

                # Add button placement
                add_button_placement(cursor, layoutId, current_refId, position)

        conn.commit()
        return len(buttons_data)

    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()
