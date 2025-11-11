"""
Simple color palette for slide-based coloring
All buttons from same slide get the same color
"""

# Orange colors (RGBA as 32-bit integers for TD Snap)
DARK_ORANGE = 4294951115   # Dark orange
LIGHT_ORANGE = 4294934323  # Light orange (peach)

def get_color_for_slide(slide_num: int) -> int:
    """
    Get color for a slide number.
    Alternates between dark and light orange

    Args:
        slide_num: Slide number (1-based)

    Returns:
        32-bit color integer
    """
    # Odd slide numbers (1, 3, 5...) get dark orange
    # Even slide numbers (2, 4, 6...) get light orange
    if slide_num % 2 == 1:
        return DARK_ORANGE
    else:
        return LIGHT_ORANGE


def rgb_to_int(r: int, g: int, b: int, a: int = 255) -> int:
    """Convert RGBA values to 32-bit integer for TD Snap."""
    return (a << 24) | (b << 16) | (g << 8) | r


def int_to_rgb(color_int: int) -> tuple:
    """Convert 32-bit integer to (R, G, B, A) tuple."""
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    a = (color_int >> 24) & 0xFF
    return (r, g, b, a)
