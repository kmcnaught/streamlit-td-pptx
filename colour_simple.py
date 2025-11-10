"""
Simple color palette for slide-based coloring
All buttons from same slide get the same color
"""

# Color palette (RGBA as 32-bit integers for TD Snap)
# Using a diverse set of colors that cycle through slides
SLIDE_COLORS = [
    4294951115,  # blue5
    4294934323,  # green2
    4294945095,  # red5
    4294956885,  # sand4
    4294956630,  # sand5
    4294946615,  # red7
    4294935067,  # green1
    4294951370,  # blue6
    4294967040,  # white1
    4289374890,  # purple
    4291624735,  # orange
    4286611584,  # teal
    4294962099,  # yellow
    4290822336,  # pink
    4282477025,  # cyan
    4286578816,  # light green
    4290019584,  # lavender
    4294309340,  # peach
    4280150535,  # slate blue
    4287245282,  # mint
]


def get_color_for_slide(slide_num: int) -> int:
    """
    Get color for a slide number.
    Colors cycle if more slides than colors.

    Args:
        slide_num: Slide number (1-based)

    Returns:
        32-bit color integer
    """
    index = (slide_num - 1) % len(SLIDE_COLORS)
    return SLIDE_COLORS[index]


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
