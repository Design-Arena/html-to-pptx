"""Canvas geometry shared by measurement and rendering.

The HTML deck is authored on a fixed 1920×1080 CSS-pixel canvas and mapped onto
a 13.333×7.5 in widescreen slide.
"""

from __future__ import annotations

SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5

SLIDE_CANVAS_WIDTH_PX = 1920
SLIDE_CANVAS_HEIGHT_PX = 1080

PIXELS_TO_INCHES_X = SLIDE_WIDTH_INCHES / SLIDE_CANVAS_WIDTH_PX
PIXELS_TO_INCHES_Y = SLIDE_HEIGHT_INCHES / SLIDE_CANVAS_HEIGHT_PX

EMU_PER_INCH = 914_400
EMU_PER_CSS_PX = round(PIXELS_TO_INCHES_X * EMU_PER_INCH)

# Font size scaling: CSS px -> PowerPoint pt, adjusted for the same layout scale
# as position coordinates so text proportions match. 96 CSS px = 1 inch at
# standard DPI; the ratio ensures font sizes and element positions use the same
# scale factor (1 CSS px ≈ 0.5 pt on this canvas).
_LAYOUT_SCALE = SLIDE_WIDTH_INCHES / (SLIDE_CANVAS_WIDTH_PX / 96.0)
CSS_PX_TO_PT = 0.75 * _LAYOUT_SCALE
