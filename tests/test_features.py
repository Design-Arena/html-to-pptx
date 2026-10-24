"""Unit tests for pure fidelity helpers (no browser needed): font substitution,
alpha flattening, gradient parsing, line-height, and alignment.
"""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from html_to_pptx import converter as C
from html_to_pptx.css_values import blend_over

# ---------------------------------------------------------------------------
# Font substitution stays in-family and rendering-safe
# ---------------------------------------------------------------------------


def test_font_resolution_in_family():
    assert C._resolve_pptx_font("'Playfair Display', Georgia, serif") == "Georgia"
    assert C._resolve_pptx_font("'Inter', 'Segoe UI', system-ui, sans-serif") == "Segoe UI"
    assert C._resolve_pptx_font("Roboto, Arial, sans-serif") == "Arial"
    assert C._resolve_pptx_font("'JetBrains Mono', monospace") == "Consolas"
    assert C._resolve_pptx_font("Montserrat, sans-serif") == "Century Gothic"
    assert C._resolve_pptx_font("'Totally Unknown', serif") == "Georgia"
    assert C._resolve_pptx_font("") == "Calibri"


# ---------------------------------------------------------------------------
# Alpha flattening over a backdrop
# ---------------------------------------------------------------------------


def test_blend_over_endpoints():
    black, white = RGBColor(0, 0, 0), RGBColor(255, 255, 255)
    assert tuple(blend_over(black, 1.0, white)) == (0, 0, 0)
    assert tuple(blend_over(black, 0.0, white)) == (255, 255, 255)
    mid = blend_over(black, 0.5, white)
    assert all(120 <= ch <= 135 for ch in mid)


def test_translucent_orange_over_navy_is_muted():
    navy = RGBColor(0x0F, 0x1B, 0x33)
    orange = RGBColor(0xE8, 0x77, 0x2E)
    out = blend_over(orange, 0.15, navy)
    # closer to navy than to orange
    assert out[0] < 0x60 and out[1] < 0x60


# ---------------------------------------------------------------------------
# Gradient parsing only accepts a single linear-gradient
# ---------------------------------------------------------------------------


def test_single_linear_gradient_parses():
    stops = C._parse_css_gradient("linear-gradient(90deg, rgb(0,0,0), rgb(255,255,255))")
    assert stops and len(stops) == 2


def test_radial_and_conic_and_multi_gradients_rejected():
    assert C._parse_css_gradient("radial-gradient(circle, rgb(0,0,0), rgb(1,1,1))") is None
    assert C._parse_css_gradient("conic-gradient(rgb(0,0,0) 0deg, rgb(1,1,1) 180deg)") is None
    assert C._parse_css_gradient(
        "radial-gradient(rgba(0,0,0,0.2), transparent), linear-gradient(rgb(1,1,1), rgb(2,2,2))"
    ) is None
    assert C._parse_css_gradient(
        "linear-gradient(rgb(1,1,1), rgb(2,2,2)), url(\"data:image/png;base64,AAAA\")"
    ) is None


def test_repeating_linear_gradient_is_not_one_gradient():
    """Stripes must reach the raster pass, not collapse into a two-stop fade."""
    assert C._parse_css_gradient(
        "repeating-linear-gradient(45deg, rgb(0,0,0) 0px, rgb(0,0,0) 10px, rgb(255,255,255) 10px)"
    ) is None


def test_gradient_stop_positions_are_honored():
    stops = C._parse_css_gradient(
        "linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 70%, rgb(0, 255, 0) 100%)"
    )
    assert [round(pos, 3) for _, pos, _ in stops] == [0.0, 0.7, 1.0]


def test_gradient_hard_stop_and_double_position():
    stops = C._parse_css_gradient(
        "linear-gradient(90deg, rgb(255, 0, 0) 0% 50%, rgb(0, 0, 255) 50% 100%)"
    )
    assert [round(pos, 3) for _, pos, _ in stops] == [0.0, 0.5, 0.5, 1.0]
    assert [tuple(c) for c, _, _ in stops][1:3] == [(255, 0, 0), (0, 0, 255)]


def test_omitted_gradient_positions_interpolate_between_known_ones():
    stops = C._parse_css_gradient(
        "linear-gradient(rgb(0, 0, 0), rgb(1, 1, 1) 20%, rgb(2, 2, 2), rgb(3, 3, 3), rgb(4, 4, 4))"
    )
    assert [round(pos, 3) for _, pos, _ in stops] == [0.0, 0.2, 0.467, 0.733, 1.0]


def test_gradient_angle_units():
    quarter_turn = C._gradient_angle_emu("linear-gradient(0.25turn, rgb(0,0,0), rgb(1,1,1))")
    assert quarter_turn == C._gradient_angle_emu("linear-gradient(90deg, rgb(0,0,0), rgb(1,1,1))")
    assert C._gradient_angle_emu("linear-gradient(-90deg, rgb(0,0,0), rgb(1,1,1))") == 180 * 60000


def test_translucent_gradient_stops_preserve_alpha():
    stops = C._parse_css_gradient(
        "linear-gradient(90deg, rgba(232,119,46,0.06), rgba(232,119,46,0.01))"
    )
    assert stops
    assert abs(stops[0][2] - 0.06) < 1e-6
    assert abs(stops[1][2] - 0.01) < 1e-6


# ---------------------------------------------------------------------------
# Line-height ratio
# ---------------------------------------------------------------------------


def test_line_height_ratio():
    assert abs(C._line_height_ratio({"lineHeight": "101.2px", "fontSize": 110}) - 0.92) < 1e-3
    assert C._line_height_ratio({"lineHeight": "normal", "fontSize": 20}) is None
    assert C._line_height_ratio({"lineHeight": "12px", "fontSize": 0}) is None


# ---------------------------------------------------------------------------
# CSS-driven alignment (flex, chips, tables)
# ---------------------------------------------------------------------------


def test_flex_center_maps_to_center_middle():
    el = {"display": "flex", "justifyContent": "center", "alignItems": "center",
          "textAlign": "start"}
    h, v = C._resolve_alignment(el, is_single_line=False, has_visual_bg=True)
    assert h == PP_ALIGN.CENTER and v == MSO_ANCHOR.MIDDLE


def test_table_cell_keeps_column_alignment():
    el = {"tag": "th", "display": "table-cell", "textAlign": "left", "backgroundColor": "rgb(0,0,0)"}
    h, _ = C._resolve_alignment(el, is_single_line=True, has_visual_bg=True)
    assert h == PP_ALIGN.LEFT  # not hijacked by chip-centering


def test_chip_centers_horizontally():
    el = {"display": "inline-block", "textAlign": "start", "backgroundColor": "rgb(0,0,0)",
          "width": 132, "paddingLeft": 16, "paddingRight": 16, "textWidth": 98}
    h, v = C._resolve_alignment(el, is_single_line=True, has_visual_bg=True)
    assert h == PP_ALIGN.CENTER and v == MSO_ANCHOR.MIDDLE


def test_full_width_row_keeps_text_at_start():
    el = {"display": "list-item", "textAlign": "start", "backgroundColor": "rgb(0,0,0)",
          "width": 1680, "paddingLeft": 24, "paddingRight": 24, "textWidth": 610}
    h, _ = C._resolve_alignment(el, is_single_line=True, has_visual_bg=True)
    assert h == PP_ALIGN.LEFT


# ---------------------------------------------------------------------------
# Font family class helper
# ---------------------------------------------------------------------------


def test_family_class():
    assert C._family_class("Georgia") == "serif"
    assert C._family_class("Consolas") == "mono"
    assert C._family_class("Segoe UI") == "sans"
