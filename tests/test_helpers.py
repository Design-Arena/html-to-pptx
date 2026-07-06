"""Unit tests for pure-Python helper functions in converter.py.

These tests do not launch a browser and run very fast.
They cover color parsing, gradient parsing, font parsing,
border-radius extraction, and element-tree utilities.
"""

from __future__ import annotations

import pytest
from pptx.dml.color import RGBColor

from html_to_pptx.converter import (
    _any_descendant_has_text,
    _count_elements,
    _css_color_to_rgb,
    _first_gradient_color,
    _gradient_angle_emu,
    _parse_border_radius_px,
    _parse_css_gradient,
    _parse_font_family,
    _resolve_font_color,
)


# ---------------------------------------------------------------------------
# _css_color_to_rgb
# ---------------------------------------------------------------------------


class TestCssColorToRgb:
    def test_opaque_rgb(self):
        result = _css_color_to_rgb("rgb(255, 0, 128)")
        assert result is not None
        color, alpha = result
        assert color == RGBColor(255, 0, 128)
        assert alpha == 1.0

    def test_rgba_with_alpha(self):
        result = _css_color_to_rgb("rgba(10, 20, 30, 0.5)")
        assert result is not None
        color, alpha = result
        assert color == RGBColor(10, 20, 30)
        assert abs(alpha - 0.5) < 0.001

    def test_rgba_fully_transparent_returns_none(self):
        assert _css_color_to_rgb("rgba(0, 0, 0, 0)") is None

    def test_rgba_near_transparent_returns_none(self):
        assert _css_color_to_rgb("rgba(255, 255, 255, 0.005)") is None

    def test_hex_six_digit(self):
        result = _css_color_to_rgb("#1e293b")
        assert result is not None
        color, alpha = result
        assert color == RGBColor(0x1E, 0x29, 0x3B)
        assert alpha == 1.0

    def test_hex_three_digit(self):
        result = _css_color_to_rgb("#f0a")
        assert result is not None
        color, _ = result
        assert color == RGBColor(0xFF, 0x00, 0xAA)

    def test_transparent_keyword_returns_none(self):
        assert _css_color_to_rgb("transparent") is None

    def test_empty_string_returns_none(self):
        assert _css_color_to_rgb("") is None


# ---------------------------------------------------------------------------
# _parse_font_family
# ---------------------------------------------------------------------------


class TestParseFontFamily:
    def test_single_font(self):
        assert _parse_font_family("Arial") == "Arial"

    def test_first_of_multiple(self):
        assert _parse_font_family("Inter, sans-serif") == "Inter"

    def test_strips_quotes(self):
        assert _parse_font_family('"Helvetica Neue", Arial') == "Helvetica Neue"

    def test_strips_single_quotes(self):
        assert _parse_font_family("'Open Sans', sans-serif") == "Open Sans"


# ---------------------------------------------------------------------------
# _parse_css_gradient
# ---------------------------------------------------------------------------


class TestParseCssGradient:
    def test_two_stop_gradient(self):
        css = "linear-gradient(90deg, rgb(0, 0, 255), rgb(255, 0, 0))"
        stops = _parse_css_gradient(css)
        assert stops is not None
        assert len(stops) == 2
        assert stops[0][1] == pytest.approx(0.0)
        assert stops[1][1] == pytest.approx(1.0)

    def test_three_stop_positions(self):
        css = "linear-gradient(180deg, rgb(0,0,0), rgb(128,128,128), rgb(255,255,255))"
        stops = _parse_css_gradient(css)
        assert stops is not None
        assert len(stops) == 3
        assert stops[1][1] == pytest.approx(0.5)

    def test_single_stop_returns_none(self):
        css = "linear-gradient(90deg, rgb(0, 0, 255))"
        assert _parse_css_gradient(css) is None

    def test_non_gradient_returns_none(self):
        assert _parse_css_gradient("none") is None
        assert _parse_css_gradient("") is None
        assert _parse_css_gradient(None) is None

    def test_rgba_transparent_stop(self):
        css = "linear-gradient(90deg, rgba(0,0,0,0), rgb(255,255,255))"
        stops = _parse_css_gradient(css)
        assert stops is not None
        _, _, alpha_first = stops[0]
        assert alpha_first == pytest.approx(0.0)


# ---------------------------------------------------------------------------
# _gradient_angle_emu
# ---------------------------------------------------------------------------


class TestGradientAngleEmu:
    def test_180deg_to_bottom(self):
        # CSS 180deg (to bottom) → OOXML (180+270)%360=90 → 90*60000=5400000
        angle = _gradient_angle_emu("linear-gradient(180deg, red, blue)")
        assert angle == 5_400_000

    def test_0deg_to_top(self):
        # CSS 0deg (to top) → OOXML (0+270)%360=270 → 270*60000=16200000
        angle = _gradient_angle_emu("linear-gradient(0deg, red, blue)")
        assert angle == 16_200_000

    def test_90deg_to_right(self):
        # CSS 90deg (to right) → OOXML (90+270)%360=0 → 0*60000=0
        angle = _gradient_angle_emu("linear-gradient(90deg, red, blue)")
        assert angle == 0

    def test_keyword_to_bottom(self):
        angle = _gradient_angle_emu("linear-gradient(to bottom, red, blue)")
        assert angle == 5_400_000

    def test_keyword_to_right(self):
        angle = _gradient_angle_emu("linear-gradient(to right, red, blue)")
        assert angle == 0

    def test_no_angle_defaults_to_180deg(self):
        # No angle keyword → defaults to 180deg
        angle = _gradient_angle_emu("linear-gradient(red, blue)")
        assert angle == 5_400_000


# ---------------------------------------------------------------------------
# _parse_border_radius_px
# ---------------------------------------------------------------------------


class TestParseBorderRadiusPx:
    def test_pixel_value(self):
        el = {"borderRadius": "16px", "width": 100, "height": 50}
        assert _parse_border_radius_px(el) == pytest.approx(16.0)

    def test_zero(self):
        el = {"borderRadius": "0", "width": 100, "height": 50}
        assert _parse_border_radius_px(el) == 0.0

    def test_percentage_resolves_against_min_dimension(self):
        # 50% on a 100x60 element → 50% of min(100,60)=60 → 30
        el = {"borderRadius": "50%", "width": 100, "height": 60}
        assert _parse_border_radius_px(el) == pytest.approx(30.0)

    def test_missing_key_returns_zero(self):
        assert _parse_border_radius_px({}) == 0.0

    def test_invalid_value_returns_zero(self):
        el = {"borderRadius": "auto", "width": 100, "height": 50}
        assert _parse_border_radius_px(el) == 0.0


# ---------------------------------------------------------------------------
# _any_descendant_has_text and _count_elements
# ---------------------------------------------------------------------------


class TestElementTreeHelpers:
    def _make_el(self, text="", children=None):
        return {"text": text, "children": children or []}

    def test_no_text_no_children(self):
        assert _any_descendant_has_text(self._make_el()) is False

    def test_direct_child_has_text(self):
        el = self._make_el(children=[self._make_el("hello")])
        assert _any_descendant_has_text(el) is True

    def test_grandchild_has_text(self):
        grandchild = self._make_el("deep")
        child = self._make_el(children=[grandchild])
        el = self._make_el(children=[child])
        assert _any_descendant_has_text(el) is True

    def test_whitespace_only_child_not_counted(self):
        el = self._make_el(children=[self._make_el("   ")])
        assert _any_descendant_has_text(el) is False

    def test_count_flat_list(self):
        elements = [self._make_el(), self._make_el(), self._make_el()]
        assert _count_elements(elements) == 3

    def test_count_nested(self):
        child = self._make_el(children=[self._make_el()])
        parent = self._make_el(children=[child])
        assert _count_elements([parent]) == 3  # parent + child + grandchild


# ---------------------------------------------------------------------------
# _first_gradient_color and _resolve_font_color
# ---------------------------------------------------------------------------


class TestFirstGradientColor:
    def test_returns_first_opaque_stop(self):
        css = "linear-gradient(90deg, rgb(255, 0, 0), rgb(0, 0, 255))"
        color = _first_gradient_color(css)
        assert color == RGBColor(255, 0, 0)

    def test_skips_transparent_stops(self):
        css = "linear-gradient(90deg, rgba(0,0,0,0), rgb(0, 255, 0))"
        color = _first_gradient_color(css)
        assert color == RGBColor(0, 255, 0)

    def test_non_gradient_returns_none(self):
        assert _first_gradient_color("none") is None
        assert _first_gradient_color("") is None


class TestResolveFontColor:
    def test_plain_color(self):
        el = {"color": "rgb(255, 255, 255)", "isGradientText": False}
        assert _resolve_font_color(el) == RGBColor(255, 255, 255)

    def test_gradient_text_uses_first_stop(self):
        el = {
            "isGradientText": True,
            "backgroundImage": "linear-gradient(90deg, rgb(200, 100, 50), rgb(0, 0, 255))",
            "color": "rgba(0,0,0,0)",
        }
        assert _resolve_font_color(el) == RGBColor(200, 100, 50)

    def test_fallback_when_color_missing(self):
        # When color is missing/unparseable, falls back to white
        el = {"color": "", "isGradientText": False}
        assert _resolve_font_color(el) == RGBColor(0xFF, 0xFF, 0xFF)
