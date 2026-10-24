"""Parsing and compositing of computed CSS values (colors, layered lists)."""

from __future__ import annotations

import re

from pptx.dml.color import RGBColor

RGB_RE = re.compile(
    r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)"
)

# Below this alpha a color is invisible and treated as transparent.
MIN_VISIBLE_ALPHA = 0.01


def rgb_match_to_color(match: re.Match) -> tuple[RGBColor, float]:
    """Turn an ``RGB_RE`` match into (RGBColor, alpha), keeping alpha 0."""
    color = RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return color, float(match.group(4)) if match.group(4) else 1.0


def css_color_to_rgb(css_color: str | None) -> tuple[RGBColor, float] | None:
    """Parse a CSS color string to (RGBColor, opacity). Returns None for transparent."""
    if not css_color or css_color == "transparent":
        return None
    match = RGB_RE.match(css_color)
    if match:
        color, alpha = rgb_match_to_color(match)
        return None if alpha < MIN_VISIBLE_ALPHA else (color, alpha)
    if css_color.startswith("#"):
        h = css_color.lstrip("#")
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)), 1.0
    return None


def blend_over(color: RGBColor, alpha: float, backdrop: RGBColor) -> RGBColor:
    """Alpha-composite ``color`` at ``alpha`` over an opaque ``backdrop``.

    LibreOffice (and PowerPoint's PDF export) do not honor per-gradient-stop
    ``<a:alpha>`` and are unreliable with solid-fill / run-color alpha, so
    translucency is flattened to an equivalent opaque color at conversion time.
    This makes the PPTX render identically to the source HTML everywhere.
    """
    alpha = max(0.0, min(1.0, alpha))
    r = round(color[0] * alpha + backdrop[0] * (1 - alpha))
    g = round(color[1] * alpha + backdrop[1] * (1 - alpha))
    b = round(color[2] * alpha + backdrop[2] * (1 - alpha))
    return RGBColor(int(r), int(g), int(b))


def split_top_level(value: str) -> list[str]:
    """Split a CSS list on commas outside parentheses (layers, stops, shadows)."""
    parts: list[str] = []
    depth = 0
    start = 0
    for index, char in enumerate(value):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif char == "," and depth == 0:
            parts.append(value[start:index].strip())
            start = index + 1
    parts.append(value[start:].strip())
    return [part for part in parts if part]
