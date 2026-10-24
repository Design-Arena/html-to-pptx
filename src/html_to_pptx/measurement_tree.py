"""Traversal of the measured element tree produced by EXTRACTION_JS."""

from __future__ import annotations

from collections.abc import Iterator


def walk_elements(elements: list[dict]) -> Iterator[dict]:
    """Yield every element in a measurement tree (depth-first, paint order)."""
    for el in elements:
        yield el
        yield from walk_elements(el.get("children", []))


def count_elements(elements: list[dict]) -> int:
    """Count total elements in a measurement tree (including nested children)."""
    return sum(1 for _ in walk_elements(elements))
