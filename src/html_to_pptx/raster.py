"""Browser raster pass for paint PPTX cannot express natively.

EXTRACTION_JS marks such elements with a ``rasterMode``:

* ``subtree`` — the element and everything inside it becomes one picture
  (inline SVG, canvas, video, WebP/AVIF/SVG/remote ``<img>``, filtered or
  clip-path'd decoration without text).
* ``paint`` — only the element's own background, border, shadow and
  pseudo-elements become a picture (radial/conic/layered/url() backgrounds,
  glows); its text and children stay native, editable shapes.

Each target is screenshotted by Chromium in isolation, so the picture is
pixel-identical to the browser render instead of an approximation.
"""

from __future__ import annotations

import base64
import io
import logging

from PIL import Image

from .measurement_tree import walk_elements

logger = logging.getLogger(__name__)

JPEG_QUALITY = 90
OPAQUE_ALPHA = 255
# Photos and full-bleed backgrounds without transparency are stored as JPEG;
# icons and charts keep PNG so edges stay crisp.
JPEG_ELIGIBLE_TAGS = {"img", "video"}
# Averaging a capture down to this many pixels per side is enough for a backdrop color.
BACKDROP_SAMPLE_PX = 32
# A capture covering less than this share of its box says little about what
# sits behind the element's children, so it is not used as their backdrop.
MIN_BACKDROP_COVERAGE = 0.6

# Fields describing paint that a ``paint`` capture now carries as pixels.
_CAPTURED_PAINT_FIELDS = {
    "backgroundColor": None,
    "backgroundImage": None,
    "borderColor": None,
    "borderWidth": 0,
    "borderLeftColor": None,
    "borderLeftWidth": 0,
    "borderLeftStyle": None,
    "borderSides": [],
    "boxShadow": None,
}


def _encode(png: bytes, image: Image.Image, allow_jpeg: bool) -> str:
    if allow_jpeg and image.getchannel("A").getextrema()[0] == OPAQUE_ALPHA:
        buffer = io.BytesIO()
        image.convert("RGB").save(buffer, "JPEG", quality=JPEG_QUALITY)
        return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode()
    return "data:image/png;base64," + base64.b64encode(png).decode()


def _backdrop_rgb(image: Image.Image) -> list[int] | None:
    """Alpha-weighted mean color of a capture, used to flatten children's alpha."""
    sample = image.resize((BACKDROP_SAMPLE_PX, BACKDROP_SAMPLE_PX), Image.Resampling.BOX)
    access = sample.load()
    pixels = [access[x, y] for x in range(BACKDROP_SAMPLE_PX) for y in range(BACKDROP_SAMPLE_PX)]
    total_alpha = sum(p[3] for p in pixels)
    if total_alpha < MIN_BACKDROP_COVERAGE * OPAQUE_ALPHA * len(pixels):
        return None
    return [round(sum(p[c] * p[3] for p in pixels) / total_alpha) for c in range(3)]


def _apply_capture(target: dict, png: bytes, geometry: dict) -> None:
    image = Image.open(io.BytesIO(png)).convert("RGBA")
    box = geometry["box"]
    if target["rasterMode"] == "subtree":
        target.update(
            isImage=True,
            isSvg=False,
            src=_encode(png, image, target.get("tag") in JPEG_ELIGIBLE_TAGS),
            children=[],
            objectFit=None,
            naturalWidth=0,
            naturalHeight=0,
            boxShadow=None,
            x=box["x"], y=box["y"], width=box["width"], height=box["height"],
        )
        # A padded capture (blur/shadow overflow) no longer matches the
        # element's corner geometry; the rounding is already in the pixels.
        if geometry["padded"]:
            target["borderRadius"] = "0px"
        return
    target.update(
        _CAPTURED_PAINT_FIELDS,
        paintSrc=_encode(png, image, allow_jpeg=True),
        paintBox=box,
        paintBackdrop=_backdrop_rgb(image),
    )


async def _capture(page, target: dict) -> None:
    mode = target["rasterMode"]
    try:
        geometry = await page.evaluate(
            "args => window.__htmlToPptx.beginCapture(args)",
            {"id": target["rasterId"], "mode": mode,
             "neutralizeTransform": bool(target.get("rotation"))},
        )
        if geometry is None:
            logger.warning("Raster target %s is gone or off-slide; skipping", target["rasterId"])
            return
        # Paint is mostly smooth gradients and blur, which gain nothing from
        # device pixels; subtrees (charts, icons, photos) keep full resolution.
        png = await page.screenshot(
            clip=geometry["clip"],
            omit_background=True,
            scale="device" if mode == "subtree" else "css",
        )
    except Exception:
        if page.is_closed():
            raise
        logger.warning("Failed to rasterize %s", target["rasterId"], exc_info=True)
        return
    finally:
        if not page.is_closed():
            await page.evaluate("() => window.__htmlToPptx.endCapture()")
    _apply_capture(target, png, geometry)


async def rasterize_marked_elements(page, measurements: list[dict]) -> None:
    """Screenshot every element EXTRACTION_JS marked and patch its measurement.

    Must run on the page EXTRACTION_JS was evaluated in. Subtree targets become
    image elements; paint targets gain ``paintSrc``/``paintBox`` pictures and
    lose the paint fields the picture now carries.
    """
    for index, slide_data in enumerate(measurements):
        targets = [el for el in walk_elements(slide_data["elements"]) if el.get("rasterId")]
        if slide_data.get("rasterId"):
            targets.insert(0, slide_data)
        if not targets:
            continue
        await page.evaluate("index => window.__htmlToPptx.prepareSlide(index)", index)
        try:
            for target in targets:
                await _capture(page, target)
        finally:
            if not page.is_closed():
                await page.evaluate("() => window.__htmlToPptx.finishSlide()")
