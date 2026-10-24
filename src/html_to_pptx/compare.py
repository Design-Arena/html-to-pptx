"""Visual comparison tool for verifying HTML-to-PPTX conversion quality.

For each HTML file:
  1. Screenshot each slide via Playwright (ground truth)
  2. Convert HTML -> PPTX via the converter
  3. Convert PPTX -> PDF via LibreOffice, then PDF pages -> PNGs
  4. Create side-by-side comparison images (HTML left, PPTX right)

Requires the ``compare`` extra: ``pip install html-to-pptx[compare]``
and LibreOffice (``soffice``) on PATH for PPTX-to-PNG rendering.

Usage::

    html-to-pptx-compare deck.html                  # full pipeline
    html-to-pptx-compare deck.html --only-convert    # skip comparison
    html-to-pptx-compare slides/*.html -o results/   # batch mode
"""

from __future__ import annotations

import argparse
import asyncio
import logging
import shutil
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

SLIDE_WIDTH_PX = 1920
SLIDE_HEIGHT_PX = 1080


async def screenshot_html_slides(html_path: Path, out_dir: Path) -> list[Path]:
    """Screenshot each slide, revealed exactly the way the converter reveals it."""
    from playwright.async_api import async_playwright

    from html_to_pptx.converter import SCRIPT_SETTLE_MS
    from html_to_pptx.extraction_js import EXTRACTION_JS

    screenshots: list[Path] = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(
            viewport={"width": SLIDE_WIDTH_PX, "height": SLIDE_HEIGHT_PX},
        )
        await page.goto(
            html_path.resolve().as_uri(), wait_until="networkidle", timeout=30_000,
        )
        await page.wait_for_timeout(SCRIPT_SETTLE_MS)
        slides = await page.evaluate(EXTRACTION_JS)
        for i in range(len(slides)):
            await page.evaluate("idx => window.__htmlToPptx.prepareSlide(idx)", i)
            out_path = out_dir / f"html_slide_{i}.png"
            await page.screenshot(path=str(out_path), full_page=False)
            await page.evaluate("() => window.__htmlToPptx.finishSlide()")
            screenshots.append(out_path)
        await browser.close()
    return screenshots


def pptx_to_pngs(pptx_path: Path, out_dir: Path) -> list[Path]:
    """Convert PPTX -> PDF via LibreOffice, then PDF pages -> PNGs via PyMuPDF."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        logger.error(
            "LibreOffice not found. Install it or add 'soffice' to PATH.\n"
            "  macOS:  brew install --cask libreoffice\n"
            "  Ubuntu: sudo apt install libreoffice\n"
        )
        return []

    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(pptx_path)],
        capture_output=True, text=True, timeout=60,
    )
    if result.returncode != 0:
        logger.error("LibreOffice PPTX->PDF failed:\n%s", result.stderr)
        return []

    pdf_path = out_dir / pptx_path.with_suffix(".pdf").name
    if not pdf_path.exists():
        logger.error("Expected PDF not found: %s", pdf_path)
        return []

    try:
        import pymupdf
    except ImportError:
        logger.error(
            "PyMuPDF not installed. Run: pip install html-to-pptx[compare]"
        )
        return []

    pngs: list[Path] = []
    doc = pymupdf.open(str(pdf_path))
    for i in range(len(doc)):
        page = doc[i]
        mat = pymupdf.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        png_path = out_dir / f"pptx_slide_{i}.png"
        pix.save(str(png_path))
        pngs.append(png_path)
    doc.close()

    return pngs


def create_comparison(
    html_pngs: list[Path],
    pptx_pngs: list[Path],
    out_dir: Path,
) -> list[Path]:
    """Create side-by-side comparison images (HTML left, PPTX right)."""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        logger.error("Pillow not installed. Run: pip install html-to-pptx[compare]")
        return []

    comparisons: list[Path] = []
    count = min(len(html_pngs), len(pptx_pngs))

    for i in range(count):
        html_img = Image.open(str(html_pngs[i])).convert("RGB")
        pptx_img = Image.open(str(pptx_pngs[i])).convert("RGB")

        target_h = max(html_img.height, pptx_img.height)
        if html_img.height != target_h:
            scale = target_h / html_img.height
            html_img = html_img.resize(
                (int(html_img.width * scale), target_h), Image.LANCZOS,
            )
        if pptx_img.height != target_h:
            scale = target_h / pptx_img.height
            pptx_img = pptx_img.resize(
                (int(pptx_img.width * scale), target_h), Image.LANCZOS,
            )

        GAP_PX = 4
        LABEL_HEIGHT_PX = 32
        total_w = html_img.width + GAP_PX + pptx_img.width
        total_h = target_h + LABEL_HEIGHT_PX

        canvas = Image.new("RGB", (total_w, total_h), (240, 240, 240))

        draw = ImageDraw.Draw(canvas)
        try:
            font = ImageFont.truetype("Arial", 18)
        except OSError:
            font = ImageFont.load_default()
        draw.text((8, 6), f"HTML (slide {i})", fill=(80, 80, 80), font=font)
        draw.text(
            (html_img.width + GAP_PX + 8, 6),
            f"PPTX (slide {i})", fill=(80, 80, 80), font=font,
        )

        canvas.paste(html_img, (0, LABEL_HEIGHT_PX))
        canvas.paste(pptx_img, (html_img.width + GAP_PX, LABEL_HEIGHT_PX))

        out_path = out_dir / f"compare_slide_{i}.png"
        canvas.save(str(out_path))
        comparisons.append(out_path)

    return comparisons


async def process_file(
    html_path: Path, output_dir: Path, *, only_convert: bool = False,
) -> None:
    """Full pipeline for one HTML file."""
    from html_to_pptx.converter import convert

    name = html_path.stem
    sample_out = output_dir / name
    if sample_out.exists():
        shutil.rmtree(sample_out)
    sample_out.mkdir(parents=True)

    logger.info("\n=== %s ===", html_path.name)

    pptx_path = sample_out / "output.pptx"
    logger.info("  Converting HTML -> PPTX...")
    await convert(str(html_path), str(pptx_path))

    if only_convert:
        logger.info("  Done (--only-convert). PPTX: %s", pptx_path)
        return

    logger.info("  Screenshotting HTML slides...")
    html_pngs = await screenshot_html_slides(html_path, sample_out)
    logger.info("  Got %d HTML screenshots", len(html_pngs))

    logger.info("  Converting PPTX -> PNGs via LibreOffice...")
    pptx_pngs = pptx_to_pngs(pptx_path, sample_out)
    logger.info("  Got %d PPTX screenshots", len(pptx_pngs))

    if html_pngs and pptx_pngs:
        logger.info("  Generating comparisons...")
        comparisons = create_comparison(html_pngs, pptx_pngs, sample_out)
        logger.info("  Created %d comparison images", len(comparisons))
    else:
        logger.warning("  Skipping comparison (missing screenshots)")

    logger.info("  Output: %s/", sample_out)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="html-to-pptx-compare",
        description=(
            "Convert HTML slide decks to PPTX and generate side-by-side "
            "visual comparisons to verify conversion quality."
        ),
    )
    parser.add_argument(
        "files", nargs="+", help="HTML file(s) to process",
    )
    parser.add_argument(
        "-o", "--output", default="output",
        help="Output directory (default: ./output)",
    )
    parser.add_argument(
        "--only-convert", action="store_true",
        help="Only convert to PPTX, skip visual comparison",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable debug logging",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    html_files = [Path(f) for f in args.files]
    for f in html_files:
        if not f.exists():
            logger.error("File not found: %s", f)
            sys.exit(1)

    async def run_all() -> None:
        for html_path in html_files:
            await process_file(
                html_path, output_dir, only_convert=args.only_convert,
            )
        logger.info("\nDone. Results in %s/", output_dir)

    asyncio.run(run_all())


if __name__ == "__main__":
    main()
