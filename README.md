# html-to-pptx

Convert HTML slide decks into editable PowerPoint files.

Unlike screenshot-based approaches that produce flat images, `html-to-pptx` measures every DOM element in a headless browser and maps it to a native PPTX shape — text boxes you can edit, images you can resize, backgrounds you can restyle. The output is a real presentation, not a picture of one.

## Install

```bash
pip install html-to-pptx
python -m playwright install chromium
```

## Quick start

### CLI

```bash
html-to-pptx deck.html deck.pptx
```

### Python

```python
import asyncio
from html_to_pptx import convert

asyncio.run(convert("deck.html", "deck.pptx"))
```

### Two-stage API

For more control, separate the measurement and rendering steps:

```python
import asyncio
from html_to_pptx import extract_measurements, render_pptx

async def main():
    # Stage 1: measure DOM elements in a headless browser
    measurements = await extract_measurements("deck.html")

    # Stage 2: render measurements to python-pptx shapes
    prs = render_pptx(measurements)

    # Modify the presentation before saving
    prs.slides[0].shapes[0].text = "Custom title"
    prs.save("deck.pptx")

asyncio.run(main())
```

### Measuring a page you already control

If you already drive Playwright (custom viewport, auth, injected data), measure the loaded page directly. Fonts are awaited, animations settled, and non-native paint rasterized in the same page:

```python
from html_to_pptx import extract_from_page, render_pptx

measurements = await extract_from_page(page)
render_pptx(measurements).save("deck.pptx")
```

## HTML format

Slides are discovered the way common deck frameworks expose them: `.slide` elements first, then reveal.js-style `.slides > section`, then any `section`, then the children of a `#deck`/`.deck`/`.slides` container. Each slide is revealed by removing inline `display:none` and adding the `active`/`present`/`current` classes, so your own `.slide.active { display: grid }` rules keep working. A 1920×1080 canvas is expected:

```html
<!DOCTYPE html>
<html>
<head>
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { background: #000; overflow: hidden; }
    .slide {
        width: 1920px;
        height: 1080px;
        display: none;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        font-family: sans-serif;
    }
    .slide.active { display: flex; }
</style>
</head>
<body>

<section class="slide active" style="background-color: #1e293b;">
    <h1 style="font-size: 64px; color: #f8fafc;">Hello World</h1>
</section>

<section class="slide" style="background-color: #ffffff;">
    <p style="font-size: 24px; color: #334155;">Second slide</p>
</section>

</body>
</html>
```

See [`examples/demo.html`](examples/demo.html) for a full five-slide deck.

## What gets converted

| HTML feature | PPTX output |
|---|---|
| Text with fonts, colors, weight, style | Editable text boxes with matching font properties |
| **Font substitution** (web fonts like Inter, Playfair, Roboto) | Mapped to a metric-compatible, install-safe font in the **same family** so headings don't reflow |
| **Single-line fit** | One-line headings/pills/buttons never wrap — the point size auto-shrinks to fit the box |
| **CSS `line-height`** | Exact paragraph line spacing (multi-line headings keep their height) |
| **CSS padding** | Reproduced as text-frame insets (bulleted/labelled text aligns correctly) |
| **Text/box/gradient alpha** | Flattened over the correct backdrop (nearest ancestor bg), so translucency renders in every viewer |
| **Translucent image overlays / scrims** | Baked into the underlying image so the photo stays visible |
| **Mixed inline content** (`<p>text <strong>bold</strong> more</p>`) | Multi-run paragraphs with per-run styling |
| **CSS `linear-gradient` backgrounds** | OOXML gradient fills with angle (`deg`/`turn`/`rad`/`grad`), exact stop positions, hard stops and double-position stops |
| **Paint PPTX can't express** (radial/conic/repeating/layered gradients, `background-image: url()`, glows, `filter`, `clip-path`, `mask`, decorative `::before`/`::after`) | Only the element's own paint is screenshotted by Chromium and placed underneath — its **text and children stay native and editable** |
| **`box-shadow`** | Native OOXML outer shadow (blur, offset, spread, alpha); layered shadows use the dominant layer |
| **Gradient text** (`-webkit-background-clip: text`) | Gradient text fill with a solid first-stop fallback |
| **Inline `<svg>`, `<canvas>`, `<video>`, `<iframe>`** | Pixel-exact picture captured in isolation at 2× device scale |
| **Any `<img>`** (WebP, AVIF, SVG, remote, relative URLs) | Native picture; non-PPTX formats are captured or transcoded to PNG |
| **`letter-spacing`, `text-decoration`** | Run character spacing, underline, strikethrough |
| **Per-side borders** (`border-bottom` dividers, table row rules) | Thin bars at the exact side and width |
| **CSS animations/transitions** | Settled to their final state before measuring, so fade-ins aren't captured mid-flight |
| **Speaker notes** (`<aside class="notes">` or `data-pptx-notes`) | PowerPoint speaker notes |
| **HTML tables** (`<table>`) | Cell-accurate text with per-column alignment preserved |
| Background colors | Solid fills (with alpha flattening) |
| Rounded corners (`border-radius`, including `%`) | OOXML adjustment guides on rounded rectangles |
| **Rounded images** (`border-radius` on `<img>` or wrapper) | Rounded-rect geometry swap on picture shapes |
| **`object-fit: cover`** | Centered crop (no distortion) instead of stretching |
| Base64-embedded images (`data:image/...`) | Native picture shapes |
| Borders (with alpha) | Shape outlines with color, width, and opacity |
| **Left/top accent bars** (`border-left`, `::before` strips) | Rounded to match the card's corners (offset under-component) |
| **`transform: rotate()`** | Shape rotation about the element center |
| **`writing-mode: vertical-*`** | Vertical PPTX text body |
| **Flex/grid alignment** (`justify-content`, `align-items`) | Horizontal + vertical text anchoring |
| `text-transform` (uppercase, etc.) | Transformed text content |
| Hyperlinks | Preserved on text runs |
| Ordered/unordered lists (incl. custom bullet dots) | Bullet/number prefixes and decorative markers |
| **CSS `::before`/`::after` pseudo-elements** | Synthetic shapes for decorative accents |
| **`white-space: pre`** (terminals/code) | Preserved line breaks and indentation |
| **RTL text direction** | Correct paragraph alignment |
| **`<br>` tags in inline runs** | Multi-paragraph text frames |
| Element opacity (incl. inherited) | OOXML alpha / picture transparency |

## How it works

```
HTML file
    │
    ▼
┌─────────────────────────┐
│  Headless Chromium       │  ← Playwright loads the HTML at 1920×1080
│  (Playwright)            │
│                          │
│  For each slide:         │
│  • Reveal it, settle     │
│    animations            │
│  • Clear CSS transforms  │
│  • Measure every visible │
│    element recursively:  │
│    position, size, color,│
│    font, text, images    │
│  • Screenshot paint PPTX │
│    can't express, in     │
│    isolation             │
└───────────┬─────────────┘
            │
            ▼
    Measurement tree (JSON)
            │
            ▼
┌─────────────────────────┐
│  python-pptx renderer    │
│                          │
│  For each element:       │
│  • Text leaf → text box  │
│  • Image → picture shape │
│  • Background → rectangle│
│  • Border → outline/shape│
│                          │
│  Coordinates:            │
│  CSS px → inches using   │
│  13.333/1920 ratio       │
└───────────┬─────────────┘
            │
            ▼
      Editable .pptx
```

## Visual comparison tool

Verify conversion quality with side-by-side comparisons:

```bash
pip install html-to-pptx[compare]
html-to-pptx-compare deck.html -o results/
```

This screenshots the original HTML, converts to PPTX, renders the PPTX back to images via LibreOffice, and creates side-by-side comparison PNGs. Requires [LibreOffice](https://www.libreoffice.org/) on PATH.

Output structure:
```
results/deck/
├── html_slide_0.png      # ground truth (HTML screenshot)
├── html_slide_1.png
├── pptx_slide_0.png      # converter output (PPTX rendered via LibreOffice)
├── pptx_slide_1.png
├── compare_slide_0.png   # side-by-side comparison
├── compare_slide_1.png
└── output.pptx           # the generated file
```

## Limitations

- **Rasterized paint is a picture** — radial gradients, glows, `url()` backgrounds and SVG/canvas content look exact but can't be restyled in PowerPoint (text on top of them stays editable)
- **Inset and multi-layer shadows** — OOXML allows one outer shadow per shape; inset layers are dropped
- **Animations and transitions** are not represented in PPTX — each slide shows its final, settled state
- **Font availability** — web fonts are mapped to a metric-compatible font in the same family (serif→Georgia, humanist sans→Segoe UI, geometric sans→Century Gothic, mono→Consolas, …). For pixel-exact display type, embed the font in PowerPoint yourself
- **Complex CSS layouts** (grid, advanced flexbox) are measured as-rendered, but deeply nested layouts may lose some positioning precision

## Development

```bash
git clone https://github.com/Design-Arena/html-to-pptx.git
cd html-to-pptx
pip install -e ".[dev]"
python -m playwright install chromium
pytest -v
```

## License

MIT
