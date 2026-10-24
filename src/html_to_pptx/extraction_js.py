"""In-page DOM measurement script.

``EXTRACTION_JS`` is evaluated in the loaded deck. For each slide it:

1. Reveals the slide the way the deck itself does (its own ``.active`` rule), so
   the authored layout mode survives — forcing ``display:flex`` made
   ``grid-template-columns`` inert and collapsed grid slides.
2. Settles CSS animations/transitions, which a reveal (re)starts from their
   first keyframe (often ``opacity:0``) and would otherwise drop content.
3. Clears ancestor transforms so ``getBoundingClientRect`` returns layout-space
   coordinates, then measures every visible element.
4. Marks paint that PPTX cannot express natively (SVG, canvas, non-PNG/JPEG
   images, radial/conic/layered/url() backgrounds, filters, clip-paths, masks,
   pseudo-element content) for the browser raster pass in ``raster.py``.

It installs ``window.__htmlToPptx`` so that pass can re-show a slide with the
identical geometry and capture one element's pixels in isolation.

Pure containers (no text, no image, no visible paint, single child) are
collapsed to reduce nesting without losing visual information.
"""

from __future__ import annotations

EXTRACTION_JS = r"""
() => {
    const SKIP_TAGS = new Set(['script', 'style', 'link', 'meta', 'template', 'noscript']);
    const INLINE_TAGS = new Set([
        'span','strong','em','b','i','a','code','mark','sub','sup',
        'small','u','s','del','abbr','cite','q','time','var','kbd',
    ]);
    // Formats python-pptx embeds directly; anything else is captured by the browser.
    const NATIVE_IMAGE_RE = /^data:image\/(png|jpe?g|gif|bmp|tiff?);base64,/i;
    const SUBTREE_RASTER_TAGS = new Set(['svg', 'canvas', 'video', 'iframe']);
    const PSEUDOS = ['::before', '::after'];
    const MAX_DEPTH = 15;
    // A Gaussian blur of radius r (sigma = r in CSS) visibly spreads ~3 sigma.
    const BLUR_EXTENT_PER_RADIUS = 3;
    const MAX_PAINT_OVERFLOW_PX = 400;
    const DECK_CONTAINER_SELECTOR = '#deck, .deck, .slides, .reveal, .slides-container, .deck-container';
    // CSS collapsible whitespace. Unlike \s (and String.trim) it excludes
    // no-break spaces, which CSS never collapses.
    const CSS_SPACE_RUN = /[ \t\n\r\f]+/g;
    // `clip` on the screen-reader-only pattern; it paints nothing at all.
    const ZERO_CLIP = 'rect(0px, 0px, 0px, 0px)';
    const SUBTREE_ATTR = 'data-pptx-isolate';
    const PAINT_ATTR = 'data-pptx-isolate-paint';
    // While one element carries an isolation attribute everything else stops
    // painting (layout is untouched, so the measured box still matches) and the
    // page background drops out, leaving only that element's pixels with real
    // alpha. Subtree mode keeps the element's descendants (an SVG chart); paint
    // mode keeps only the element's own background, border, shadow and
    // pseudo-elements, because its text and children are exported natively.
    // Opacity is forced to 1 because render_pptx re-applies measured opacity.
    const ISOLATION_CSS = `
        html:has([${SUBTREE_ATTR}]) *:not([${SUBTREE_ATTR}], [${SUBTREE_ATTR}] *),
        html:has([${PAINT_ATTR}]) *:not([${PAINT_ATTR}]) { visibility: hidden !important; }
        [${SUBTREE_ATTR}], [${PAINT_ATTR}] { visibility: visible !important; }
        :is([${SUBTREE_ATTR}], [${PAINT_ATTR}], :has([${SUBTREE_ATTR}], [${PAINT_ATTR}])) {
            opacity: 1 !important;
        }
        html:has([${SUBTREE_ATTR}], [${PAINT_ATTR}]),
        html:has([${SUBTREE_ATTR}], [${PAINT_ATTR}]) body { background: transparent !important; }
        [${PAINT_ATTR}] {
            color: transparent !important;
            -webkit-text-fill-color: transparent !important;
            -webkit-text-stroke-color: transparent !important;
            text-shadow: none !important;
            text-decoration-color: transparent !important;
        }
    `;

    let rasterCounter = 0;

    function findSlides() {
        let list = Array.from(document.querySelectorAll('.slide'));
        if (list.length) return list;
        list = Array.from(document.querySelectorAll('.slides > section, .reveal .slides > section'));
        if (list.length) return list;
        list = Array.from(document.querySelectorAll('section'));
        if (list.length) return list;
        const deck = document.getElementById('deck')
            || document.querySelector('.deck, .slides, .slides-container, .deck-container');
        if (deck) return Array.from(deck.children).filter(c => !isSkipped(c));
        return [];
    }

    function settleAnimations() {
        // finish() jumps a finite animation/transition to its end state; an
        // infinite one cannot finish and is cancelled back to the authored style.
        for (const animation of document.getAnimations()) {
            if (animation.effect && animation.effect.getComputedTiming().endTime === Infinity) {
                animation.cancel();
            } else {
                animation.finish();
            }
        }
    }

    function showSlide(slides, idx) {
        document.querySelectorAll(DECK_CONTAINER_SELECTOR).forEach(c => {
            c.style.setProperty('transition', 'none', 'important');
        });
        slides.forEach((s, i) => {
            if (i === idx) {
                s.style.removeProperty('display');
                s.style.setProperty('opacity', '1', 'important');
                s.style.setProperty('visibility', 'visible', 'important');
                s.classList.add('active', 'present', 'current');
                s.classList.remove('past', 'future');
                s.removeAttribute('hidden');
                // Last resort only: decks that hide slides without an `.active`
                // rule would otherwise stay invisible.
                if (getComputedStyle(s).display === 'none') {
                    s.style.setProperty('display', 'flex', 'important');
                }
            } else {
                s.style.setProperty('display', 'none', 'important');
                s.classList.remove('active', 'present', 'current');
            }
        });
        settleAnimations();
    }

    function clearAncestorTransforms(el) {
        const saved = [];
        let current = el;
        while (current && current !== document.documentElement) {
            const computed = getComputedStyle(current);
            if (computed.transform && computed.transform !== 'none') {
                saved.push({el: current, original: current.style.transform});
                current.style.transform = 'none';
            }
            current = current.parentElement;
        }
        return saved;
    }

    function restoreTransforms(saved) {
        for (const item of saved) {
            item.el.style.transform = item.original;
        }
    }

    function isSkipped(el) {
        const tag = el.tagName.toLowerCase();
        return SKIP_TAGS.has(tag) || (tag === 'aside' && el.classList.contains('notes'));
    }

    function splitTopLevel(value) {
        const parts = [];
        let depth = 0, start = 0;
        for (let i = 0; i < value.length; i++) {
            const ch = value[i];
            if (ch === '(') depth++;
            else if (ch === ')') depth--;
            else if (ch === ',' && depth === 0) { parts.push(value.slice(start, i)); start = i + 1; }
        }
        parts.push(value.slice(start));
        return parts.map(p => p.trim()).filter(Boolean);
    }

    // Only a single linear-gradient maps to OOXML; everything else is captured.
    function hasComplexBackground(bg) {
        if (!bg || bg === 'none') return false;
        return splitTopLevel(bg).length > 1 || !bg.startsWith('linear-gradient(');
    }

    function maskOf(style) {
        return style.maskImage || style.webkitMaskImage || 'none';
    }

    function transformsPaint(style) {
        return style.filter !== 'none' || style.clipPath !== 'none' || maskOf(style) !== 'none';
    }

    function pxSum(text) {
        let sum = 0;
        for (const m of text.matchAll(/(-?[\d.]+)px/g)) sum += Math.abs(parseFloat(m[1]));
        return sum;
    }

    // How far blur, drop-shadow and outer box-shadow paint beyond the border box.
    function paintOverflowPx(style) {
        let extent = 0;
        const filter = style.filter || 'none';
        for (const m of filter.matchAll(/blur\(([\d.]+)px\)/g)) {
            extent = Math.max(extent, BLUR_EXTENT_PER_RADIUS * parseFloat(m[1]));
        }
        if (filter.includes('drop-shadow')) extent = Math.max(extent, pxSum(filter));
        if (style.boxShadow && style.boxShadow !== 'none') {
            for (const layer of splitTopLevel(style.boxShadow)) {
                if (!layer.includes('inset')) extent = Math.max(extent, pxSum(layer));
            }
        }
        return Math.min(extent, MAX_PAINT_OVERFLOW_PX);
    }

    function pseudoIsRendered(ps) {
        return ps.content && ps.content !== 'none' && ps.content !== 'normal'
            && ps.display !== 'none' && ps.visibility !== 'hidden';
    }

    // Pseudo content glyphs (quote marks, counters, icons) and non-native paint
    // have no editable equivalent, so the host's own paint is captured instead.
    function pseudoNeedsRaster(ps) {
        if (!pseudoIsRendered(ps)) return false;
        const hasContent = ps.content !== '""' && ps.content !== "''";
        return hasContent || hasComplexBackground(ps.backgroundImage)
            || transformsPaint(ps) || ps.boxShadow !== 'none';
    }

    function rasterModeOf(el, tag, style, src, hasOwnPaint, clipsBackgroundToText) {
        if (SUBTREE_RASTER_TAGS.has(tag)) return 'subtree';
        if (tag === 'img') return NATIVE_IMAGE_RE.test(src) && !transformsPaint(style) ? null : 'subtree';
        const hasText = !!el.textContent.trim();
        if (transformsPaint(style) && !hasText) return 'subtree';
        if (clipsBackgroundToText) return null;
        if (hasComplexBackground(style.backgroundImage)) return 'paint';
        if (transformsPaint(style) && hasOwnPaint) return 'paint';
        if (PSEUDOS.some(p => pseudoNeedsRaster(getComputedStyle(el, p)))) return 'paint';
        return null;
    }

    function markForRaster(el) {
        const id = 'pptx-raster-' + (rasterCounter++);
        el.setAttribute('data-pptx-id', id);
        return id;
    }

    function letterSpacingPx(style) {
        return style.letterSpacing === 'normal' ? 0 : (parseFloat(style.letterSpacing) || 0);
    }

    // Width the element's content actually occupies, as opposed to its box.
    function textExtentPx(el) {
        const range = document.createRange();
        range.selectNodeContents(el);
        return range.getBoundingClientRect().width;
    }

    function runStyle(cs) {
        return {
            color: cs.color,
            fontSize: parseFloat(cs.fontSize),
            fontFamily: cs.fontFamily,
            fontWeight: cs.fontWeight,
            fontStyle: cs.fontStyle,
            textTransform: cs.textTransform,
            letterSpacing: letterSpacingPx(cs),
            textDecoration: cs.textDecorationLine,
        };
    }

    function readBorders(style) {
        const sides = [];
        for (const side of ['top', 'right', 'bottom', 'left']) {
            const cap = side[0].toUpperCase() + side.slice(1);
            const width = parseFloat(style['border' + cap + 'Width']) || 0;
            const lineStyle = style['border' + cap + 'Style'];
            if (width > 0 && lineStyle !== 'none' && lineStyle !== 'hidden') {
                sides.push({side, width, color: style['border' + cap + 'Color'], style: lineStyle});
            }
        }
        const same = (a, b) => a.width === b.width && a.color === b.color && a.style === b.style;
        const uniform = sides.length === 4 && sides.every(s => same(s, sides[0]));
        const left = sides.find(s => s.side === 'left');
        return {
            any: sides.length > 0,
            borderColor: uniform ? sides[0].color : null,
            borderWidth: uniform ? sides[0].width : 0,
            borderStyle: uniform ? sides[0].style : null,
            borderLeftColor: !uniform && left ? left.color : null,
            borderLeftWidth: !uniform && left ? left.width : 0,
            borderLeftStyle: !uniform && left ? left.style : null,
            borderSides: uniform ? [] : sides,
        };
    }

    // 'preserve' keeps spaces and newlines (pre, pre-wrap, break-spaces),
    // 'breaks' keeps only newlines (pre-line), 'collapse' keeps neither.
    function whitespaceMode(style) {
        if (style.whiteSpace === 'pre-line') return 'breaks';
        if (style.whiteSpace.startsWith('pre') || style.whiteSpace === 'break-spaces') return 'preserve';
        return 'collapse';
    }

    // Text as the browser renders it, before trimming at the element's edges.
    function renderedText(text, mode) {
        if (mode === 'preserve') return text;
        if (mode === 'breaks') return text.replace(/[ \t]*\n[ \t]*/g, '\n').replace(/[ \t]+/g, ' ');
        return text.replace(CSS_SPACE_RUN, ' ');
    }

    function getDirectText(el, mode) {
        const parts = [];
        for (const node of el.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) parts.push(renderedText(node.textContent, mode));
        }
        if (mode === 'preserve') return parts.join('');
        return parts.map(p => p.replace(/^ | $/g, '')).filter(Boolean).join(' ');
    }

    function collectInlineRuns(el) {
        const runs = [];
        const parentStyle = getComputedStyle(el);
        const parentRun = runStyle(parentStyle);
        const parentMode = whitespaceMode(parentStyle);
        const lineBreak = () => ({ ...parentRun, text: '\n', textTransform: 'none' });
        // Preserved newlines become '\n' runs, which render as paragraph breaks.
        const pushText = (text, run) => {
            text.split('\n').forEach((segment, i) => {
                if (i > 0) runs.push(lineBreak());
                if (segment) runs.push({ ...run, text: segment });
            });
        };
        const nodes = el.childNodes;
        for (let ni = 0; ni < nodes.length; ni++) {
            const node = nodes[ni];
            if (node.nodeType === Node.TEXT_NODE) {
                // Collapsed spaces survive between inline siblings but not at the
                // element's edges.
                let t = renderedText(node.textContent, parentMode);
                if (parentMode !== 'preserve') {
                    if (ni === 0) t = t.replace(/^ /, '');
                    if (ni === nodes.length - 1) t = t.replace(/ $/, '');
                }
                pushText(t, parentRun);
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                const tag = node.tagName.toLowerCase();
                if (SKIP_TAGS.has(tag)) continue;
                const cs = getComputedStyle(node);
                if (cs.display === 'none' || cs.visibility === 'hidden') continue;
                if (tag === 'br') {
                    runs.push(lineBreak());
                    continue;
                }
                const isInline = cs.display.startsWith('inline') || INLINE_TAGS.has(tag);
                if (isInline) {
                    const mode = whitespaceMode(cs);
                    let text = renderedText(node.textContent, mode);
                    if (mode !== 'preserve') text = text.replace(/^ | $/g, '');
                    if (text) {
                        const childBgImage = cs.backgroundImage !== 'none' ? cs.backgroundImage : null;
                        const childFillColor = cs.webkitTextFillColor || '';
                        const childIsGradientText = (childFillColor === 'transparent' || childFillColor === 'rgba(0, 0, 0, 0)') && childBgImage && childBgImage.includes('gradient');
                        pushText(text, {
                            ...runStyle(cs),
                            href: tag === 'a' ? node.getAttribute('href') : null,
                            isGradientText: childIsGradientText,
                            backgroundImage: childIsGradientText ? childBgImage : null,
                        });
                    }
                }
            }
        }
        return runs;
    }

    function hasChildElementText(el) {
        for (const child of el.children) {
            if (child.textContent && child.textContent.trim()) return true;
        }
        return false;
    }

    // Text-free ::before/::after decorations with plain paint (accent bars,
    // bullet dots) become native shapes; anything richer is captured with the host.
    function measurePseudo(el, pseudo, relX, relY, rect) {
        const ps = getComputedStyle(el, pseudo);
        if (!pseudoIsRendered(ps)) return null;

        const psBg = ps.backgroundColor !== 'rgba(0, 0, 0, 0)' && ps.backgroundColor !== 'transparent';
        const pw = parseFloat(ps.width) || 0;
        const ph = parseFloat(ps.height) || 0;
        if (pw <= 0 || ph <= 0) return null;

        let px = relX, py = relY;
        const pt = parseFloat(ps.top); const pl = parseFloat(ps.left);
        const pr = parseFloat(ps.right); const pb = parseFloat(ps.bottom);
        if (ps.position === 'absolute') {
            if (!isNaN(pl)) px = relX + pl;
            else if (!isNaN(pr)) px = relX + rect.width - pw - pr;
            if (!isNaN(pt)) py = relY + pt;
            else if (!isNaN(pb)) py = relY + rect.height - ph - pb;
        }

        return {
            tag: '_pseudo',
            x: px, y: py, width: pw, height: ph,
            text: '',
            color: ps.color,
            backgroundColor: psBg ? ps.backgroundColor : null,
            backgroundImage: ps.backgroundImage !== 'none' ? ps.backgroundImage : null,
            fontSize: 0, fontFamily: '', fontWeight: '400', fontStyle: 'normal',
            textAlign: 'left', lineHeight: 'normal', direction: 'ltr',
            opacity: parseFloat(ps.opacity), borderRadius: ps.borderRadius,
            ...readBorders(ps),
            boxShadow: null,
            textTransform: 'none', isImage: false, src: null, href: null,
            markerColor: null, children: [], isGradientText: false,
        };
    }

    function measureElement(el, slideRect, depth) {
        if (depth > MAX_DEPTH) return null;

        const style = getComputedStyle(el);
        if (style.display === 'none' || style.visibility === 'hidden') return null;
        if (parseFloat(style.opacity) === 0) return null;

        // A CSS rotate makes getBoundingClientRect return the (larger) enclosing
        // box. Detect the angle, then measure the element with its own transform
        // neutralized so we get the true unrotated box; PPTX re-applies the angle
        // about the shape center. (Restored immediately to avoid side effects.)
        let ownRotation = 0;
        const _tf = style.transform;
        if (_tf && _tf.indexOf('matrix') === 0) {
            const mm = _tf.match(/matrix\(([^)]+)\)/);
            if (mm) {
                const p = mm[1].split(',').map(parseFloat);
                const ang = Math.atan2(p[1], p[0]) * 180 / Math.PI;
                if (Math.abs(ang) > 0.5) ownRotation = ang;
            }
        }
        let _savedT = null;
        if (ownRotation) { _savedT = el.style.transform; el.style.transform = 'none'; el.getBoundingClientRect(); }

        const rect = el.getBoundingClientRect();
        if (ownRotation) { el.style.transform = _savedT; }
        const relX = rect.left - slideRect.left;
        const relY = rect.top - slideRect.top;

        if (rect.width <= 0 || rect.height <= 0) return null;
        if (relX + rect.width <= 0 || relY + rect.height <= 0) return null;
        if (relX >= slideRect.width || relY >= slideRect.height) return null;

        // Visually-hidden text (the sr-only pattern) stays in the DOM but never
        // paints: a zero `clip`, or overflow clipped to a box at most a pixel wide.
        const isPositioned = style.position === 'absolute' || style.position === 'fixed';
        if (isPositioned && style.clip === ZERO_CLIP) return null;
        const clipsContent = (style.overflowX !== 'visible' && rect.width <= 1)
            || (style.overflowY !== 'visible' && rect.height <= 1);

        let directText = clipsContent ? '' : getDirectText(el, whitespaceMode(style));
        const tag = el.tagName.toLowerCase();

        let markerColor = null;
        let markerPrefix = '';
        if (tag === 'li') {
            const parentTag = el.parentElement ? el.parentElement.tagName.toLowerCase() : '';
            const listStyle = style.listStyleType;
            if (parentTag === 'ol') {
                const index = Array.from(el.parentElement.children).indexOf(el) + 1;
                markerPrefix = index + '. ';
                directText = markerPrefix + directText;
            } else if (listStyle !== 'none') {
                markerPrefix = '\u2022 ';
                directText = markerPrefix + directText;
            }
            const ms = getComputedStyle(el, '::marker');
            if (ms && ms.color) markerColor = ms.color;
        }

        const isImg = tag === 'img';
        const isSvg = tag === 'svg';
        const src = isImg ? (el.currentSrc || el.getAttribute('src') || '') : '';
        const hasVisibleBg = style.backgroundColor !== 'rgba(0, 0, 0, 0)' &&
                             style.backgroundColor !== 'transparent';
        const borders = readBorders(style);
        const boxShadow = style.boxShadow !== 'none' ? style.boxShadow : null;

        const bgImage = style.backgroundImage !== 'none' ? style.backgroundImage : null;
        const textFillColor = style.webkitTextFillColor || style.WebkitTextFillColor || '';
        const isFillTransparent = textFillColor === 'transparent' || textFillColor === 'rgba(0, 0, 0, 0)';
        const isGradientText = isFillTransparent && bgImage && bgImage.includes('gradient');
        const clipsBackgroundToText = style.backgroundClip === 'text' || style.webkitBackgroundClip === 'text';
        const hasOwnPaint = hasVisibleBg || !!bgImage || borders.any || !!boxShadow;
        const rasterMode = rasterModeOf(
            el, tag, style, src, hasOwnPaint, isGradientText || clipsBackgroundToText,
        );

        const data = {
            tag: tag,
            x: relX,
            y: relY,
            width: rect.width,
            height: rect.height,
            text: rasterMode === 'subtree' ? '' : directText,
            color: style.color,
            backgroundColor: hasVisibleBg ? style.backgroundColor : null,
            fontSize: parseFloat(style.fontSize),
            fontFamily: style.fontFamily,
            fontWeight: style.fontWeight,
            fontStyle: style.fontStyle,
            textAlign: style.textAlign,
            lineHeight: style.lineHeight,
            letterSpacing: letterSpacingPx(style),
            textDecoration: style.textDecorationLine,
            textWidth: directText ? textExtentPx(el) : 0,
            direction: style.direction,
            position: style.position,
            display: style.display,
            justifyContent: style.justifyContent,
            alignItems: style.alignItems,
            writingMode: style.writingMode,
            rotation: ownRotation,
            opacity: parseFloat(style.opacity),
            paddingLeft: parseFloat(style.paddingLeft) || 0,
            paddingRight: parseFloat(style.paddingRight) || 0,
            paddingTop: parseFloat(style.paddingTop) || 0,
            paddingBottom: parseFloat(style.paddingBottom) || 0,
            objectFit: isImg ? style.objectFit : null,
            naturalWidth: isImg ? (el.naturalWidth || 0) : 0,
            naturalHeight: isImg ? (el.naturalHeight || 0) : 0,
            borderRadius: style.borderRadius,
            borderColor: borders.borderColor,
            borderWidth: borders.borderWidth,
            borderStyle: borders.borderStyle,
            borderLeftColor: borders.borderLeftColor,
            borderLeftWidth: borders.borderLeftWidth,
            borderLeftStyle: borders.borderLeftStyle,
            borderSides: borders.borderSides,
            boxShadow: boxShadow,
            textTransform: style.textTransform,
            backgroundImage: bgImage,
            isGradientText: isGradientText,
            isImage: isImg,
            isSvg: isSvg,
            src: isImg ? src : null,
            href: tag === 'a' ? el.getAttribute('href') : null,
            markerColor: markerColor,
            rasterMode: rasterMode,
            rasterId: rasterMode ? markForRaster(el) : null,
            children: []
        };

        // A captured subtree is one picture: its descendants are already in it.
        if (rasterMode === 'subtree') return data;

        // ::before is painted before the element's children and ::after after
        // them. A paint-raster host already carries its pseudo-elements.
        const pseudoData = rasterMode === 'paint'
            ? [null, null]
            : PSEUDOS.map(p => measurePseudo(el, p, relX, relY, rect));
        if (pseudoData[0]) data.children.push(pseudoData[0]);
        for (const child of clipsContent ? [] : el.children) {
            if (isSkipped(child)) continue;
            const childData = measureElement(child, slideRect, depth + 1);
            if (childData) data.children.push(childData);
        }
        if (pseudoData[1]) data.children.push(pseudoData[1]);

        // Preserved newlines need paragraph breaks, which only runs carry.
        if (directText && (hasChildElementText(el) || directText.includes('\n'))) {
            const runs = collectInlineRuns(el);
            if (runs.length > 0) {
                if (markerPrefix) {
                    runs[0].text = markerPrefix + runs[0].text;
                }
                data.inlineRuns = runs;
                data.text = '';
            }
        }

        const isContainer = !data.text && !isImg && !isSvg && !hasVisibleBg && !borders.any &&
                           data.backgroundImage === null && !data.inlineRuns &&
                           !rasterMode && !boxShadow;
        if (isContainer && data.children.length === 1 && depth > 0) {
            const child = data.children[0];
            // A pure wrapper is collapsed away, but its visual effects must be
            // folded into the surviving child, or they are silently lost:
            //  - opacity (e.g. a faint hero-image wrapper at opacity:0.18)
            //  - border-radius + overflow:hidden clipping (rounded image frames)
            if (data.opacity < 1) {
                child.opacity = (child.opacity == null ? 1 : child.opacity) * data.opacity;
            }
            const childHasRadius = child.borderRadius && child.borderRadius !== '0px'
                && child.borderRadius !== '0';
            if (!childHasRadius && data.borderRadius && data.borderRadius !== '0px'
                && data.borderRadius !== '0') {
                child.borderRadius = data.borderRadius;
            }
            return child;
        }

        return data;
    }

    function slideNotes(slide) {
        const attr = slide.getAttribute('data-pptx-notes');
        if (attr) return attr.trim();
        const aside = slide.querySelector('aside.notes');
        return aside ? aside.textContent.replace(/\s+/g, ' ').trim() : '';
    }

    function measureSlide(slide, slideIndex) {
        const slideRect = slide.getBoundingClientRect();
        const slideStyle = getComputedStyle(slide);
        const relX = 0, relY = 0;
        const needsPaint = hasComplexBackground(slideStyle.backgroundImage)
            || PSEUDOS.some(p => pseudoNeedsRaster(getComputedStyle(slide, p)));
        const slideData = {
            index: slideIndex,
            width: slideRect.width,
            height: slideRect.height,
            backgroundColor: slideStyle.backgroundColor,
            backgroundImage: slideStyle.backgroundImage !== 'none' ? slideStyle.backgroundImage : null,
            notes: slideNotes(slide),
            rasterMode: needsPaint ? 'paint' : null,
            rasterId: needsPaint ? markForRaster(slide) : null,
            elements: []
        };

        const pseudoData = needsPaint
            ? [null, null]
            : PSEUDOS.map(p => measurePseudo(slide, p, relX, relY, slideRect));
        if (pseudoData[0]) slideData.elements.push(pseudoData[0]);
        for (const child of slide.children) {
            if (isSkipped(child)) continue;
            const measured = measureElement(child, slideRect, 0);
            if (measured) slideData.elements.push(measured);
        }
        if (pseudoData[1]) slideData.elements.push(pseudoData[1]);
        return slideData;
    }

    const slides = findSlides();
    const results = slides.map((slide, slideIndex) => {
        showSlide(slides, slideIndex);
        const savedTransforms = clearAncestorTransforms(slide);
        slide.offsetHeight;
        const slideData = measureSlide(slide, slideIndex);
        restoreTransforms(savedTransforms);
        return slideData;
    });

    const isolationStyle = document.createElement('style');
    isolationStyle.textContent = ISOLATION_CSS;
    const pseudoColorStyle = document.createElement('style');
    let activeSlide = null;
    let activeTransforms = [];
    let capture = null;

    // Paint mode makes the host's text transparent; pseudo-elements inherit that,
    // so pin their authored colors (read before isolating) to keep their glyphs.
    function pseudoColorRules(el) {
        return PSEUDOS.map(p => {
            const ps = getComputedStyle(el, p);
            return `[${PAINT_ATTR}]${p} { color: ${ps.color} !important; `
                + `-webkit-text-fill-color: ${ps.webkitTextFillColor} !important; }`;
        }).join('\n');
    }

    window.__htmlToPptx = {
        prepareSlide(idx) {
            if (!isolationStyle.isConnected) document.head.append(isolationStyle, pseudoColorStyle);
            restoreTransforms(activeTransforms);
            activeSlide = slides[idx];
            showSlide(slides, idx);
            activeTransforms = clearAncestorTransforms(activeSlide);
            activeSlide.scrollIntoView({block: 'start', inline: 'start'});
        },
        finishSlide() {
            restoreTransforms(activeTransforms);
            activeTransforms = [];
            activeSlide = null;
        },
        // Returns the viewport clip to screenshot and the same box relative to the
        // slide, or null when the element is gone or entirely off-slide.
        beginCapture(args) {
            const el = document.querySelector('[data-pptx-id="' + args.id + '"]');
            if (!el) return null;
            const attr = args.mode === 'paint' ? PAINT_ATTR : SUBTREE_ATTR;
            // Slide reveals (ours and callers') set inline `visibility: visible
            // !important`, which beats the isolation stylesheet, so ancestors
            // are hidden inline for the duration of the capture.
            const ancestors = [];
            for (let a = el.parentElement; a; a = a.parentElement) {
                ancestors.push({
                    el: a,
                    value: a.style.getPropertyValue('visibility'),
                    priority: a.style.getPropertyPriority('visibility'),
                });
                a.style.setProperty('visibility', 'hidden', 'important');
            }
            capture = {el, attr, ancestors, transform: el.style.transform};
            if (args.neutralizeTransform) el.style.transform = 'none';
            if (args.mode === 'paint') pseudoColorStyle.textContent = pseudoColorRules(el);
            el.setAttribute(attr, '');

            const pad = paintOverflowPx(getComputedStyle(el));
            const r = el.getBoundingClientRect();
            const s = activeSlide.getBoundingClientRect();
            const left = Math.max(r.left - pad, s.left, 0);
            const top = Math.max(r.top - pad, s.top, 0);
            const right = Math.min(r.right + pad, s.right, window.innerWidth);
            const bottom = Math.min(r.bottom + pad, s.bottom, window.innerHeight);
            if (right - left < 1 || bottom - top < 1) return null;
            const size = {width: right - left, height: bottom - top};
            return {
                clip: {x: left, y: top, ...size},
                box: {x: left - s.left, y: top - s.top, ...size},
                padded: pad > 0,
            };
        },
        endCapture() {
            if (!capture) return;
            capture.el.removeAttribute(capture.attr);
            capture.el.style.transform = capture.transform;
            for (const a of capture.ancestors) {
                a.el.style.setProperty('visibility', a.value, a.priority);
            }
            pseudoColorStyle.textContent = '';
            capture = null;
        },
    };

    return results;
}
"""
