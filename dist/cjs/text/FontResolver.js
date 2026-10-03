"use strict";
/**
 * Resolves font names from PPTX to system fonts.
 * Handles theme fonts (+mj-lt, +mn-lt) and font substitution fallback chains.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.FontResolver = exports.PIXELS_PER_POINT = void 0;
exports.pointsToPixels = pointsToPixels;
exports.setCustomFontFallbacks = setCustomFontFallbacks;
exports.getCustomFontFallbacks = getCustomFontFallbacks;
exports.createFontResolver = createFontResolver;
const skia_canvas_1 = require("skia-canvas");
const theme_js_1 = require("../types/theme.js");
const Logger_js_1 = require("../utils/Logger.js");
const FontManager_js_1 = require("./FontManager.js");
/**
 * Pixels per point at the canvas-native 96 DPI (1pt = 1/72 inch, 1px = 1/96 inch).
 */
exports.PIXELS_PER_POINT = 96 / 72;
/**
 * Converts a size in points to pixels at the canvas-native 96 DPI.
 *
 * @param points Size in points
 * @returns Size in pixels (unscaled)
 */
function pointsToPixels(points) {
    return points * exports.PIXELS_PER_POINT;
}
/**
 * Font substitution fallback chains.
 * When a requested font is unavailable, try these alternatives in order.
 */
const FONT_FALLBACK_CHAINS = {
    // Common Windows fonts
    Calibri: ['Calibri', 'Arial', 'Helvetica', 'sans-serif'],
    'Calibri Light': ['Calibri Light', 'Calibri', 'Arial', 'Helvetica', 'sans-serif'],
    Arial: ['Arial', 'Helvetica', 'sans-serif'],
    'Times New Roman': ['Times New Roman', 'Times', 'Georgia', 'serif'],
    Cambria: ['Cambria', 'Georgia', 'Times New Roman', 'serif'],
    Consolas: ['Consolas', 'Monaco', 'Courier New', 'monospace'],
    'Courier New': ['Courier New', 'Courier', 'monospace'],
    Georgia: ['Georgia', 'Times New Roman', 'serif'],
    Tahoma: ['Tahoma', 'Arial', 'Helvetica', 'sans-serif'],
    Verdana: ['Verdana', 'Arial', 'Helvetica', 'sans-serif'],
    'Trebuchet MS': ['Trebuchet MS', 'Arial', 'Helvetica', 'sans-serif'],
    Impact: ['Impact', 'Arial Black', 'sans-serif'],
    'Comic Sans MS': ['Comic Sans MS', 'cursive'],
    'Segoe UI': ['Segoe UI', 'Arial', 'Helvetica', 'sans-serif'],
    // Common macOS fonts
    Helvetica: ['Helvetica', 'Arial', 'sans-serif'],
    'Helvetica Neue': ['Helvetica Neue', 'Helvetica', 'Arial', 'sans-serif'],
    'San Francisco': ['San Francisco', 'Helvetica Neue', 'Helvetica', 'sans-serif'],
    // Common Japanese fonts
    'MS Gothic': ['MS Gothic', 'Yu Gothic', 'Meiryo', 'Hiragino Kaku Gothic ProN', 'sans-serif'],
    'MS PGothic': [
        'MS PGothic',
        'MS Gothic',
        'Yu Gothic',
        'Meiryo',
        'Hiragino Kaku Gothic ProN',
        'sans-serif',
    ],
    'MS Mincho': ['MS Mincho', 'Yu Mincho', 'Hiragino Mincho ProN', 'serif'],
    'MS PMincho': ['MS PMincho', 'MS Mincho', 'Yu Mincho', 'Hiragino Mincho ProN', 'serif'],
    'Yu Gothic': ['Yu Gothic', 'YuGothic', 'Meiryo', 'Hiragino Kaku Gothic ProN', 'sans-serif'],
    'Yu Mincho': ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'serif'],
    Meiryo: ['Meiryo', 'Yu Gothic', 'Hiragino Kaku Gothic ProN', 'sans-serif'],
    // Common Chinese fonts
    SimSun: ['SimSun', 'Songti SC', 'STSong', 'serif'],
    SimHei: ['SimHei', 'Heiti SC', 'PingFang SC', 'sans-serif'],
    'Microsoft YaHei': ['Microsoft YaHei', 'PingFang SC', 'Heiti SC', 'sans-serif'],
    DengXian: ['DengXian', 'Microsoft YaHei', 'PingFang SC', 'sans-serif'],
    PMingLiU: ['PMingLiU', 'Songti TC', 'serif'],
    'Microsoft JhengHei': ['Microsoft JhengHei', 'PingFang TC', 'Heiti TC', 'sans-serif'],
    // Common Korean fonts
    'Malgun Gothic': ['Malgun Gothic', 'Apple SD Gothic Neo', 'AppleGothic', 'sans-serif'],
    Batang: ['Batang', 'AppleMyungjo', 'serif'],
    Gulim: ['Gulim', 'Apple SD Gothic Neo', 'AppleGothic', 'sans-serif'],
    // Default fallbacks
    'sans-serif': ['Arial', 'Helvetica', 'sans-serif'],
    serif: ['Georgia', 'Times New Roman', 'Times', 'serif'],
    monospace: ['Consolas', 'Monaco', 'Courier New', 'monospace'],
};
/**
 * Default fallback chain for unknown fonts.
 */
const DEFAULT_FALLBACK = ['Arial', 'Helvetica', 'sans-serif'];
/**
 * Generic CSS family keywords: requesting these is not a "missing font"
 * (they are fallback vocabulary, not concrete fonts), so no missing-font
 * warning is emitted for them.
 */
const GENERIC_FAMILIES = new Set([
    'sans-serif',
    'serif',
    'monospace',
    'cursive',
    'fantasy',
    'system-ui',
]);
/**
 * Metric-compatible substitutes, tried immediately after the requested
 * family and before the geometric fallbacks — but only when actually
 * installed. These fonts share glyph metrics with their Windows
 * counterparts, so substituting them preserves line-wrap points:
 * Carlito↔Calibri, Caladea↔Cambria, Liberation Sans↔Arial/Helvetica,
 * Liberation Serif↔Times New Roman.
 */
const METRIC_COMPATIBLE_SUBSTITUTES = {
    Calibri: ['Carlito'],
    'Calibri Light': ['Carlito'],
    Cambria: ['Caladea'],
    Arial: ['Liberation Sans'],
    Helvetica: ['Liberation Sans'],
    'Times New Roman': ['Liberation Serif'],
};
/**
 * Session-level custom fallback chains, merged over the built-in
 * FONT_FALLBACK_CHAINS by every FontResolver instance. Set from
 * PptxRenderOptions.fonts.fallbacks by the renderer entry points.
 *
 * Note: this is process-global state (the resolver is constructed deep
 * inside the text pipeline); concurrent renders with different fallback
 * maps will share the most recently set value.
 */
let sessionFallbacks = {};
/**
 * Sets (replaces) the session custom font fallback chains.
 *
 * @param fallbacks Chains keyed by requested family name; undefined clears
 *   any previously set session fallbacks
 */
function setCustomFontFallbacks(fallbacks) {
    sessionFallbacks = { ...(fallbacks ?? {}) };
}
/**
 * Returns the current session custom font fallback chains.
 */
function getCustomFontFallbacks() {
    return sessionFallbacks;
}
/**
 * Default font availability check: FontManager-registered families
 * (embedded/user fonts) first, then the skia-canvas FontLibrary
 * (system fonts plus registrations).
 */
function defaultFontAvailability(family) {
    if ((0, FontManager_js_1.getRegisteredFontFamilies)().has(family)) {
        return true;
    }
    try {
        return skia_canvas_1.FontLibrary.has(family);
    }
    catch {
        return false;
    }
}
/**
 * Resolves font names from PPTX to canvas-compatible fonts.
 */
class FontResolver {
    logger;
    fontScheme;
    metricsCache = new Map();
    /** Isolated glyph advances per font string, for unkerned (noKern) runs */
    charAdvanceCache = new Map();
    customFallbacks;
    fontAvailability;
    availabilityCache = new Map();
    warnings;
    /** Families already reported as missing (one warning per family) */
    warnedMissingFamilies = new Set();
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'FontResolver');
        this.fontScheme = config.fontScheme;
        this.customFallbacks = config.customFallbacks;
        this.fontAvailability = config.fontAvailability ?? defaultFontAvailability;
        this.warnings = config.warnings;
    }
    /**
     * Resolves a font family name from PPTX to a canvas-compatible name.
     * Handles theme font references (+mj-lt, +mn-lt) and substitution.
     *
     * @param fontFamily Font family name from PPTX (may include theme references)
     * @returns Resolved font family name for canvas
     */
    resolveFontFamily(fontFamily) {
        if (!fontFamily) {
            return this.fontScheme.minorFont;
        }
        // Handle theme font references
        // +mj-lt = Major Latin font (headings)
        // +mn-lt = Minor Latin font (body)
        // +mj-ea = Major East Asian font
        // +mn-ea = Minor East Asian font
        // +mj-cs = Major Complex Script font
        // +mn-cs = Minor Complex Script font
        if (fontFamily.startsWith('+mj-lt') || fontFamily === '+mj-lt') {
            return this.fontScheme.majorFont;
        }
        if (fontFamily.startsWith('+mn-lt') || fontFamily === '+mn-lt') {
            return this.fontScheme.minorFont;
        }
        // The ea/cs slots commonly hold an empty typeface="" in real themes,
        // meaning "not specified" — fall back to the Latin font in that case too.
        const nonEmptyOr = (value, fallback) => value !== undefined && value !== '' ? value : fallback;
        if (fontFamily.startsWith('+mj-ea') || fontFamily === '+mj-ea') {
            return nonEmptyOr(this.fontScheme.majorFontEastAsian, this.fontScheme.majorFont);
        }
        if (fontFamily.startsWith('+mn-ea') || fontFamily === '+mn-ea') {
            return nonEmptyOr(this.fontScheme.minorFontEastAsian, this.fontScheme.minorFont);
        }
        if (fontFamily.startsWith('+mj-cs') || fontFamily === '+mj-cs') {
            return nonEmptyOr(this.fontScheme.majorFontComplexScript, this.fontScheme.majorFont);
        }
        if (fontFamily.startsWith('+mn-cs') || fontFamily === '+mn-cs') {
            return nonEmptyOr(this.fontScheme.minorFontComplexScript, this.fontScheme.minorFont);
        }
        // Return as-is for regular font names
        return fontFamily;
    }
    /**
     * Gets a font family with fallback chain for CSS/canvas use.
     *
     * The chain is composed as: requested family, then installed
     * metric-compatible substitutes (Carlito, Caladea, Liberation Sans/Serif),
     * then the family's fallback chain. Custom chains (per-instance config or
     * session-level via setCustomFontFallbacks) replace the built-in chain for
     * their family and suppress metric-substitute injection, since an explicit
     * user chain expresses complete intent.
     *
     * @param fontFamily Primary font family name
     * @returns Comma-separated font family string with fallbacks
     */
    getFontFamilyWithFallbacks(fontFamily) {
        const resolved = this.resolveFontFamily(fontFamily);
        // Custom chains (instance config wins over session, session over built-in)
        const customChain = this.customFallbacks?.[resolved] ?? sessionFallbacks[resolved];
        const baseChain = customChain ?? FONT_FALLBACK_CHAINS[resolved] ?? DEFAULT_FALLBACK;
        // Metric-compatible substitutes go before the geometric fallbacks, but
        // only when actually installed and only for non-custom chains.
        const metricSubstitutes = customChain
            ? []
            : (METRIC_COMPATIBLE_SUBSTITUTES[resolved] ?? []).filter((f) => this.isFontAvailable(f));
        // Compose and dedupe, keeping the resolved family first
        const chain = [];
        for (const family of [resolved, ...metricSubstitutes, ...baseChain]) {
            if (!chain.includes(family)) {
                chain.push(family);
            }
        }
        this.reportMissingFont(resolved, metricSubstitutes, chain);
        // Quote font names that contain spaces
        const quoted = chain.map((f) => (f.includes(' ') ? `"${f}"` : f));
        return quoted.join(', ');
    }
    /**
     * Emits a 'missing-font' warning (once per family per resolver) when the
     * requested family is unavailable and no metric-compatible substitute is
     * installed — i.e., text falls back to a geometrically different font, so
     * line breaks and overflow may not match PowerPoint. Metric-compatible
     * substitutions (Carlito for Calibri, ...) are silent by design.
     */
    reportMissingFont(resolved, installedMetricSubstitutes, chain) {
        if (GENERIC_FAMILIES.has(resolved) ||
            this.warnedMissingFamilies.has(resolved) ||
            installedMetricSubstitutes.length > 0 ||
            this.isFontAvailable(resolved)) {
            return;
        }
        this.warnedMissingFamilies.add(resolved);
        // The family actually used: first installed fallback, else the chain's
        // final generic keyword (whatever the canvas maps it to).
        const fallback = chain.slice(1).find((family) => this.isFontAvailable(family)) ?? chain[chain.length - 1];
        this.logger.warn('Font not available, falling back', { family: resolved, fallback });
        this.warnings?.push({
            code: 'missing-font',
            message: `Font "${resolved}" is not installed; falling back to "${fallback}"`,
            detail: { family: resolved, fallback },
        });
    }
    /**
     * Reports whether a font family is installed or registered, using the
     * configured availability predicate. Results are cached per instance.
     *
     * @param family Font family name
     * @returns True when the family can be rendered exactly
     */
    isFontAvailable(family) {
        const cached = this.availabilityCache.get(family);
        if (cached !== undefined) {
            return cached;
        }
        const available = this.fontAvailability(family);
        this.availabilityCache.set(family, available);
        return available;
    }
    /**
     * Resolves complete font information for canvas rendering.
     *
     * PPTX font sizes are in points; the canvas is addressed in pixels at
     * 96 DPI native resolution and rendered at the caller's render scale,
     * so the final glyph size is sizePoints x (96/72) x scale.
     *
     * @param fontFamily Font family name (may include theme refs)
     * @param sizePoints Font size in points (unscaled PPTX value)
     * @param bold Whether the font should be bold
     * @param italic Whether the font should be italic
     * @param scale Render scale factor (target pixels / native 96 DPI pixels)
     * @param letterSpacingPoints Letter tracking (a:rPr spc) in points, scaled
     *   to pixels with the same pt->px conversion as the font size
     * @param kerning Whether pair kerning applies (a:rPr kern threshold met).
     *   Defaults to true — the shaped fast path.
     * @returns Complete resolved font information
     */
    resolveFont(fontFamily, sizePoints, bold = false, italic = false, scale = 1, letterSpacingPoints = 0, kerning = true) {
        const family = this.getFontFamilyWithFallbacks(fontFamily ?? this.fontScheme.minorFont);
        // Final glyph size in pixels (rounded to avoid float noise in font strings)
        const sizePx = Math.round(pointsToPixels(sizePoints) * scale * 100) / 100;
        // Letter tracking in pixels, same pt->px scale as the glyph size
        const letterSpacingPx = Math.round(pointsToPixels(letterSpacingPoints) * scale * 100) / 100;
        // Build canvas font string: "bold italic 24px Arial"
        const parts = [];
        if (bold)
            parts.push('bold');
        if (italic)
            parts.push('italic');
        parts.push(`${sizePx}px`);
        parts.push(family);
        return {
            family,
            sizePoints,
            sizePx,
            bold,
            italic,
            fontString: parts.join(' '),
            ...(letterSpacingPx !== 0 && { letterSpacingPx }),
            ...(kerning === false && { noKern: true }),
        };
    }
    /**
     * Measures text using the specified font, honoring letter tracking.
     *
     * Tracking is applied natively via ctx.letterSpacing so shaping matches
     * drawing. skia-canvas inserts the spacing between glyphs ((n-1) gaps),
     * while PowerPoint's a:spc is part of every glyph's advance — so one
     * trailing spacing unit is added to keep wrap math faithful to PowerPoint.
     *
     * @param ctx Canvas 2D context
     * @param text Text to measure
     * @param font Resolved font information
     * @returns Width of the text in pixels (advance width incl. tracking)
     */
    measureText(ctx, text, font) {
        if (font.noKern) {
            // Unkerned runs advance glyph by glyph — kerning-free widths cannot
            // be obtained from shaped measurement (skia always kerns).
            let width = 0;
            for (const advance of this.charAdvances(ctx, text, font))
                width += advance;
            return width;
        }
        const spacingPx = font.letterSpacingPx ?? 0;
        ctx.save();
        ctx.font = font.fontString;
        if (spacingPx !== 0) {
            ctx.letterSpacing = `${spacingPx}px`;
        }
        const metrics = ctx.measureText(text);
        ctx.restore();
        if (spacingPx === 0) {
            return metrics.width;
        }
        // Trailing advance for the last glyph (PowerPoint advance semantics)
        const glyphCount = [...text].length;
        return glyphCount > 0 ? metrics.width + spacingPx : metrics.width;
    }
    /**
     * Per-glyph advances for unkerned (noKern) text: each code point's
     * isolated fractional advance width (cached per font+glyph) plus the
     * run's letter tracking. Measuring characters in isolation is exactly
     * kerning-free layout; the same advances drive both wrap measurement
     * (measureText above) and glyph placement (TextRenderer's unkerned draw
     * path), so they never disagree. Tracking follows PowerPoint semantics:
     * part of every glyph's advance, including the last.
     *
     * @param ctx Canvas 2D context
     * @param text Text to measure
     * @param font Resolved font information (noKern runs)
     * @returns Advance width per code point, in pixels
     */
    charAdvances(ctx, text, font) {
        const spacingPx = font.letterSpacingPx ?? 0;
        const advances = [];
        let cache = this.charAdvanceCache.get(font.fontString);
        if (!cache) {
            cache = new Map();
            this.charAdvanceCache.set(font.fontString, cache);
        }
        ctx.save();
        ctx.font = font.fontString;
        for (const char of text) {
            let advance = cache.get(char);
            if (advance === undefined) {
                advance = ctx.measureText(char).width;
                cache.set(char, advance);
            }
            advances.push(advance + spacingPx);
        }
        ctx.restore();
        return advances;
    }
    /**
     * Gets font metrics for the specified font.
     * Results are cached for performance.
     *
     * @param ctx Canvas 2D context
     * @param font Resolved font information
     * @returns Font metrics
     */
    getFontMetrics(ctx, font) {
        const cacheKey = font.fontString;
        const cached = this.metricsCache.get(cacheKey);
        if (cached) {
            return cached;
        }
        ctx.save();
        ctx.font = font.fontString;
        // Measure using canvas text metrics
        const metrics = ctx.measureText('Mgy');
        // Get ascent and descent from canvas metrics (pixels for the px-sized font)
        // Note: Some canvas implementations may not support all metrics
        const ascent = metrics.fontBoundingBoxAscent ?? metrics.actualBoundingBoxAscent ?? font.sizePx * 0.8;
        const descent = metrics.fontBoundingBoxDescent ?? metrics.actualBoundingBoxDescent ?? font.sizePx * 0.2;
        // Measure 'M' width for em-based calculations
        const emMetrics = ctx.measureText('M');
        const emWidth = emMetrics.width;
        // Measure average character width using a sample string
        const avgMetrics = ctx.measureText('abcdefghijklmnopqrstuvwxyz');
        const avgCharWidth = avgMetrics.width / 26;
        ctx.restore();
        const result = {
            ascent,
            descent,
            lineHeight: ascent + descent,
            emWidth,
            avgCharWidth,
        };
        this.metricsCache.set(cacheKey, result);
        this.logger.debug('Computed font metrics', {
            font: font.fontString,
            ascent,
            descent,
            lineHeight: result.lineHeight,
        });
        return result;
    }
    /**
     * Calculates line height based on font pixel size and line spacing.
     * Used as a fallback when real font metrics (ascent + descent from
     * measureText) are unavailable.
     *
     * @param fontSizePx Font size in pixels (ResolvedFont.sizePx)
     * @param lineSpacingPercent Line spacing as percentage (100 = single, 200 = double)
     * @returns Line height in pixels
     */
    calculateLineHeight(fontSizePx, lineSpacingPercent = 100) {
        // Standard line height is approximately 1.2x font size
        // Line spacing percentage modifies this
        const baseLineHeight = fontSizePx * 1.2;
        return baseLineHeight * (lineSpacingPercent / 100);
    }
    /**
     * Clears the metrics and font-availability caches.
     */
    clearCache() {
        this.metricsCache.clear();
        this.availabilityCache.clear();
        this.charAdvanceCache.clear();
    }
}
exports.FontResolver = FontResolver;
/**
 * Creates a FontResolver with the given font scheme.
 */
function createFontResolver(fontScheme = theme_js_1.DEFAULT_FONT_SCHEME, logger) {
    return new FontResolver({ fontScheme, logger });
}
