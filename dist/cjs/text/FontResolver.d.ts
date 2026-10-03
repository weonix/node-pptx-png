/**
 * Resolves font names from PPTX to system fonts.
 * Handles theme fonts (+mj-lt, +mn-lt) and font substitution fallback chains.
 */
import { type CanvasRenderingContext2D } from 'skia-canvas';
import type { ResolvedFontScheme } from '../types/theme.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Pixels per point at the canvas-native 96 DPI (1pt = 1/72 inch, 1px = 1/96 inch).
 */
export declare const PIXELS_PER_POINT: number;
/**
 * Converts a size in points to pixels at the canvas-native 96 DPI.
 *
 * @param points Size in points
 * @returns Size in pixels (unscaled)
 */
export declare function pointsToPixels(points: number): number;
/**
 * Font metrics for a specific font at a specific size.
 */
export interface FontMetrics {
    /** Font ascent (distance from baseline to top) in pixels */
    ascent: number;
    /** Font descent (distance from baseline to bottom) in pixels */
    descent: number;
    /** Total line height (ascent + descent) in pixels */
    lineHeight: number;
    /** Em width (approximate width of 'M' character) in pixels */
    emWidth: number;
    /** Average character width in pixels */
    avgCharWidth: number;
}
/**
 * Resolved font information ready for canvas use.
 */
export interface ResolvedFont {
    /** Canvas-compatible font family string */
    family: string;
    /** Original PPTX font size in points (unscaled) */
    sizePoints: number;
    /** Final glyph size in pixels: sizePoints x (96/72) x render scale */
    sizePx: number;
    /** Whether the font is bold */
    bold: boolean;
    /** Whether the font is italic */
    italic: boolean;
    /** Complete font string for canvas (e.g., "bold italic 24px Arial") */
    fontString: string;
    /**
     * Letter tracking (a:rPr spc) in render pixels, already scaled like sizePx.
     * Applied natively via ctx.letterSpacing for measurement and drawing.
     * Absent (or 0) means no tracking — the untracked fast path.
     */
    letterSpacingPx?: number;
    /**
     * Pair kerning is disabled for this run (a:rPr kern threshold not met:
     * the run is smaller than the minimum kerning size, so PowerPoint lays
     * glyphs out on bare advances). skia-canvas always shapes with kerning,
     * so unkerned runs are measured and drawn per glyph instead. Absent
     * means kerning stays on — the shaped fast path.
     */
    noKern?: true;
}
/**
 * Configuration for FontResolver.
 */
export interface FontResolverConfig {
    /** Resolved font scheme from theme */
    fontScheme: ResolvedFontScheme;
    /** Logger instance */
    logger?: ILogger;
    /**
     * Custom fallback chains merged over the built-in chains (and over any
     * session fallbacks set via setCustomFontFallbacks). A custom chain
     * replaces the built-in chain for that family.
     */
    customFallbacks?: Record<string, string[]>;
    /**
     * Predicate reporting whether a font family is installed/registered.
     * Defaults to checking FontManager-registered families and the
     * skia-canvas FontLibrary. Injectable for tests.
     */
    fontAvailability?: (family: string) => boolean;
    /**
     * Structured warning collector. Receives a 'missing-font' warning (once
     * per family per resolver) when a requested font is unavailable and no
     * metric-compatible substitute is installed, so the render falls back to
     * a geometrically different font.
     */
    warnings?: WarningCollector;
}
/**
 * Sets (replaces) the session custom font fallback chains.
 *
 * @param fallbacks Chains keyed by requested family name; undefined clears
 *   any previously set session fallbacks
 */
export declare function setCustomFontFallbacks(fallbacks?: Record<string, string[]>): void;
/**
 * Returns the current session custom font fallback chains.
 */
export declare function getCustomFontFallbacks(): Readonly<Record<string, string[]>>;
/**
 * Resolves font names from PPTX to canvas-compatible fonts.
 */
export declare class FontResolver {
    private readonly logger;
    private readonly fontScheme;
    private readonly metricsCache;
    /** Isolated glyph advances per font string, for unkerned (noKern) runs */
    private readonly charAdvanceCache;
    private readonly customFallbacks;
    private readonly fontAvailability;
    private readonly availabilityCache;
    private readonly warnings;
    /** Families already reported as missing (one warning per family) */
    private readonly warnedMissingFamilies;
    constructor(config: FontResolverConfig);
    /**
     * Resolves a font family name from PPTX to a canvas-compatible name.
     * Handles theme font references (+mj-lt, +mn-lt) and substitution.
     *
     * @param fontFamily Font family name from PPTX (may include theme references)
     * @returns Resolved font family name for canvas
     */
    resolveFontFamily(fontFamily: string | undefined): string;
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
    getFontFamilyWithFallbacks(fontFamily: string): string;
    /**
     * Emits a 'missing-font' warning (once per family per resolver) when the
     * requested family is unavailable and no metric-compatible substitute is
     * installed — i.e., text falls back to a geometrically different font, so
     * line breaks and overflow may not match PowerPoint. Metric-compatible
     * substitutions (Carlito for Calibri, ...) are silent by design.
     */
    private reportMissingFont;
    /**
     * Reports whether a font family is installed or registered, using the
     * configured availability predicate. Results are cached per instance.
     *
     * @param family Font family name
     * @returns True when the family can be rendered exactly
     */
    isFontAvailable(family: string): boolean;
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
    resolveFont(fontFamily: string | undefined, sizePoints: number, bold?: boolean, italic?: boolean, scale?: number, letterSpacingPoints?: number, kerning?: boolean): ResolvedFont;
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
    measureText(ctx: CanvasRenderingContext2D, text: string, font: ResolvedFont): number;
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
    charAdvances(ctx: CanvasRenderingContext2D, text: string, font: ResolvedFont): number[];
    /**
     * Gets font metrics for the specified font.
     * Results are cached for performance.
     *
     * @param ctx Canvas 2D context
     * @param font Resolved font information
     * @returns Font metrics
     */
    getFontMetrics(ctx: CanvasRenderingContext2D, font: ResolvedFont): FontMetrics;
    /**
     * Calculates line height based on font pixel size and line spacing.
     * Used as a fallback when real font metrics (ascent + descent from
     * measureText) are unavailable.
     *
     * @param fontSizePx Font size in pixels (ResolvedFont.sizePx)
     * @param lineSpacingPercent Line spacing as percentage (100 = single, 200 = double)
     * @returns Line height in pixels
     */
    calculateLineHeight(fontSizePx: number, lineSpacingPercent?: number): number;
    /**
     * Clears the metrics and font-availability caches.
     */
    clearCache(): void;
}
/**
 * Creates a FontResolver with the given font scheme.
 */
export declare function createFontResolver(fontScheme?: ResolvedFontScheme, logger?: ILogger): FontResolver;
