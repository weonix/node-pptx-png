/**
 * Text layout engine for measuring and positioning text within bounds.
 * Handles paragraph properties, alignment, spacing, and indentation.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rect, Rgba } from '../types/geometry.js';
import type { TextBody } from '../types/elements.js';
import type { ResolvedFontScheme } from '../types/theme.js';
import { FontResolver, type ResolvedFont } from './FontResolver.js';
import { BulletFormatter } from './BulletFormatter.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * A positioned text run ready for rendering.
 */
export interface PositionedTextRun {
    /** Text content */
    text: string;
    /** X position in pixels */
    x: number;
    /** Y position (baseline) in pixels */
    y: number;
    /** Width of the text in pixels (pre-measured during layout) */
    width: number;
    /** Font for rendering */
    font: ResolvedFont;
    /**
     * Text color from the run/inheritance chain. Undefined when no source in
     * the chain specifies a color; the renderer then falls back to the shape's
     * default text color (style fontRef or fill contrast) and finally black.
     */
    color?: Rgba;
    /** Whether text is underlined */
    underline?: boolean;
    /** Whether text has strikethrough */
    strikethrough?: boolean;
    /** Baseline offset for super/subscript (percentage) */
    baselineOffset?: number;
}
/**
 * A positioned bullet ready for rendering.
 */
export interface PositionedBullet {
    /** Bullet text */
    text: string;
    /** X position in pixels */
    x: number;
    /** Y position (baseline) in pixels */
    y: number;
    /** Font for rendering */
    font: ResolvedFont;
    /** Bullet color (undefined = use text color) */
    color?: Rgba;
}
/**
 * A laid out line of text.
 */
export interface LayoutLine {
    /** Positioned text runs in this line */
    runs: PositionedTextRun[];
    /** Bullet for this line (if any, only first line of paragraph) */
    bullet?: PositionedBullet;
    /** Line Y position (top of line) in pixels */
    y: number;
    /** Line height in pixels */
    height: number;
    /** Total width of line content in pixels */
    width: number;
}
/**
 * Complete text layout result.
 */
export interface TextLayout {
    /** All laid out lines */
    lines: LayoutLine[];
    /** Total height of text content in pixels */
    totalHeight: number;
    /** Maximum width of any line in pixels */
    maxWidth: number;
    /** Bounds used for layout */
    bounds: Rect;
}
/**
 * Configuration for TextLayoutEngine.
 */
export interface TextLayoutEngineConfig {
    /** Font scheme from theme */
    fontScheme: ResolvedFontScheme;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector, forwarded to the FontResolver */
    warnings?: WarningCollector;
}
/**
 * Resolves the line advance in pixels from a parsed line-spacing value.
 *
 * TextParser encodes a:lnSpc as a single number: positive values are
 * percentages in 1000ths of a percent (a:spcPct, 100000 = 100%), negative
 * values are fixed spacing in hundredths of a point (a:spcPts).
 *
 * @param lineSpacing Encoded line spacing (undefined = single spacing)
 * @param singleLineHeightPx Single-spaced line height in pixels
 * @param scaleY Vertical render scale factor
 * @returns Line height in pixels
 */
export declare function resolveLineHeightPx(lineSpacing: number | undefined, singleLineHeightPx: number, scaleY: number): number;
/**
 * Resolves paragraph spacing (space before/after) in pixels.
 *
 * TextParser encodes a:spcBef / a:spcAft as a single number: positive values
 * are absolute spacing in EMU (a:spcPts), negative values are percentages of
 * the line height in 1000ths of a percent (a:spcPct, -50000 = 50%).
 *
 * @param spacing Encoded spacing value (undefined = no spacing)
 * @param singleLineHeightPx Single-spaced line height in pixels
 * @param scaleY Vertical render scale factor
 * @returns Spacing in pixels
 */
export declare function resolveParagraphSpacingPx(spacing: number | undefined, singleLineHeightPx: number, scaleY: number): number;
/**
 * Splits text into segments for cap="small" rendering: characters with a
 * distinct uppercase form (i.e. lowercase letters) become smallCap segments
 * carrying their uppercased text, drawn at SMALL_CAPS_SCALE of the font
 * size; all other characters (capitals, digits, punctuation, whitespace)
 * stay full-size. Consecutive characters of the same class are grouped.
 * Uses locale-independent Unicode default case conversion.
 *
 * @param text Run text to segment
 * @returns Ordered segments whose concatenated text is the transformed run
 */
export declare function segmentSmallCaps(text: string): {
    text: string;
    smallCap: boolean;
}[];
/**
 * Text layout engine for measuring and positioning text.
 */
export declare class TextLayoutEngine {
    private readonly logger;
    private readonly fontResolver;
    private readonly wordWrapper;
    private readonly bulletFormatter;
    constructor(config: TextLayoutEngineConfig);
    /**
     * Lays out text within the specified bounds.
     *
     * A text body with an empty a:normAutofit (no stored fontScale /
     * lnSpcReduction) is solved iteratively: the text is laid out at
     * decreasing font scales until it fits the box, matching PowerPoint's
     * render-time shrink-to-fit. Stored values, spAutoFit and noAutofit are
     * honored as before.
     *
     * @param ctx Canvas 2D context for text measurement
     * @param textBody Text body to lay out
     * @param shapeBounds Shape bounds in pixels
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     * @returns Complete text layout
     */
    layoutText(ctx: CanvasRenderingContext2D, textBody: TextBody, shapeBounds: Rect, scaleX: number, scaleY: number): TextLayout;
    /**
     * Runs one full layout pass (wrapping, positioning, alignment) with the
     * given autofit multipliers applied. Both the stored-value path and each
     * solver candidate go through this exact code path, so a solved scale
     * renders identically to the same scale stored in the file.
     */
    private layoutWithScales;
    /**
     * Calculates text bounds with insets applied.
     */
    private calculateTextBounds;
    /**
     * Lays out a single paragraph.
     */
    private layoutParagraph;
    /**
     * Stretches a positioned line to fill targetWidth by distributing the
     * shortfall evenly across its word gaps (maximal whitespace stretches
     * with words on both sides, including gaps spanning run boundaries).
     * The runs array is re-split in place into word- and gap-level runs so
     * each piece carries its own x position; gap runs keep their source
     * run's decorations so underlines span the widened gaps.
     * @returns true when the line was stretched to targetWidth
     */
    private justifyLine;
    /**
     * Builds text fragments from paragraph runs.
     * Fonts are resolved with the render scale so measurement, wrapping and
     * rendering all use the same scaled glyph size; the normAutofit fontScale
     * multiplier is applied to every resolved size.
     *
     * Capitalization (a:rPr cap) is applied here, before wrapping, so line
     * breaks are computed against the transformed text: cap="all" uppercases
     * the fragment text; cap="small" splits the run into full-size and
     * reduced-size uppercase segments. Letter tracking (a:rPr spc) rides on
     * the resolved font so measurement and drawing stay in sync. Tracking is
     * kept absolute (not multiplied by the autofit fontScale), matching how
     * PowerPoint stores spc independently of normAutofit font scaling.
     */
    private buildFragments;
    /**
     * Merges default run properties with specific run properties.
     */
    private mergeRunProperties;
    /**
     * Converts BulletConfig to BulletProps.
     */
    private convertBulletConfig;
    /**
     * Whether the text body requires a render-time shrink-to-fit solve: an
     * a:normAutofit that carries no stored values. PowerPoint persists solved
     * fontScale/lnSpcReduction on save; when either attribute is present the
     * file already holds the solved state and is honored as-is (stored values
     * win — no re-solving). Only a bare `<a:normAutofit/>` is solved here.
     */
    private needsAutofitSolve;
    /**
     * Iteratively solves shrink-to-fit for an empty a:normAutofit: finds the
     * largest candidate scale (see AutofitSolver for the PowerPoint-observed
     * ladder) whose full layout — including re-wrapping at the scaled font
     * sizes — fits the available text-box height, then returns that layout.
     * Layout passes are cached per candidate so the winner is never recomputed;
     * the solver binary-searches the ladder, costing at most ~6 passes.
     */
    private solveAndLayout;
    /**
     * Resolves the normAutofit multipliers stored on the text body. Only
     * meaningful values are honored: fontScale must be a positive fraction
     * and lnSpcReduction a fraction below 1.
     */
    private getAutofitScales;
    /**
     * Resolves the default font and single-spaced line metrics for a paragraph.
     *
     * PowerPoint bases single (and percentage) line spacing on 1.2x the font
     * point size, NOT on the font's own ascent+descent metrics: an 8pt Open
     * Sans paragraph advances 9.6pt per line even though the font's hhea
     * metrics sum to 1.36em. Using font metrics here made every dense text
     * block drift vertically against PowerPoint output.
     */
    private getParagraphLineMetrics;
    /**
     * Gets the default font for a paragraph, scaled to render pixels and by
     * the normAutofit fontScale multiplier.
     */
    private getDefaultFont;
    /**
     * Applies horizontal alignment to a layout line.
     */
    private applyHorizontalAlignment;
    /**
     * Calculates vertical offset for alignment.
     */
    private calculateVerticalOffset;
    /**
     * Gets the font resolver for external use.
     */
    getFontResolver(): FontResolver;
    /**
     * Gets the bullet formatter for external use.
     */
    getBulletFormatter(): BulletFormatter;
}
/**
 * Creates a TextLayoutEngine instance.
 */
export declare function createTextLayoutEngine(fontScheme: ResolvedFontScheme, logger?: ILogger): TextLayoutEngine;
