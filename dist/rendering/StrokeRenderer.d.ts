/**
 * Renders shape outlines (strokes) to canvas context.
 * Handles line properties including width, color, dash patterns, caps, and joins.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rgba, Path, Point } from '../types/geometry.js';
import type { Stroke, LineEnd } from '../types/elements.js';
import type { ResolvedTheme, ColorMap, ThemeLineStyle } from '../types/theme.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Default stroke width in EMU (1 point = 12700 EMU).
 */
export declare const DEFAULT_STROKE_WIDTH_EMU = 12700;
/**
 * Endpoint and outward direction for one end of an open path.
 */
export interface PathEndInfo {
    /** The terminal point of the path */
    point: Point;
    /** Unit vector pointing outward, away from the path interior */
    direction: Point;
}
/**
 * Start (head) and end (tail) information for an open path.
 */
export interface PathEndPoints {
    start: PathEndInfo;
    end: PathEndInfo;
}
/**
 * Renderable geometry for a single line end decoration.
 * `inset` is how far the underlying line should be pulled back along the
 * outward direction so it does not poke past the decoration tip.
 */
export type LineEndGeometry = {
    kind: 'polygon';
    points: Point[];
    inset: number;
} | {
    kind: 'ellipse';
    center: Point;
    radiusAlong: number;
    radiusAcross: number;
    rotation: number;
    inset: number;
} | {
    kind: 'openArrow';
    tip: Point;
    barbs: [Point, Point];
    inset: number;
};
/**
 * Computes the endpoints and outward tangent directions of an open path from
 * its first and last drawable segments. For bezier segments the tangent runs
 * through the control point adjacent to the endpoint (falling back to the
 * segment chord when the control point is degenerate); arcs use their chord
 * as an approximation. Returns undefined for closed or empty paths, which
 * have no free line ends to decorate.
 * @param path The path to analyze
 * @returns Start/end info, or undefined when the path has no open ends
 */
export declare function getPathEndPoints(path: Path): PathEndPoints | undefined;
/**
 * Computes the drawable geometry for a line end decoration.
 * Sizes follow Office conventions: width/length are the stroke width times
 * the sm/med/lg multiplier (2/3/5). Triangle and stealth heads sit with the
 * tip on the endpoint and extend back into the line; diamond and oval are
 * centered on the endpoint; the open arrow is two stroked barbs meeting at
 * the endpoint.
 * @param point Path endpoint the decoration is anchored to
 * @param direction Unit vector pointing outward from the path
 * @param end Line end decoration properties
 * @param lineWidthPx Stroke width in pixels
 * @returns The geometry, or undefined for type 'none'
 */
export declare function computeLineEndGeometry(point: Point, direction: Point, end: LineEnd, lineWidthPx: number): LineEndGeometry | undefined;
/**
 * Returns a copy of the path with its straight terminal segments pulled back
 * by the given insets so the line does not poke past filled arrowhead tips.
 * Only lineTo terminal segments are shortened; curved/arc ends keep their
 * full extent (acceptable simplification: the arrowhead fill covers the
 * overlap in most cases). Insets are clamped to the terminal segment length,
 * and scaled down proportionally when both ends shorten the same segment.
 * @param path The path to shorten (not mutated)
 * @param headInset Pixels to pull back the start of the path
 * @param tailInset Pixels to pull back the end of the path
 * @returns A new path with adjusted terminal points
 */
export declare function shortenPathEnds(path: Path, headInset: number, tailInset: number): Path;
/**
 * Configuration for StrokeRenderer.
 */
export interface StrokeRendererConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Renders strokes for shapes.
 */
export declare class StrokeRenderer {
    private readonly logger;
    private readonly colorResolver;
    private readonly unitConverter;
    constructor(config: StrokeRendererConfig);
    /**
     * Renders a stroke to the canvas for the given path.
     * @param ctx Canvas 2D context
     * @param path Path to stroke
     * @param stroke Stroke definition
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     */
    renderStroke(ctx: CanvasRenderingContext2D, path: Path, stroke: Stroke, scaleX: number, scaleY: number): void;
    /**
     * Parses stroke (outline) properties from a shape properties node.
     * Explicit a:ln properties win; properties the outline does not specify
     * fall back to the referenced style stroke (from a:lnRef), then defaults,
     * per the ECMA-376 style matrix precedence.
     * @param spPr Shape properties node containing a:ln
     * @param styleStroke Stroke resolved from the shape's a:lnRef, if any
     * @returns Parsed stroke or undefined if no stroke should render
     */
    parseStroke(spPr: PptxXmlNode | undefined, styleStroke?: Stroke): Stroke | undefined;
    /**
     * Resolves a theme line style (from lnStyleLst) into a renderable Stroke,
     * substituting phClr placeholders with the a:lnRef child color.
     * @param style The theme line style
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved stroke, or undefined when the style has no line fill
     */
    resolveThemeLineStyle(style: ThemeLineStyle, phClr?: Rgba): Stroke | undefined;
    /**
     * Maps a raw OOXML cap attribute value to a LineCap.
     */
    private mapOoxmlCap;
    /**
     * Parses stroke color from outline element.
     * @returns The explicit outline color, or undefined when none is declared
     */
    private parseStrokeColor;
    /**
     * Parses dash pattern from outline element.
     */
    private parseDashPattern;
    /**
     * Resolves a line end decoration for one end of the outline, applying the
     * ECMA-376 precedence: an explicit a:headEnd/a:tailEnd wins (including an
     * explicit or defaulted type="none", which clears the decoration), an
     * absent element inherits from the style stroke.
     * @param ln Outline element
     * @param tagName 'a:headEnd' or 'a:tailEnd'
     * @param inherited Decoration inherited from the style stroke, if any
     * @returns The effective decoration, or undefined for none
     */
    private resolveLineEnd;
    /**
     * Parses an a:headEnd / a:tailEnd child of an outline element.
     * Per ECMA-376 CT_LineEndProperties, type defaults to 'none' and w/len
     * default to 'med'.
     * @returns The decoration, the sentinel 'none' when the element is present
     * but declares no decoration, or undefined when the element is absent
     */
    private parseLineEnd;
    /**
     * Parses an ST_LineEndWidth / ST_LineEndLength attribute value, defaulting
     * to 'med' per ECMA-376.
     */
    private parseLineEndSize;
    /**
     * Draws a single line end decoration.
     * Filled shapes (triangle, stealth, diamond, oval) fill with the stroke
     * color; the open arrow strokes its two barbs at the line width with round
     * caps and joins.
     */
    private drawLineEnd;
    /**
     * Parses line cap from outline element.
     * @returns The explicit cap, or undefined when the attribute is absent
     */
    private parseLineCap;
    /**
     * Parses line join from outline element.
     * @returns The explicit join, or undefined when no join child is present
     */
    private parseLineJoin;
    /**
     * Maps LineCap type to canvas lineCap value.
     */
    private mapLineCap;
    /**
     * Maps LineJoin type to canvas lineJoin value.
     */
    private mapLineJoin;
    /**
     * Gets stroke style string for a color (utility method).
     */
    getStrokeStyle(color: Rgba): string;
    /**
     * Converts EMU stroke width to pixels.
     */
    strokeWidthToPixels(widthEmu: number, scale: number): number;
}
/**
 * Default stroke renderer factory.
 */
export declare function createStrokeRenderer(theme: ResolvedTheme, logger?: ILogger): StrokeRenderer;
//# sourceMappingURL=StrokeRenderer.d.ts.map