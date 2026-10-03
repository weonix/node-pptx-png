/**
 * Renders fills (solid, gradient, pattern, picture) to canvas context.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rgba, Rect, Path } from '../types/geometry.js';
import type { Fill, PictureFill } from '../types/elements.js';
import type { ResolvedTheme, ColorMap, ThemeFill } from '../types/theme.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
import type { ImageRenderer, CropRect, TileInfo } from './ImageRenderer.js';
/**
 * Extended picture fill with parsed source rect and tile info.
 * Note: CropRect fields use OpenXML percentage format (0-100000 = 0-100%),
 * which is distinct from geometry.Rect that uses pixel coordinates.
 */
export interface ExtendedPictureFill extends PictureFill {
    /** Source rectangle for cropping (percentage-based, parsed from blipFill) */
    srcRect?: CropRect;
    /** Tile info (parsed from blipFill) */
    tile?: TileInfo;
    /** Fill rectangle for stretch mode (percentage-based) */
    fillRect?: CropRect;
    /** Whether stretch mode is enabled */
    stretch?: boolean;
    /** The raw blipFill node for ImageRenderer */
    blipFillNode?: PptxXmlNode;
}
/**
 * Start and end points of a linear gradient line.
 */
export interface LinearGradientPoints {
    /** Gradient line start x */
    x0: number;
    /** Gradient line start y */
    y0: number;
    /** Gradient line end x */
    x1: number;
    /** Gradient line end y */
    y1: number;
}
/**
 * Computes the start/end points of a linear gradient line through a center point.
 * Per ECMA-376 (a:lin@ang), the gradient angle is measured clockwise from the
 * positive x-axis in y-down slide space: 0 = left-to-right, 90 = top-to-bottom,
 * 180 = right-to-left, 270 = bottom-to-top. Canvas coordinates are also y-down,
 * so the angle maps directly to (cos, sin) without adjustment.
 * @param centerX Center x of the fill area
 * @param centerY Center y of the fill area
 * @param halfDiagonal Half the diagonal length of the fill area (ensures coverage)
 * @param angleDegrees Gradient angle in degrees, clockwise from the positive x-axis
 * @returns Gradient line start and end points
 */
export declare function computeLinearGradientPoints(centerX: number, centerY: number, halfDiagonal: number, angleDegrees: number): LinearGradientPoints;
/**
 * Configuration for FillRenderer.
 */
export interface FillRendererConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Renders fills for shapes.
 */
export declare class FillRenderer {
    private readonly logger;
    private readonly colorResolver;
    constructor(config: FillRendererConfig);
    /**
     * Renders a fill to the canvas for the given path.
     * For picture fills, use renderFillAsync instead.
     * @param ctx Canvas 2D context
     * @param path Path to fill
     * @param fill Fill definition
     * @param bounds Bounding rectangle of the shape
     */
    renderFill(ctx: CanvasRenderingContext2D, path: Path, fill: Fill, bounds: Rect): void;
    /**
     * Renders a fill to the canvas for the given path, with async support for picture fills.
     * @param ctx Canvas 2D context
     * @param path Path to fill
     * @param fill Fill definition
     * @param bounds Bounding rectangle of the shape
     * @param imageRenderer Optional ImageRenderer for picture fills
     */
    renderFillAsync(ctx: CanvasRenderingContext2D, path: Path, fill: Fill, bounds: Rect, imageRenderer?: ImageRenderer): Promise<void>;
    /**
     * Renders a picture fill.
     * Delegates all parsing and rendering to ImageRenderer.
     */
    private renderPictureFill;
    /**
     * Renders a solid fill.
     */
    private renderSolidFill;
    /**
     * Renders a gradient fill.
     */
    private renderGradientFill;
    /**
     * Creates a linear gradient for the given bounds and fill.
     */
    private createLinearGradient;
    /**
     * Creates a radial gradient for the given bounds and fill.
     */
    private createRadialGradient;
    /**
     * Parses fill properties from a shape properties node.
     * @param spPr Shape properties node
     * @returns Parsed fill or undefined if no fill specified
     */
    parseFill(spPr: PptxXmlNode | undefined): Fill | undefined;
    /**
     * Parses a solid fill element.
     */
    private parseSolidFill;
    /**
     * Parses a gradient fill element.
     */
    private parseGradientFill;
    /**
     * Parses a pattern fill element.
     */
    private parsePatternFill;
    /**
     * Parses a picture fill element.
     * Note: Full parsing of srcRect, tile, fillRect is deferred to ImageRenderer
     * which will parse the blipFillNode when rendering. This avoids duplicate parsing logic.
     */
    private parsePictureFill;
    /**
     * Resolves a theme style-matrix fill (from fillStyleLst/bgFillStyleLst)
     * into a renderable Fill, substituting phClr placeholders with the color
     * supplied by the style reference (a:fillRef child color).
     * @param themeFill The theme fill definition
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved fill, or undefined when required colors are missing
     */
    resolveThemeFill(themeFill: ThemeFill, phClr?: Rgba): Fill | undefined;
    /**
     * Gets fill style string for a solid color (utility method).
     */
    getFillStyle(color: Rgba): string;
}
/**
 * Default fill renderer factory.
 */
export declare function createFillRenderer(theme: ResolvedTheme, logger?: ILogger): FillRenderer;
//# sourceMappingURL=FillRenderer.d.ts.map