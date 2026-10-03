/**
 * Renders images to canvas context.
 * Handles picture shapes (p:pic elements), blipFill for shape fills,
 * source rectangle cropping, and stretch/tile fill modes.
 */
import type { CanvasRenderingContext2D, Image } from 'skia-canvas';
import type { Rect } from '../types/geometry.js';
import type { PptxParser, PptxXmlNode } from '../core/PptxParser.js';
import { type DecodedImage } from '../utils/ImageDecoder.js';
import { type PixelTransform } from '../geometry/TransformCalculator.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Crop rectangle with percentage-based values.
 * Used for image cropping in OpenXML where values represent percentages.
 * Values are in percentage units (0-100000 = 0-100%).
 *
 * Note: This is distinct from geometry.Rect which uses pixel coordinates.
 * OpenXML uses this percentage format for srcRect and fillRect in blipFill elements.
 */
export interface CropRect {
    /** Left crop percentage (0-100000 where 100000 = 100%) */
    left: number;
    /** Top crop percentage (0-100000 where 100000 = 100%) */
    top: number;
    /** Right crop percentage (0-100000 where 100000 = 100%) */
    right: number;
    /** Bottom crop percentage (0-100000 where 100000 = 100%) */
    bottom: number;
}
/**
 * Tile fill settings for repeating an image.
 */
export interface TileInfo {
    /** Horizontal scale percentage (100000 = 100%) */
    sx: number;
    /** Vertical scale percentage (100000 = 100%) */
    sy: number;
    /** Horizontal offset in EMU */
    tx: number;
    /** Vertical offset in EMU */
    ty: number;
    /** Flip mode for tiles */
    flip: 'none' | 'x' | 'y' | 'xy';
    /** Alignment anchor */
    alignment: string;
}
/**
 * Picture data parsed from p:pic or blipFill elements.
 */
export interface PictureData {
    /** Relationship ID for embedded image (r:embed) */
    blipRelId: string;
    /** Source rectangle for cropping (percentage-based, see CropRect) */
    srcRect?: CropRect;
    /** Whether to stretch the image to fill bounds */
    stretch?: boolean;
    /** Tile fill settings */
    tile?: TileInfo;
    /** Fill rectangle for stretch mode (percentage-based, see CropRect) */
    fillRect?: CropRect;
    /**
     * Picture opacity from a:blip/a:alphaModFix (0-1). Undefined = opaque.
     * Watermark-style images commonly use low values (e.g. 0.16).
     */
    alpha?: number;
}
/**
 * Computes the destination rectangle for a stretched image with an a:fillRect.
 * Per ECMA-376 (20.1.8.30), positive fillRect percentages are INSETS — the
 * image is drawn inside a rectangle smaller than the shape bounds — while
 * negative percentages are outsets that extend beyond the bounds.
 * @param bounds The shape bounds in pixels
 * @param fillRect The fill rectangle percentages (100000 = 100%)
 * @returns The destination rectangle in pixels
 */
export declare function computeStretchDestRect(bounds: Rect, fillRect: CropRect): Rect;
/**
 * Configuration for ImageRenderer.
 */
export interface ImageRendererConfig {
    /** PPTX parser instance */
    parser: PptxParser;
    /** Source file path for relationship resolution (e.g., ppt/slides/slide1.xml) */
    sourcePath: string;
    /** Horizontal scale factor (EMU to pixels) */
    scaleX: number;
    /** Vertical scale factor (EMU to pixels) */
    scaleY: number;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector for media fidelity events */
    warnings?: WarningCollector;
}
/**
 * Renders images from PPTX media folder to canvas.
 */
export declare class ImageRenderer {
    private readonly logger;
    private readonly parser;
    private readonly sourcePath;
    private readonly scaleX;
    private readonly scaleY;
    private readonly relationshipParser;
    private readonly imageDecoder;
    private readonly transformCalculator;
    private readonly unitConverter;
    /** Cache of loaded images, keyed by relationship ID */
    private readonly imageCache;
    /** Maximum number of images to cache (LRU eviction) */
    private readonly maxCacheSize;
    /** Track insertion order for LRU eviction */
    private readonly cacheOrder;
    private readonly warnings;
    constructor(config: ImageRendererConfig);
    /**
     * Adds an image to the cache with LRU eviction.
     * @param key The cache key (relationship ID)
     * @param image The decoded image to cache
     */
    private addToCache;
    /**
     * Loads an image by its relationship ID.
     *
     * @param relationshipId The r:embed relationship ID
     * @returns The decoded image or undefined if not found
     */
    loadImage(relationshipId: string): Promise<DecodedImage | undefined>;
    /**
     * Parses a blipFill element to extract picture data.
     *
     * @param blipFill The a:blipFill element node
     * @returns Parsed picture data or undefined
     */
    parseBlipFill(blipFill: PptxXmlNode | undefined): PictureData | undefined;
    /**
     * Parses a p:pic element to extract picture data.
     *
     * @param picNode The p:pic element node
     * @returns Parsed picture data or undefined
     */
    parsePicElement(picNode: PptxXmlNode | undefined): PictureData | undefined;
    /**
     * Parses a crop rectangle element (srcRect or fillRect).
     * Values are in OpenXML percentage format (0-100000).
     */
    private parseCropRect;
    /**
     * Parses tile fill settings.
     */
    private parseTileInfo;
    /**
     * Converts a percentage value (0-100000) to a decimal (0-1).
     */
    private percentToDecimal;
    /**
     * Renders an image to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param image The decoded image to render
     * @param bounds The destination bounds in pixels
     * @param pictureData Optional picture data with cropping/fill settings
     */
    renderImage(ctx: CanvasRenderingContext2D, image: Image, bounds: Rect, pictureData?: PictureData): void;
    /**
     * Renders a tiled image pattern.
     */
    private renderTiledImage;
    /**
     * Renders a picture element (p:pic) to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param picNode The p:pic XML node
     * @param transform The pixel transform for the picture
     */
    renderPictureElement(ctx: CanvasRenderingContext2D, picNode: PptxXmlNode, transform: PixelTransform): Promise<void>;
    /**
     * Renders a picture fill (blipFill) to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param blipFill The a:blipFill XML node
     * @param bounds The destination bounds in pixels
     */
    renderPictureFill(ctx: CanvasRenderingContext2D, blipFill: PptxXmlNode, bounds: Rect): Promise<void>;
    /**
     * Clears the image cache.
     */
    clearCache(): void;
    /**
     * Updates the source path for relationship resolution.
     * Call this when rendering a different slide.
     */
    setSourcePath(sourcePath: string): ImageRenderer;
}
/**
 * Creates an ImageRenderer instance.
 */
export declare function createImageRenderer(parser: PptxParser, sourcePath: string, scaleX: number, scaleY: number, logger?: ILogger): ImageRenderer;
//# sourceMappingURL=ImageRenderer.d.ts.map