/**
 * Decodes image data from Buffer to canvas-compatible Image objects.
 * Handles different image formats (PNG, JPEG, GIF, BMP, WebP), plus
 * EMF/WMF vector metafiles which are rasterized via the emf module.
 */
import { type Image } from 'skia-canvas';
import type { ILogger } from './Logger.js';
import type { WarningCollector } from './WarningCollector.js';
/**
 * Supported image formats.
 */
export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tiff' | 'emf' | 'wmf' | 'unknown';
/**
 * Result of decoding an image.
 */
export interface DecodedImage {
    /** The decoded Image object */
    image: Image;
    /** Image width in pixels */
    width: number;
    /** Image height in pixels */
    height: number;
    /** Detected format */
    format: ImageFormat;
}
/**
 * Configuration for ImageDecoder.
 */
export interface ImageDecoderConfig {
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector for decode fidelity events */
    warnings?: WarningCollector;
}
/**
 * Decodes images from binary data to canvas-compatible Image objects.
 * Note: Caching is handled at the ImageRenderer level by relationship ID,
 * which is the correct key for deduplication.
 */
export declare class ImageDecoder {
    private readonly logger;
    private readonly warnings;
    constructor(config?: ImageDecoderConfig);
    /**
     * Detects the image format from the buffer's magic bytes.
     */
    detectFormat(buffer: Buffer): ImageFormat;
    /**
     * Checks if a buffer starts with the given signature bytes.
     */
    private matchesSignature;
    /**
     * Decodes an image from a Buffer.
     *
     * @param buffer The image data as a Buffer
     * @returns The decoded image with metadata
     * @throws Error if the image cannot be decoded
     */
    decode(buffer: Buffer): Promise<DecodedImage>;
    /**
     * Rasterizes an EMF/WMF vector metafile (parse -> replay onto an
     * offscreen canvas -> PNG -> Image) so callers can treat it like any
     * other decoded image.
     *
     * @param buffer The metafile contents
     * @param format 'emf' or 'wmf'
     * @returns The rasterized image with metadata
     * @throws Error if the metafile cannot be parsed or rendered (callers
     * treat this like any other decode failure: warn and skip the picture)
     */
    private decodeMetafile;
    /**
     * Decodes a TIFF image by transcoding to PNG via the optional `sharp`
     * peer dependency (skia has no TIFF codec). Decks commonly embed TIFF
     * logos; without sharp installed the picture is skipped with a warning,
     * matching the behavior for other undecodable media.
     *
     * @param buffer TIFF bytes
     * @returns The decoded image
     * @throws Error when sharp is unavailable or the TIFF is corrupt
     */
    private decodeTiff;
    /**
     * Decodes an image from a data URI.
     *
     * @param dataUri The data URI string (e.g., data:image/png;base64,...)
     * @returns The decoded image with metadata
     */
    decodeDataUri(dataUri: string): Promise<DecodedImage>;
    /**
     * Gets the MIME type for an image format.
     */
    getMimeType(format: ImageFormat): string;
    /**
     * Gets the file extension for an image format.
     */
    getExtension(format: ImageFormat): string;
}
/**
 * Creates an ImageDecoder instance.
 */
export declare function createImageDecoder(logger?: ILogger): ImageDecoder;
