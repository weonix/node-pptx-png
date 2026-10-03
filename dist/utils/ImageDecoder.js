/**
 * Decodes image data from Buffer to canvas-compatible Image objects.
 * Handles different image formats (PNG, JPEG, GIF, BMP, WebP), plus
 * EMF/WMF vector metafiles which are rasterized via the emf module.
 */
import { loadImage } from 'skia-canvas';
import { createLogger } from './Logger.js';
import { EmfParser, EmfRenderer, WmfParser, isEmf, isWmf } from '../emf/index.js';
/**
 * Image signature bytes for format detection.
 */
const IMAGE_SIGNATURES = {
    png: [0x89, 0x50, 0x4e, 0x47], // .PNG
    jpeg: [0xff, 0xd8, 0xff], // JPEG SOI marker
    gif: [0x47, 0x49, 0x46], // GIF
    bmp: [0x42, 0x4d], // BM
    webp: [0x52, 0x49, 0x46, 0x46], // RIFF (WebP container)
};
/**
 * Decodes images from binary data to canvas-compatible Image objects.
 * Note: Caching is handled at the ImageRenderer level by relationship ID,
 * which is the correct key for deduplication.
 */
export class ImageDecoder {
    logger;
    warnings;
    constructor(config = {}) {
        this.logger = config.logger ?? createLogger('warn', 'ImageDecoder');
        this.warnings = config.warnings;
    }
    /**
     * Detects the image format from the buffer's magic bytes.
     */
    detectFormat(buffer) {
        if (buffer.length < 4) {
            return 'unknown';
        }
        // Check PNG signature
        if (this.matchesSignature(buffer, IMAGE_SIGNATURES.png)) {
            return 'png';
        }
        // Check JPEG signature
        if (this.matchesSignature(buffer, IMAGE_SIGNATURES.jpeg)) {
            return 'jpeg';
        }
        // Check GIF signature
        if (this.matchesSignature(buffer, IMAGE_SIGNATURES.gif)) {
            return 'gif';
        }
        // Check BMP signature
        if (this.matchesSignature(buffer, IMAGE_SIGNATURES.bmp)) {
            return 'bmp';
        }
        // Check WebP (RIFF container with WEBP)
        if (this.matchesSignature(buffer, IMAGE_SIGNATURES.webp) &&
            buffer.length >= 12 &&
            buffer[8] === 0x57 && // W
            buffer[9] === 0x45 && // E
            buffer[10] === 0x42 && // B
            buffer[11] === 0x50 // P
        ) {
            return 'webp';
        }
        // Check TIFF (II*\0 little-endian or MM\0* big-endian)
        if (buffer.length >= 4 &&
            ((buffer[0] === 0x49 && buffer[1] === 0x49 && buffer[2] === 0x2a && buffer[3] === 0x00) ||
                (buffer[0] === 0x4d && buffer[1] === 0x4d && buffer[2] === 0x00 && buffer[3] === 0x2a))) {
            return 'tiff';
        }
        // Check EMF (EMR_HEADER iType=1 with ' EMF' signature at offset 40)
        if (isEmf(buffer)) {
            return 'emf';
        }
        // Check WMF (placeable 0x9AC6CDD7 key or standard header)
        if (isWmf(buffer)) {
            return 'wmf';
        }
        return 'unknown';
    }
    /**
     * Checks if a buffer starts with the given signature bytes.
     */
    matchesSignature(buffer, signature) {
        if (buffer.length < signature.length) {
            return false;
        }
        for (let i = 0; i < signature.length; i++) {
            if (buffer[i] !== signature[i]) {
                return false;
            }
        }
        return true;
    }
    /**
     * Decodes an image from a Buffer.
     *
     * @param buffer The image data as a Buffer
     * @returns The decoded image with metadata
     * @throws Error if the image cannot be decoded
     */
    async decode(buffer) {
        const format = this.detectFormat(buffer);
        this.logger.debug('Decoding image', {
            format,
            size: buffer.length,
        });
        if (format === 'emf' || format === 'wmf') {
            return this.decodeMetafile(buffer, format);
        }
        if (format === 'tiff') {
            return this.decodeTiff(buffer);
        }
        try {
            // loadImage can take a Buffer directly
            const image = await loadImage(buffer);
            const result = {
                image,
                width: image.width,
                height: image.height,
                format,
            };
            this.logger.debug('Image decoded successfully', {
                format,
                width: image.width,
                height: image.height,
            });
            return result;
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to decode image', {
                format,
                size: buffer.length,
                error: message,
            });
            this.warnings?.push({
                code: 'image-decode-failed',
                message: `Failed to decode ${format} image (${buffer.length} bytes): ${message}`,
                detail: { format, size: buffer.length },
            });
            throw new Error(`Failed to decode image: ${message}`, { cause: error });
        }
    }
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
    async decodeMetafile(buffer, format) {
        try {
            const parsed = format === 'emf'
                ? new EmfParser({ logger: this.logger.child?.('EmfParser') }).parse(buffer)
                : new WmfParser({ logger: this.logger.child?.('WmfParser') }).parse(buffer);
            const renderer = new EmfRenderer({
                logger: this.logger.child?.('EmfRenderer'),
                warnings: this.warnings,
            });
            const canvas = await renderer.render(parsed);
            const png = await canvas.toBuffer('png');
            const image = await loadImage(png);
            this.logger.debug('Metafile rasterized', {
                format,
                width: image.width,
                height: image.height,
                records: parsed.records.length,
            });
            return {
                image,
                width: image.width,
                height: image.height,
                format,
            };
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.warn('Failed to rasterize metafile', {
                format,
                size: buffer.length,
                error: message,
            });
            this.warnings?.push({
                code: 'image-decode-failed',
                message: `Failed to rasterize ${format} metafile (${buffer.length} bytes): ${message}`,
                detail: { format, size: buffer.length },
            });
            throw new Error(`Failed to decode image: ${message}`, { cause: error });
        }
    }
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
    async decodeTiff(buffer) {
        let sharpModule;
        try {
            sharpModule = (await import('sharp')).default;
        }
        catch {
            this.warnings?.push({
                code: 'image-decode-failed',
                message: 'TIFF image requires the optional sharp dependency (npm install sharp)',
            });
            throw new Error('TIFF decoding requires the optional sharp dependency');
        }
        const png = await sharpModule(buffer).png().toBuffer();
        const image = await loadImage(png);
        this.logger.debug('TIFF transcoded via sharp', {
            width: image.width,
            height: image.height,
        });
        return { image, width: image.width, height: image.height, format: 'tiff' };
    }
    /**
     * Decodes an image from a data URI.
     *
     * @param dataUri The data URI string (e.g., data:image/png;base64,...)
     * @returns The decoded image with metadata
     */
    async decodeDataUri(dataUri) {
        // Parse data URI
        const match = dataUri.match(/^data:([^;]+);base64,(.+)$/);
        if (!match) {
            throw new Error('Invalid data URI format');
        }
        const mimeType = match[1];
        const base64Data = match[2];
        if (!base64Data) {
            throw new Error('No image data in data URI');
        }
        const buffer = Buffer.from(base64Data, 'base64');
        this.logger.debug('Decoding image from data URI', {
            mimeType,
            size: buffer.length,
        });
        return this.decode(buffer);
    }
    /**
     * Gets the MIME type for an image format.
     */
    getMimeType(format) {
        switch (format) {
            case 'png':
                return 'image/png';
            case 'jpeg':
                return 'image/jpeg';
            case 'gif':
                return 'image/gif';
            case 'bmp':
                return 'image/bmp';
            case 'webp':
                return 'image/webp';
            case 'tiff':
                return 'image/tiff';
            case 'emf':
                return 'image/x-emf';
            case 'wmf':
                return 'image/x-wmf';
            default:
                return 'application/octet-stream';
        }
    }
    /**
     * Gets the file extension for an image format.
     */
    getExtension(format) {
        switch (format) {
            case 'png':
                return '.png';
            case 'jpeg':
                return '.jpg';
            case 'gif':
                return '.gif';
            case 'bmp':
                return '.bmp';
            case 'webp':
                return '.webp';
            case 'tiff':
                return '.tiff';
            case 'emf':
                return '.emf';
            case 'wmf':
                return '.wmf';
            default:
                return '.bin';
        }
    }
}
/**
 * Creates an ImageDecoder instance.
 */
export function createImageDecoder(logger) {
    return new ImageDecoder({ logger });
}
//# sourceMappingURL=ImageDecoder.js.map