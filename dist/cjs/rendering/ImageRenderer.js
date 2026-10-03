"use strict";
/**
 * Renders images to canvas context.
 * Handles picture shapes (p:pic elements), blipFill for shape fills,
 * source rectangle cropping, and stretch/tile fill modes.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ImageRenderer = void 0;
exports.computeStretchDestRect = computeStretchDestRect;
exports.createImageRenderer = createImageRenderer;
const PptxParser_js_1 = require("../core/PptxParser.js");
const RelationshipParser_js_1 = require("../parsers/RelationshipParser.js");
const ImageDecoder_js_1 = require("../utils/ImageDecoder.js");
const TransformCalculator_js_1 = require("../geometry/TransformCalculator.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Computes the destination rectangle for a stretched image with an a:fillRect.
 * Per ECMA-376 (20.1.8.30), positive fillRect percentages are INSETS — the
 * image is drawn inside a rectangle smaller than the shape bounds — while
 * negative percentages are outsets that extend beyond the bounds.
 * @param bounds The shape bounds in pixels
 * @param fillRect The fill rectangle percentages (100000 = 100%)
 * @returns The destination rectangle in pixels
 */
function computeStretchDestRect(bounds, fillRect) {
    const leftPct = fillRect.left / 100000;
    const topPct = fillRect.top / 100000;
    const rightPct = fillRect.right / 100000;
    const bottomPct = fillRect.bottom / 100000;
    return {
        x: bounds.x + bounds.width * leftPct,
        y: bounds.y + bounds.height * topPct,
        width: bounds.width * (1 - leftPct - rightPct),
        height: bounds.height * (1 - topPct - bottomPct),
    };
}
/**
 * Renders images from PPTX media folder to canvas.
 */
class ImageRenderer {
    logger;
    parser;
    sourcePath;
    scaleX;
    scaleY;
    relationshipParser;
    imageDecoder;
    transformCalculator;
    unitConverter;
    /** Cache of loaded images, keyed by relationship ID */
    imageCache = new Map();
    /** Maximum number of images to cache (LRU eviction) */
    maxCacheSize = 50;
    /** Track insertion order for LRU eviction */
    cacheOrder = [];
    warnings;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'ImageRenderer');
        this.parser = config.parser;
        this.sourcePath = config.sourcePath;
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
        this.warnings = config.warnings;
        this.relationshipParser = new RelationshipParser_js_1.RelationshipParser({
            parser: config.parser,
            logger: this.logger.child?.('RelParser'),
        });
        this.imageDecoder = new ImageDecoder_js_1.ImageDecoder({
            logger: this.logger.child?.('Decoder'),
            warnings: config.warnings,
        });
        this.transformCalculator = new TransformCalculator_js_1.TransformCalculator();
        this.unitConverter = new UnitConverter_js_1.UnitConverter();
    }
    /**
     * Adds an image to the cache with LRU eviction.
     * @param key The cache key (relationship ID)
     * @param image The decoded image to cache
     */
    addToCache(key, image) {
        if (this.imageCache.has(key)) {
            // Move to end of order (most recently used)
            const idx = this.cacheOrder.indexOf(key);
            if (idx > -1) {
                this.cacheOrder.splice(idx, 1);
            }
        }
        this.imageCache.set(key, image);
        this.cacheOrder.push(key);
        // Evict oldest entries if over limit
        while (this.cacheOrder.length > this.maxCacheSize) {
            const oldest = this.cacheOrder.shift();
            if (oldest) {
                this.imageCache.delete(oldest);
                this.logger.debug('Evicted image from cache', { id: oldest });
            }
        }
    }
    /**
     * Loads an image by its relationship ID.
     *
     * @param relationshipId The r:embed relationship ID
     * @returns The decoded image or undefined if not found
     */
    async loadImage(relationshipId) {
        // Check cache (and update access order for LRU)
        const cached = this.imageCache.get(relationshipId);
        if (cached) {
            // Move to end of order (most recently used)
            const idx = this.cacheOrder.indexOf(relationshipId);
            if (idx > -1) {
                this.cacheOrder.splice(idx, 1);
                this.cacheOrder.push(relationshipId);
            }
            return cached;
        }
        try {
            // Resolve relationship to media path
            const mediaPath = await this.relationshipParser.resolveImageRelationship(this.sourcePath, relationshipId);
            if (!mediaPath) {
                this.logger.warn('Image relationship not found', { id: relationshipId });
                this.warnings?.push({
                    code: 'media-missing',
                    message: `Image relationship "${relationshipId}" not found; picture skipped`,
                    detail: { relationshipId, sourcePath: this.sourcePath },
                });
                return undefined;
            }
            // Load the image data from the PPTX
            const buffer = await this.parser.readBinary(mediaPath);
            // Decode the image
            const decoded = await this.imageDecoder.decode(buffer);
            // Cache it with LRU eviction
            this.addToCache(relationshipId, decoded);
            this.logger.debug('Image loaded', {
                id: relationshipId,
                path: mediaPath,
                width: decoded.width,
                height: decoded.height,
                format: decoded.format,
            });
            return decoded;
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to load image', {
                id: relationshipId,
                error: message,
            });
            // Decode failures are already collected by ImageDecoder as
            // 'image-decode-failed'; anything else here is unreadable media.
            if (!message.startsWith('Failed to decode image')) {
                this.warnings?.push({
                    code: 'media-missing',
                    message: `Failed to load image "${relationshipId}": ${message}`,
                    detail: { relationshipId, sourcePath: this.sourcePath },
                });
            }
            return undefined;
        }
    }
    /**
     * Parses a blipFill element to extract picture data.
     *
     * @param blipFill The a:blipFill element node
     * @returns Parsed picture data or undefined
     */
    parseBlipFill(blipFill) {
        if (!blipFill) {
            return undefined;
        }
        // Get the blip element with the relationship ID
        const blip = (0, PptxParser_js_1.getXmlChild)(blipFill, 'a:blip');
        if (!blip) {
            this.logger.debug('blipFill has no blip element');
            return undefined;
        }
        const blipRelId = (0, PptxParser_js_1.getXmlAttr)(blip, 'r:embed');
        if (!blipRelId) {
            // Could be r:link for external images, but we only support embedded
            this.logger.debug('blip has no r:embed attribute');
            return undefined;
        }
        const result = {
            blipRelId,
        };
        // Picture-level transparency (a:alphaModFix amt in 1000ths of a percent)
        const alphaModFix = (0, PptxParser_js_1.getXmlChild)(blip, 'a:alphaModFix');
        if (alphaModFix) {
            const amt = (0, PptxParser_js_1.getXmlAttr)(alphaModFix, 'amt');
            if (amt !== undefined) {
                const parsed = parseInt(amt, 10);
                if (Number.isFinite(parsed)) {
                    result.alpha = Math.min(1, Math.max(0, parsed / 100000));
                }
            }
        }
        // Parse source rectangle (cropping)
        const srcRect = (0, PptxParser_js_1.getXmlChild)(blipFill, 'a:srcRect');
        if (srcRect) {
            result.srcRect = this.parseCropRect(srcRect);
        }
        // Check for stretch fill
        const stretch = (0, PptxParser_js_1.getXmlChild)(blipFill, 'a:stretch');
        if (stretch) {
            result.stretch = true;
            // Parse fill rectangle within stretch
            const fillRect = (0, PptxParser_js_1.getXmlChild)(stretch, 'a:fillRect');
            if (fillRect) {
                result.fillRect = this.parseCropRect(fillRect);
            }
        }
        // Check for tile fill
        const tile = (0, PptxParser_js_1.getXmlChild)(blipFill, 'a:tile');
        if (tile) {
            result.tile = this.parseTileInfo(tile);
        }
        return result;
    }
    /**
     * Parses a p:pic element to extract picture data.
     *
     * @param picNode The p:pic element node
     * @returns Parsed picture data or undefined
     */
    parsePicElement(picNode) {
        if (!picNode) {
            return undefined;
        }
        // Get blipFill from the picture element
        const blipFill = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:blipFill');
        return this.parseBlipFill(blipFill);
    }
    /**
     * Parses a crop rectangle element (srcRect or fillRect).
     * Values are in OpenXML percentage format (0-100000).
     */
    parseCropRect(rectNode) {
        return {
            left: parseInt((0, PptxParser_js_1.getXmlAttr)(rectNode, 'l') ?? '0', 10),
            top: parseInt((0, PptxParser_js_1.getXmlAttr)(rectNode, 't') ?? '0', 10),
            right: parseInt((0, PptxParser_js_1.getXmlAttr)(rectNode, 'r') ?? '0', 10),
            bottom: parseInt((0, PptxParser_js_1.getXmlAttr)(rectNode, 'b') ?? '0', 10),
        };
    }
    /**
     * Parses tile fill settings.
     */
    parseTileInfo(tileNode) {
        const flipAttr = (0, PptxParser_js_1.getXmlAttr)(tileNode, 'flip') ?? 'none';
        let flip = 'none';
        if (flipAttr === 'x')
            flip = 'x';
        else if (flipAttr === 'y')
            flip = 'y';
        else if (flipAttr === 'xy')
            flip = 'xy';
        return {
            sx: parseInt((0, PptxParser_js_1.getXmlAttr)(tileNode, 'sx') ?? '100000', 10),
            sy: parseInt((0, PptxParser_js_1.getXmlAttr)(tileNode, 'sy') ?? '100000', 10),
            tx: parseInt((0, PptxParser_js_1.getXmlAttr)(tileNode, 'tx') ?? '0', 10),
            ty: parseInt((0, PptxParser_js_1.getXmlAttr)(tileNode, 'ty') ?? '0', 10),
            flip,
            alignment: (0, PptxParser_js_1.getXmlAttr)(tileNode, 'algn') ?? 'tl',
        };
    }
    /**
     * Converts a percentage value (0-100000) to a decimal (0-1).
     */
    percentToDecimal(percent) {
        return percent / 100000;
    }
    /**
     * Renders an image to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param image The decoded image to render
     * @param bounds The destination bounds in pixels
     * @param pictureData Optional picture data with cropping/fill settings
     */
    renderImage(ctx, image, bounds, pictureData) {
        ctx.save();
        // Picture-level transparency (a:alphaModFix)
        if (pictureData?.alpha !== undefined && pictureData.alpha < 1) {
            ctx.globalAlpha = ctx.globalAlpha * pictureData.alpha;
        }
        // Calculate source rectangle (cropping)
        let srcX = 0;
        let srcY = 0;
        let srcWidth = image.width;
        let srcHeight = image.height;
        if (pictureData?.srcRect) {
            const { left, top, right, bottom } = pictureData.srcRect;
            const leftPct = this.percentToDecimal(left);
            const topPct = this.percentToDecimal(top);
            const rightPct = this.percentToDecimal(right);
            const bottomPct = this.percentToDecimal(bottom);
            srcX = image.width * leftPct;
            srcY = image.height * topPct;
            srcWidth = image.width * (1 - leftPct - rightPct);
            srcHeight = image.height * (1 - topPct - bottomPct);
        }
        // Calculate destination rectangle
        let destX = bounds.x;
        let destY = bounds.y;
        let destWidth = bounds.width;
        let destHeight = bounds.height;
        // Apply fill rectangle if specified (positive values inset the destination
        // within bounds, negative values outset it — ECMA-376 20.1.8.30)
        if (pictureData?.fillRect) {
            const dest = computeStretchDestRect(bounds, pictureData.fillRect);
            destX = dest.x;
            destY = dest.y;
            destWidth = dest.width;
            destHeight = dest.height;
        }
        if (pictureData?.tile) {
            // Render as tiled pattern
            this.renderTiledImage(ctx, image, bounds, pictureData.tile, srcX, srcY, srcWidth, srcHeight);
        }
        else {
            // Render stretched or scaled image
            ctx.drawImage(image, srcX, srcY, srcWidth, srcHeight, destX, destY, destWidth, destHeight);
        }
        ctx.restore();
        this.logger.debug('Image rendered', {
            src: { x: srcX, y: srcY, width: srcWidth, height: srcHeight },
            dest: { x: destX, y: destY, width: destWidth, height: destHeight },
            tiled: !!pictureData?.tile,
        });
    }
    /**
     * Renders a tiled image pattern.
     */
    renderTiledImage(ctx, image, bounds, tile, srcX, srcY, srcWidth, srcHeight) {
        // Calculate tile dimensions in output pixels: source pixels scaled by the
        // tile sx/sy percentages, then by the render scale like all other geometry
        const tileWidth = ((srcWidth * tile.sx) / 100000) * this.scaleX;
        const tileHeight = ((srcHeight * tile.sy) / 100000) * this.scaleY;
        // Guard against infinite loops from zero/negative tile dimensions
        if (tileWidth <= 0 || tileHeight <= 0) {
            this.logger.warn('Invalid tile dimensions, skipping tile fill', {
                tileWidth,
                tileHeight,
                srcWidth,
                srcHeight,
                scaleX: tile.sx,
                scaleY: tile.sy,
            });
            return;
        }
        // Calculate offset in pixels (from EMU)
        const offsetX = this.unitConverter.emuToPixels(tile.tx) * this.scaleX;
        const offsetY = this.unitConverter.emuToPixels(tile.ty) * this.scaleY;
        // Set clip region to bounds
        ctx.beginPath();
        ctx.rect(bounds.x, bounds.y, bounds.width, bounds.height);
        ctx.clip();
        // Calculate starting position based on alignment
        let startX = bounds.x + offsetX;
        let startY = bounds.y + offsetY;
        // Adjust for alignment (tl = top-left, etc.)
        switch (tile.alignment) {
            case 'tr':
                startX = bounds.x + bounds.width - tileWidth + offsetX;
                break;
            case 'bl':
                startY = bounds.y + bounds.height - tileHeight + offsetY;
                break;
            case 'br':
                startX = bounds.x + bounds.width - tileWidth + offsetX;
                startY = bounds.y + bounds.height - tileHeight + offsetY;
                break;
            case 'ctr':
                startX = bounds.x + (bounds.width - tileWidth) / 2 + offsetX;
                startY = bounds.y + (bounds.height - tileHeight) / 2 + offsetY;
                break;
            // tl and others default to top-left
        }
        // Adjust start to ensure we cover from the beginning
        while (startX > bounds.x)
            startX -= tileWidth;
        while (startY > bounds.y)
            startY -= tileHeight;
        // Draw tiles
        let row = 0;
        for (let y = startY; y < bounds.y + bounds.height; y += tileHeight) {
            let col = 0;
            for (let x = startX; x < bounds.x + bounds.width; x += tileWidth) {
                ctx.save();
                // Apply flip if needed
                let flipX = 1;
                let flipY = 1;
                if (tile.flip === 'x' || tile.flip === 'xy') {
                    flipX = col % 2 === 0 ? 1 : -1;
                }
                if (tile.flip === 'y' || tile.flip === 'xy') {
                    flipY = row % 2 === 0 ? 1 : -1;
                }
                if (flipX !== 1 || flipY !== 1) {
                    ctx.translate(x + tileWidth / 2, y + tileHeight / 2);
                    ctx.scale(flipX, flipY);
                    ctx.translate(-(x + tileWidth / 2), -(y + tileHeight / 2));
                }
                ctx.drawImage(image, srcX, srcY, srcWidth, srcHeight, x, y, tileWidth, tileHeight);
                ctx.restore();
                col++;
            }
            row++;
        }
    }
    /**
     * Renders a picture element (p:pic) to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param picNode The p:pic XML node
     * @param transform The pixel transform for the picture
     */
    async renderPictureElement(ctx, picNode, transform) {
        // Skip hidden pictures (p:nvPicPr/p:cNvPr @hidden — e.g. OLE previews)
        const nvPicPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:nvPicPr');
        const cNvPr = (0, PptxParser_js_1.getXmlChild)(nvPicPr, 'p:cNvPr');
        if ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1') {
            this.logger.debug('Skipping hidden picture');
            return;
        }
        // Parse the picture data
        const pictureData = this.parsePicElement(picNode);
        if (!pictureData) {
            this.logger.debug('No picture data found in p:pic element');
            return;
        }
        // Load the image
        const decodedImage = await this.loadImage(pictureData.blipRelId);
        if (!decodedImage) {
            this.logger.warn('Could not load image', { relId: pictureData.blipRelId });
            return;
        }
        // Apply transform
        ctx.save();
        this.transformCalculator.applyTransform(ctx, transform);
        // Create bounds at origin (transform already applied)
        const bounds = {
            x: 0,
            y: 0,
            width: transform.width,
            height: transform.height,
        };
        // Render the image
        this.renderImage(ctx, decodedImage.image, bounds, pictureData);
        ctx.restore();
    }
    /**
     * Renders a picture fill (blipFill) to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param blipFill The a:blipFill XML node
     * @param bounds The destination bounds in pixels
     */
    async renderPictureFill(ctx, blipFill, bounds) {
        // Parse the picture data
        const pictureData = this.parseBlipFill(blipFill);
        if (!pictureData) {
            this.logger.debug('No picture data found in blipFill');
            return;
        }
        // Load the image
        const decodedImage = await this.loadImage(pictureData.blipRelId);
        if (!decodedImage) {
            this.logger.warn('Could not load image for fill', { relId: pictureData.blipRelId });
            return;
        }
        // Render the image
        this.renderImage(ctx, decodedImage.image, bounds, pictureData);
    }
    /**
     * Clears the image cache.
     */
    clearCache() {
        this.imageCache.clear();
        this.cacheOrder.length = 0;
        this.logger.debug('Image cache cleared');
    }
    /**
     * Updates the source path for relationship resolution.
     * Call this when rendering a different slide.
     */
    setSourcePath(sourcePath) {
        return new ImageRenderer({
            parser: this.parser,
            sourcePath,
            scaleX: this.scaleX,
            scaleY: this.scaleY,
            logger: this.logger,
        });
    }
}
exports.ImageRenderer = ImageRenderer;
/**
 * Creates an ImageRenderer instance.
 */
function createImageRenderer(parser, sourcePath, scaleX, scaleY, logger) {
    return new ImageRenderer({
        parser,
        sourcePath,
        scaleX,
        scaleY,
        logger,
    });
}
