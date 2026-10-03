"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PptxImageRenderer = void 0;
exports.createRenderer = createRenderer;
exports.renderPresentation = renderPresentation;
exports.renderSlide = renderSlide;
exports.getSlideCount = getSlideCount;
exports.getSlideDimensions = getSlideDimensions;
const PptxDocument_js_1 = require("./PptxDocument.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Main entry point for rendering PPTX presentations to images.
 *
 * Each call opens, renders, and closes the file. All calls delegate to
 * {@link PptxDocument} internally, so this class and the document API share
 * a single render path; when rendering the same file repeatedly, prefer
 * `openPresentation()` to parse it once and reuse the document.
 */
class PptxImageRenderer {
    logger;
    constructor(options) {
        this.logger = (0, Logger_js_1.createLogger)(options?.logLevel ?? 'warn', 'PptxImageRenderer');
    }
    /**
     * Renders all slides in a presentation to images.
     *
     * Never rejects: invalid input resolves to a presentation-level error
     * result, and per-slide failures resolve to slide-level failure results.
     */
    async renderPresentation(input, options = {}) {
        let document;
        try {
            document = await PptxDocument_js_1.PptxDocument.open(input, {}, this.logger);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const stack = error instanceof Error ? error.stack : undefined;
            this.logger.error('Failed to render presentation', { error: message });
            return {
                slides: [],
                totalSlides: 0,
                successfulSlides: 0,
                allSuccessful: false,
                errors: [{ level: 'presentation', message, stack }],
            };
        }
        try {
            return await document.renderAll(options);
        }
        finally {
            document.close();
        }
    }
    /**
     * Renders a single slide to an image.
     *
     * Note: `slideIndex` is **0-based** (legacy contract). The PptxDocument
     * API uses 1-based slide numbers instead, consistent with the
     * `slideNumbers` render option.
     *
     * Rejects on unreadable input; bad per-call input (out-of-range index,
     * invalid scale or preset) resolves to a result with `success: false`.
     */
    async renderSlide(input, slideIndex, options = {}) {
        const document = await PptxDocument_js_1.PptxDocument.open(input, {}, this.logger);
        try {
            // Validate here to preserve this method's legacy 0-based error message
            if (slideIndex < 0 || slideIndex >= document.slideCount) {
                return {
                    slideIndex,
                    slideNumber: slideIndex + 1,
                    imageData: Buffer.alloc(0),
                    width: 0,
                    height: 0,
                    success: false,
                    errorMessage: `Slide index ${slideIndex} out of range (0-${document.slideCount - 1})`,
                };
            }
            // The await is required: the finally block closes the document as soon
            // as this method returns, so returning the pending promise directly
            // would close it before rendering runs.
            return await document.slide(slideIndex + 1, options);
        }
        finally {
            document.close();
        }
    }
    /**
     * Gets the number of slides in a presentation.
     */
    async getSlideCount(input) {
        const document = await PptxDocument_js_1.PptxDocument.open(input, {}, this.logger);
        try {
            return document.slideCount;
        }
        finally {
            document.close();
        }
    }
    /**
     * Gets the slide dimensions in EMU.
     */
    async getSlideDimensions(input) {
        const document = await PptxDocument_js_1.PptxDocument.open(input, {}, this.logger);
        try {
            return {
                width: document.size.widthEmu,
                height: document.size.heightEmu,
            };
        }
        finally {
            document.close();
        }
    }
}
exports.PptxImageRenderer = PptxImageRenderer;
/**
 * Creates a new PptxImageRenderer instance.
 */
function createRenderer(options) {
    return new PptxImageRenderer(options);
}
/**
 * Convenience function to render a presentation.
 */
async function renderPresentation(input, options) {
    const renderer = new PptxImageRenderer({ logLevel: options?.logLevel });
    return renderer.renderPresentation(input, options);
}
/**
 * Convenience function to render a single slide.
 */
async function renderSlide(input, slideIndex, options) {
    const renderer = new PptxImageRenderer({ logLevel: options?.logLevel });
    return renderer.renderSlide(input, slideIndex, options);
}
/**
 * Convenience function to get the number of slides in a presentation.
 */
async function getSlideCount(input) {
    return new PptxImageRenderer().getSlideCount(input);
}
/**
 * Convenience function to get the native slide dimensions in EMU.
 */
async function getSlideDimensions(input) {
    return new PptxImageRenderer().getSlideDimensions(input);
}
