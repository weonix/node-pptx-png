import type { PptxRenderOptions, SlideRenderResult, PresentationRenderResult } from '../types/index.js';
/**
 * Interface for the PPTX image renderer.
 */
export interface IPptxImageRenderer {
    /**
     * Renders all slides in a presentation to images.
     */
    renderPresentation(input: Buffer | string, options?: PptxRenderOptions): Promise<PresentationRenderResult>;
    /**
     * Renders a single slide to an image.
     */
    renderSlide(input: Buffer | string, slideIndex: number, options?: PptxRenderOptions): Promise<SlideRenderResult>;
    /**
     * Gets the number of slides in a presentation.
     */
    getSlideCount(input: Buffer | string): Promise<number>;
    /**
     * Gets the slide dimensions in EMU.
     */
    getSlideDimensions(input: Buffer | string): Promise<{
        width: number;
        height: number;
    }>;
}
/**
 * Main entry point for rendering PPTX presentations to images.
 *
 * Each call opens, renders, and closes the file. All calls delegate to
 * {@link PptxDocument} internally, so this class and the document API share
 * a single render path; when rendering the same file repeatedly, prefer
 * `openPresentation()` to parse it once and reuse the document.
 */
export declare class PptxImageRenderer implements IPptxImageRenderer {
    private readonly logger;
    constructor(options?: {
        logLevel?: PptxRenderOptions['logLevel'];
    });
    /**
     * Renders all slides in a presentation to images.
     *
     * Never rejects: invalid input resolves to a presentation-level error
     * result, and per-slide failures resolve to slide-level failure results.
     */
    renderPresentation(input: Buffer | string, options?: PptxRenderOptions): Promise<PresentationRenderResult>;
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
    renderSlide(input: Buffer | string, slideIndex: number, options?: PptxRenderOptions): Promise<SlideRenderResult>;
    /**
     * Gets the number of slides in a presentation.
     */
    getSlideCount(input: Buffer | string): Promise<number>;
    /**
     * Gets the slide dimensions in EMU.
     */
    getSlideDimensions(input: Buffer | string): Promise<{
        width: number;
        height: number;
    }>;
}
/**
 * Creates a new PptxImageRenderer instance.
 */
export declare function createRenderer(options?: {
    logLevel?: PptxRenderOptions['logLevel'];
}): IPptxImageRenderer;
/**
 * Convenience function to render a presentation.
 */
export declare function renderPresentation(input: Buffer | string, options?: PptxRenderOptions): Promise<PresentationRenderResult>;
/**
 * Convenience function to render a single slide.
 */
export declare function renderSlide(input: Buffer | string, slideIndex: number, options?: PptxRenderOptions): Promise<SlideRenderResult>;
/**
 * Convenience function to get the number of slides in a presentation.
 */
export declare function getSlideCount(input: Buffer | string): Promise<number>;
/**
 * Convenience function to get the native slide dimensions in EMU.
 */
export declare function getSlideDimensions(input: Buffer | string): Promise<{
    width: number;
    height: number;
}>;
//# sourceMappingURL=PptxImageRenderer.d.ts.map