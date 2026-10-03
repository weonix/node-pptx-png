/**
 * Category of a structured render warning.
 *
 * Warnings report fidelity gaps that did not fail the render: content that
 * was skipped, substituted, or only partially drawn. Codes are stable
 * kebab-case identifiers suitable for CI gates and programmatic filtering.
 */
export type WarningCode = 'missing-font' | 'embedded-font-unsupported' | 'unsupported-element' | 'unsupported-geometry' | 'unsupported-chart-type' | 'chart-failed' | 'image-decode-failed' | 'metafile-partial' | 'diagram-missing-part' | 'layout-load-failed' | 'media-missing' | 'effect-skipped' | 'other';
/**
 * A structured warning emitted during rendering: a fidelity event (missing
 * font, unsupported element, decode failure, ...) that degraded the output
 * without failing it.
 */
export interface RenderWarning {
    /** Stable category code for programmatic handling. */
    code: WarningCode;
    /** Human-readable description of what was skipped or degraded. */
    message: string;
    /**
     * One-based slide number the warning occurred on. Undefined for
     * presentation-level warnings (e.g., embedded font registration, which
     * happens once per document rather than per slide).
     */
    slideNumber?: number;
    /** Identifier of the affected element (p:cNvPr id), when known. */
    elementId?: string;
    /** Machine-readable context (e.g., font family, graphic URI). */
    detail?: Record<string, unknown>;
}
/**
 * Result of rendering a single slide.
 */
export interface SlideRenderResult {
    /**
     * Zero-based slide index.
     */
    slideIndex: number;
    /**
     * One-based slide number for display purposes.
     */
    slideNumber: number;
    /**
     * Rendered image data as a Buffer. The encoding follows the `format`
     * render option: PNG (image/png), JPEG (image/jpeg), WebP (image/webp),
     * or SVG (image/svg+xml; the buffer holds the UTF-8 SVG document).
     */
    imageData: Buffer;
    /**
     * Rendered image width in pixels.
     */
    width: number;
    /**
     * Rendered image height in pixels.
     */
    height: number;
    /**
     * Whether the slide was rendered successfully.
     */
    success: boolean;
    /**
     * Error message if rendering failed.
     */
    errorMessage?: string;
    /**
     * Detailed error stack if available.
     */
    errorStack?: string;
    /**
     * Structured warnings collected while rendering this slide (fidelity
     * events such as missing fonts or unsupported elements). Present only
     * when at least one warning was emitted; a clean render omits the field.
     */
    warnings?: RenderWarning[];
}
/**
 * Result of rendering an entire presentation.
 */
export interface PresentationRenderResult {
    /**
     * Results for each slide in order.
     */
    slides: SlideRenderResult[];
    /**
     * Total number of slides in the presentation.
     */
    totalSlides: number;
    /**
     * Number of slides that rendered successfully.
     */
    successfulSlides: number;
    /**
     * Whether all slides rendered successfully.
     */
    allSuccessful: boolean;
    /**
     * Errors encountered during rendering.
     * Contains one `slide`-level entry per failed or invalid slide and a
     * `presentation`-level entry when the presentation itself could not be
     * processed (e.g., invalid PPTX file). Omitted when no errors occurred.
     */
    errors?: RenderError[];
    /**
     * Aggregate of all structured warnings: presentation-level warnings
     * (slideNumber undefined, e.g. embedded-font registration) plus every
     * per-slide warning, deduplicated by code + message + slideNumber.
     * Present only when at least one warning was emitted.
     */
    warnings?: RenderWarning[];
}
/**
 * Level at which a rendering error occurred.
 */
export type RenderErrorLevel = 'presentation' | 'slide' | 'element';
/**
 * Detailed error information for rendering failures.
 */
export interface RenderError {
    /**
     * Level at which the error occurred.
     */
    level: RenderErrorLevel;
    /**
     * Slide index if applicable (zero-based).
     */
    slideIndex?: number;
    /**
     * Type of element that caused the error.
     */
    elementType?: string;
    /**
     * Element ID or identifier if available.
     */
    elementId?: string;
    /**
     * Human-readable error message.
     */
    message: string;
    /**
     * Stack trace if available.
     */
    stack?: string;
}
