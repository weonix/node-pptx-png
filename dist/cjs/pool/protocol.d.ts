/**
 * Message protocol shared between RenderPool (main thread) and the pool
 * worker entry. All payloads are structured-clone-safe; image bytes travel
 * as transferable ArrayBuffers so they move between threads without
 * re-serialization.
 */
import type { PptxRenderOptions, RenderError, RenderWarning } from '../types/index.js';
/**
 * A render job sent from the pool to one worker: the worker opens the
 * input itself and renders its shard of slide numbers.
 */
export interface PoolRenderJob {
    /** Pool-unique job id, echoed back in the response. */
    id: number;
    /**
     * PPTX source: a file path, or the file bytes (structured-clone copied
     * per worker; Buffers arrive as Uint8Array on the worker side).
     */
    input: string | Uint8Array;
    /** Render options (structured-clone-safe subset; no callbacks/signals). */
    options: PptxRenderOptions;
    /** 1-based slide numbers this worker renders, in the order to render. */
    slideNumbers: number[];
}
/**
 * One rendered slide as sent back from a worker. Mirrors SlideRenderResult
 * except that imageData is a transferable ArrayBuffer instead of a Buffer.
 */
export interface PoolSlidePayload {
    /** Zero-based slide index. */
    slideIndex: number;
    /** One-based slide number. */
    slideNumber: number;
    /** Encoded image bytes (transferred, not copied). */
    imageData: ArrayBuffer;
    /** Rendered width in pixels. */
    width: number;
    /** Rendered height in pixels. */
    height: number;
    /** Whether the slide rendered successfully. */
    success: boolean;
    /** Error message if rendering failed. */
    errorMessage?: string;
    /** Error stack if available. */
    errorStack?: string;
    /** Structured warnings collected for this slide. */
    warnings?: RenderWarning[];
}
/**
 * Successful worker response: the worker's shard rendered (individual
 * slides may still have failed; see per-slide `success`).
 */
export interface PoolSuccessResponse {
    /** Job id this response answers. */
    id: number;
    /** Discriminant: the shard was processed. */
    ok: true;
    /** Results in the same order as the job's slideNumbers. */
    slides: PoolSlidePayload[];
    /** Total number of slides in the deck as seen by this worker. */
    totalSlides: number;
    /** Presentation-level warnings (slideNumber undefined) from this worker. */
    presentationWarnings: RenderWarning[];
    /** Slide/presentation-level errors reported by the render. */
    errors: RenderError[];
}
/**
 * Failed worker response: the worker could not process the job at all
 * (e.g. the input is not a readable PPTX).
 */
export interface PoolErrorResponse {
    /** Job id this response answers. */
    id: number;
    /** Discriminant: the job failed wholesale. */
    ok: false;
    /** Human-readable failure message. */
    message: string;
    /** Stack trace if available. */
    stack?: string;
}
/** Any response a worker can post for a job. */
export type PoolResponse = PoolSuccessResponse | PoolErrorResponse;
