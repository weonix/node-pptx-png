/**
 * Bounded collector for structured render warnings.
 *
 * One collector is created per slide-render invocation (so warnings carry
 * per-slide attribution) plus one per document for presentation-level
 * events such as embedded-font registration. Renderers receive the
 * collector through their configs and push warnings alongside their
 * existing log output — collecting never replaces logging.
 */
import type { RenderWarning } from '../types/results.js';
/**
 * Default maximum number of warnings a collector stores. Pathological decks
 * (e.g., thousands of shapes with unknown geometry) would otherwise grow
 * the warning list without bound.
 */
export declare const DEFAULT_WARNING_LIMIT = 500;
/**
 * Options for creating a WarningCollector.
 */
export interface WarningCollectorOptions {
    /**
     * One-based slide number stamped onto pushed warnings that do not carry
     * their own. Omit for presentation-level collectors, whose warnings keep
     * slideNumber undefined.
     */
    slideNumber?: number;
    /** Maximum stored warnings before further pushes are counted as dropped. */
    limit?: number;
}
/**
 * Collects structured render warnings with a storage cap.
 *
 * Warnings pushed beyond the cap are not stored; they increment
 * {@link droppedCount} instead so callers can detect truncation.
 */
export declare class WarningCollector {
    private readonly warnings;
    private readonly slideNumber;
    private readonly limit;
    private dropped;
    constructor(options?: WarningCollectorOptions);
    /**
     * Records a warning. When the collector was created with a slide number,
     * warnings without their own slideNumber are attributed to that slide.
     * Pushes beyond the storage cap are counted as dropped.
     */
    push(warning: RenderWarning): void;
    /**
     * Returns the collected warnings (a copy, in push order).
     */
    list(): RenderWarning[];
    /**
     * Number of warnings dropped because the storage cap was reached.
     */
    get droppedCount(): number;
}
/**
 * Deduplicates warnings by code + message + slideNumber, preserving the
 * first occurrence's order. Used to build the presentation-level aggregate,
 * where the same fidelity event repeated across shapes or render calls
 * would otherwise drown the signal.
 */
export declare function dedupeWarnings(warnings: RenderWarning[]): RenderWarning[];
