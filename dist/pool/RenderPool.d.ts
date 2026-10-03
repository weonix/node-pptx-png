import type { PptxRenderOptions, PresentationRenderResult, SlideSelection } from '../types/index.js';
/**
 * Options for {@link createRenderPool}.
 */
export interface RenderPoolOptions {
    /**
     * Number of worker threads, or 'auto' to size from the machine
     * (available parallelism minus one, at least 1).
     * @default 'auto'
     */
    workers?: number | 'auto';
    /**
     * Advanced: explicit path or file URL of the worker entry file.
     * Normally the pool locates the built worker next to its own compiled
     * module (dist/pool/worker.js or dist/cjs/pool/worker.js); pass this
     * when bundling relocates files.
     */
    workerUrl?: string | URL;
}
/**
 * Options for {@link RenderPool.render}: standard render options plus a
 * slide selection.
 */
export interface PoolRenderOptions extends PptxRenderOptions {
    /**
     * Slides to render, as an explicit list of 1-based slide numbers
     * (results keep the given order) or an inclusive `from`/`to` range
     * clamped to the deck. If omitted, `slideNumbers` is honored as a list;
     * otherwise all slides are rendered in deck order.
     */
    slides?: SlideSelection;
}
/**
 * Distributes items round-robin across `shards` buckets: item i goes to
 * bucket i % shards, preserving relative order within each bucket. Buckets
 * beyond the item count come back empty.
 */
export declare function shardRoundRobin<T>(items: readonly T[], shards: number): T[][];
/**
 * A pool of worker threads that renders presentations in parallel.
 *
 * Create via {@link createRenderPool}. Workers are spawned lazily on the
 * first render and reused until {@link close} is called; a pool left open
 * keeps the process alive, so always close it when done.
 */
export declare class RenderPool {
    private readonly targetWorkers;
    private readonly workerUrlOption;
    private readonly workers;
    private workerUrlPromise;
    private closePromise;
    private nextJobId;
    /**
     * Creates a pool. Prefer {@link createRenderPool}.
     *
     * @param options Worker count ('auto' sizes from the machine) and
     *   advanced worker-entry override
     */
    constructor(options?: RenderPoolOptions);
    /** Maximum number of worker threads this pool will spawn. */
    get workerCount(): number;
    /** Whether {@link close} has been called. */
    get closed(): boolean;
    /**
     * Renders a presentation across the pool's workers.
     *
     * Slide numbers are distributed round-robin across the workers; each
     * worker opens the input itself (a Buffer input is copied to each
     * worker), renders its shard, and transfers the encoded images back.
     * Results come back in the requested order with per-slide warnings
     * preserved and the aggregate matching `renderAll` semantics
     * (presentation-level warnings first, deduplicated).
     *
     * Like `renderPresentation`, this never rejects for input problems:
     * unreadable input or a failed worker resolves to a presentation-level
     * error result, and per-slide failures resolve to slide-level failure
     * results. It only rejects when the pool is already closed.
     *
     * @param input File path or Buffer containing PPTX data
     * @param options Render options plus a `slides` selection
     */
    render(input: string | Buffer, options?: PoolRenderOptions): Promise<PresentationRenderResult>;
    /**
     * Terminates all workers and rejects any in-flight jobs. Idempotent:
     * subsequent calls return the same completed promise. Render calls made
     * after close() reject with a clear error.
     */
    close(): Promise<void>;
    /**
     * Expands a selection into 1-based slide numbers: list form as-is,
     * range form clamped to the deck and expanded ascending, undefined
     * meaning all slides (mirrors PptxDocument.slides()).
     */
    private resolveSelection;
    /**
     * Spawns workers until `count` are available (bounded by the pool's
     * configured size). Workers persist across render calls.
     */
    private ensureWorkers;
    /**
     * Sends one job to a worker and resolves with its response.
     */
    private dispatch;
    /**
     * Reassembles worker responses into a PresentationRenderResult in the
     * requested slide order, merging errors (sorted by slide index) and
     * deduplicating warnings with presentation-level ones first.
     */
    private reassemble;
    /**
     * Converts a worker slide payload back into a SlideRenderResult,
     * wrapping the transferred ArrayBuffer in a Buffer without copying.
     */
    private fromPayload;
    /**
     * Locates the worker entry file, cached after the first call.
     */
    private resolveWorkerUrl;
    /**
     * Resolves the worker entry file. Priority: explicit workerUrl option;
     * `worker.js` next to this module (dist/pool for the ESM build via
     * import.meta.url, dist/cjs/pool for the CJS build via __dirname); then
     * `dist/pool/worker.js` under the nearest package root (covers running
     * from src through a TypeScript-aware runner, where the sibling is .ts).
     */
    private findWorkerUrl;
}
/**
 * Creates a {@link RenderPool} of worker threads for parallel rendering.
 * Workers spawn lazily on the first render; call {@link RenderPool.close}
 * when done or the workers will keep the process alive.
 *
 * @param options Worker count ('auto' by default) and advanced overrides
 * @example
 * ```typescript
 * const pool = createRenderPool({ workers: 4 });
 * try {
 *   const result = await pool.render('./deck.pptx', { preset: 'thumb' });
 * } finally {
 *   await pool.close();
 * }
 * ```
 */
export declare function createRenderPool(options?: RenderPoolOptions): RenderPool;
//# sourceMappingURL=RenderPool.d.ts.map