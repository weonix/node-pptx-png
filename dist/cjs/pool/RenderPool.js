"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.RenderPool = void 0;
exports.shardRoundRobin = shardRoundRobin;
exports.createRenderPool = createRenderPool;
/**
 * Multi-threaded render pool built on node:worker_threads.
 *
 * Slides are sharded round-robin across a set of persistent workers; each
 * worker opens the PPTX itself (file path, or bytes cloned per worker),
 * renders its shard through the normal PptxDocument path, and transfers
 * the encoded images back. The main thread reassembles the results in
 * request order, merging errors and structured warnings.
 */
const node_worker_threads_1 = require("node:worker_threads");
const fs = __importStar(require("node:fs"));
const os = __importStar(require("node:os"));
const path = __importStar(require("node:path"));
const node_url_1 = require("node:url");
const PptxDocument_js_1 = require("../core/PptxDocument.js");
const WarningCollector_js_1 = require("../utils/WarningCollector.js");
/**
 * Distributes items round-robin across `shards` buckets: item i goes to
 * bucket i % shards, preserving relative order within each bucket. Buckets
 * beyond the item count come back empty.
 */
function shardRoundRobin(items, shards) {
    if (!Number.isInteger(shards) || shards < 1) {
        throw new Error(`shardRoundRobin: invalid shard count ${shards}`);
    }
    const buckets = Array.from({ length: shards }, () => []);
    items.forEach((item, index) => {
        buckets[index % shards].push(item);
    });
    return buckets;
}
/**
 * Walks up from a directory until a package.json is found; returns the
 * containing directory, or undefined when none exists.
 */
function findPackageRoot(startDir) {
    let dir = startDir;
    for (;;) {
        if (fs.existsSync(path.join(dir, 'package.json'))) {
            return dir;
        }
        const parent = path.dirname(dir);
        if (parent === dir) {
            return undefined;
        }
        dir = parent;
    }
}
/**
 * A pool of worker threads that renders presentations in parallel.
 *
 * Create via {@link createRenderPool}. Workers are spawned lazily on the
 * first render and reused until {@link close} is called; a pool left open
 * keeps the process alive, so always close it when done.
 */
class RenderPool {
    targetWorkers;
    workerUrlOption;
    workers = [];
    workerUrlPromise = null;
    closePromise = null;
    nextJobId = 1;
    /**
     * Creates a pool. Prefer {@link createRenderPool}.
     *
     * @param options Worker count ('auto' sizes from the machine) and
     *   advanced worker-entry override
     */
    constructor(options = {}) {
        const workers = options.workers ?? 'auto';
        if (workers === 'auto') {
            this.targetWorkers = Math.max(1, os.availableParallelism() - 1);
        }
        else {
            if (!Number.isInteger(workers) || workers < 1) {
                throw new Error(`RenderPool: invalid workers option ${String(workers)}. ` +
                    `Expected a positive integer or 'auto'.`);
            }
            this.targetWorkers = workers;
        }
        this.workerUrlOption = options.workerUrl;
    }
    /** Maximum number of worker threads this pool will spawn. */
    get workerCount() {
        return this.targetWorkers;
    }
    /** Whether {@link close} has been called. */
    get closed() {
        return this.closePromise !== null;
    }
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
    async render(input, options = {}) {
        if (this.closed) {
            throw new Error('RenderPool is closed. Create a new pool with createRenderPool().');
        }
        const { slides, ...renderOptions } = options;
        try {
            // Open once on the main thread for validation and deck metadata
            // (slide count for sharding and the totalSlides aggregate)
            const document = await PptxDocument_js_1.PptxDocument.open(input, {
                logLevel: renderOptions.logLevel ?? 'warn',
            });
            let slideCount;
            try {
                slideCount = document.slideCount;
            }
            finally {
                document.close();
            }
            const selection = this.resolveSelection(slides ?? renderOptions.slideNumbers, slideCount).map((slideNumber, position) => ({ position, slideNumber }));
            if (selection.length === 0) {
                return { slides: [], totalSlides: slideCount, successfulSlides: 0, allSuccessful: true };
            }
            const shards = shardRoundRobin(selection, this.targetWorkers).filter((shard) => shard.length > 0);
            await this.ensureWorkers(shards.length);
            const responses = await Promise.all(shards.map((shard, index) => this.dispatch(this.workers[index], {
                id: this.nextJobId++,
                input,
                options: renderOptions,
                slideNumbers: shard.map((selected) => selected.slideNumber),
            })));
            return this.reassemble(selection, shards, responses, slideCount);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const stack = error instanceof Error ? error.stack : undefined;
            return {
                slides: [],
                totalSlides: 0,
                successfulSlides: 0,
                allSuccessful: false,
                errors: [{ level: 'presentation', message, stack }],
            };
        }
    }
    /**
     * Terminates all workers and rejects any in-flight jobs. Idempotent:
     * subsequent calls return the same completed promise. Render calls made
     * after close() reject with a clear error.
     */
    close() {
        if (this.closePromise) {
            return this.closePromise;
        }
        const entries = this.workers.splice(0, this.workers.length);
        for (const entry of entries) {
            for (const pending of entry.pending.values()) {
                pending.reject(new Error('RenderPool closed while a job was in flight'));
            }
            entry.pending.clear();
        }
        this.closePromise = Promise.all(entries.map((entry) => entry.worker.terminate())).then(() => undefined);
        return this.closePromise;
    }
    /**
     * Expands a selection into 1-based slide numbers: list form as-is,
     * range form clamped to the deck and expanded ascending, undefined
     * meaning all slides (mirrors PptxDocument.slides()).
     */
    resolveSelection(selection, slideCount) {
        if (Array.isArray(selection)) {
            return [...selection];
        }
        const from = Math.max(1, Math.floor(selection?.from ?? 1));
        const to = Math.min(slideCount, Math.floor(selection?.to ?? slideCount));
        const slideNumbers = [];
        for (let slideNumber = from; slideNumber <= to; slideNumber++) {
            slideNumbers.push(slideNumber);
        }
        return slideNumbers;
    }
    /**
     * Spawns workers until `count` are available (bounded by the pool's
     * configured size). Workers persist across render calls.
     */
    async ensureWorkers(count) {
        const url = await this.resolveWorkerUrl();
        const needed = Math.min(count, this.targetWorkers);
        while (this.workers.length < needed) {
            // execArgv: [] — workers must not inherit parent process module flags
            // (e.g. --input-type=module breaks worker file resolution)
            const entry = {
                worker: new node_worker_threads_1.Worker(url, { execArgv: [] }),
                pending: new Map(),
            };
            entry.worker.on('message', (response) => {
                const pending = entry.pending.get(response.id);
                if (pending) {
                    entry.pending.delete(response.id);
                    pending.resolve(response);
                }
            });
            const fail = (error) => {
                // The worker is unusable: drop it and fail its in-flight jobs
                const index = this.workers.indexOf(entry);
                if (index !== -1) {
                    this.workers.splice(index, 1);
                }
                for (const pending of entry.pending.values()) {
                    pending.reject(error);
                }
                entry.pending.clear();
            };
            entry.worker.on('error', fail);
            entry.worker.on('exit', (code) => {
                // ANY exit with jobs pending is a failure — a worker that exits
                // cleanly (code 0) without answering would otherwise hang the pool
                if (!this.closed) {
                    fail(new Error(`Render worker exited unexpectedly with code ${code}`));
                }
            });
            this.workers.push(entry);
        }
    }
    /**
     * Sends one job to a worker and resolves with its response.
     */
    dispatch(entry, job) {
        return new Promise((resolve, reject) => {
            entry.pending.set(job.id, { resolve, reject });
            try {
                entry.worker.postMessage(job);
            }
            catch (error) {
                entry.pending.delete(job.id);
                reject(error instanceof Error ? error : new Error(String(error)));
            }
        });
    }
    /**
     * Reassembles worker responses into a PresentationRenderResult in the
     * requested slide order, merging errors (sorted by slide index) and
     * deduplicating warnings with presentation-level ones first.
     */
    reassemble(selection, shards, responses, totalSlides) {
        const ordered = new Array(selection.length);
        const errors = [];
        const presentationWarnings = [];
        for (let shardIndex = 0; shardIndex < shards.length; shardIndex++) {
            const shard = shards[shardIndex];
            const response = responses[shardIndex];
            if (!response.ok) {
                // A whole shard failed (unreadable input, worker-side crash):
                // surface it as a presentation-level failure, matching the
                // renderPresentation contract for unreadable input
                throw new Error(`Render worker failed: ${response.message}`);
            }
            presentationWarnings.push(...response.presentationWarnings);
            errors.push(...response.errors);
            response.slides.forEach((payload, index) => {
                const selected = shard[index];
                if (selected) {
                    ordered[selected.position] = this.fromPayload(payload);
                }
            });
        }
        const slides = ordered.filter((slide) => slide !== undefined);
        const successfulSlides = slides.filter((slide) => slide.success).length;
        errors.sort((a, b) => (a.slideIndex ?? -1) - (b.slideIndex ?? -1));
        const warnings = (0, WarningCollector_js_1.dedupeWarnings)([
            ...presentationWarnings,
            ...slides.flatMap((slide) => slide.warnings ?? []),
        ]);
        return {
            slides,
            totalSlides,
            successfulSlides,
            allSuccessful: successfulSlides === slides.length && errors.length === 0,
            ...(errors.length > 0 ? { errors } : {}),
            ...(warnings.length > 0 ? { warnings } : {}),
        };
    }
    /**
     * Converts a worker slide payload back into a SlideRenderResult,
     * wrapping the transferred ArrayBuffer in a Buffer without copying.
     */
    fromPayload(payload) {
        return {
            slideIndex: payload.slideIndex,
            slideNumber: payload.slideNumber,
            imageData: Buffer.from(payload.imageData),
            width: payload.width,
            height: payload.height,
            success: payload.success,
            ...(payload.errorMessage !== undefined ? { errorMessage: payload.errorMessage } : {}),
            ...(payload.errorStack !== undefined ? { errorStack: payload.errorStack } : {}),
            ...(payload.warnings !== undefined ? { warnings: payload.warnings } : {}),
        };
    }
    /**
     * Locates the worker entry file, cached after the first call.
     */
    resolveWorkerUrl() {
        this.workerUrlPromise ??= this.findWorkerUrl();
        return this.workerUrlPromise;
    }
    /**
     * Resolves the worker entry file. Priority: explicit workerUrl option;
     * `worker.js` next to this module (dist/pool for the ESM build via
     * import.meta.url, dist/cjs/pool for the CJS build via __dirname); then
     * `dist/pool/worker.js` under the nearest package root (covers running
     * from src through a TypeScript-aware runner, where the sibling is .ts).
     */
    async findWorkerUrl() {
        if (this.workerUrlOption !== undefined) {
            if (this.workerUrlOption instanceof URL) {
                return this.workerUrlOption;
            }
            return this.workerUrlOption.startsWith('file:')
                ? new URL(this.workerUrlOption)
                : (0, node_url_1.pathToFileURL)(this.workerUrlOption);
        }
        const candidates = [];
        let moduleDir;
        // CJS build: __dirname is defined and worker.js sits next to this file
        if (typeof __dirname !== 'undefined') {
            moduleDir = __dirname;
        }
        else {
            // ESM build: the import.meta capture lives in workerUrl.js, whose
            // CJS transpile cannot be parsed — it is only imported on this path
            const { moduleUrl } = await Promise.resolve().then(() => __importStar(require('./workerUrl.js')));
            moduleDir = path.dirname((0, node_url_1.fileURLToPath)(moduleUrl));
        }
        candidates.push(path.join(moduleDir, 'worker.js'));
        // Dev/test fallback: running from src via a TS runner — use the built
        // ESM worker under the package's dist directory
        const packageRoot = findPackageRoot(moduleDir);
        if (packageRoot) {
            candidates.push(path.join(packageRoot, 'dist', 'pool', 'worker.js'));
        }
        for (const candidate of candidates) {
            if (fs.existsSync(candidate)) {
                return (0, node_url_1.pathToFileURL)(candidate);
            }
        }
        throw new Error(`RenderPool: cannot locate the worker entry file (checked: ${candidates.join(', ')}). ` +
            `Build the package (npm run build) or pass the workerUrl option.`);
    }
}
exports.RenderPool = RenderPool;
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
function createRenderPool(options = {}) {
    return new RenderPool(options);
}
