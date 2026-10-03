"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Worker entry for RenderPool (node:worker_threads).
 *
 * Each worker opens the PPTX input itself (file path, or bytes cloned from
 * the main thread), renders its shard of slides via the normal
 * PptxDocument path, and posts the encoded images back as transferable
 * ArrayBuffers. The file is loadable from both the ESM build
 * (dist/pool/worker.js) and the CJS build (dist/cjs/pool/worker.js): it
 * contains no import.meta and no top-level await.
 */
const node_worker_threads_1 = require("node:worker_threads");
const PptxDocument_js_1 = require("../core/PptxDocument.js");
/**
 * Copies a Buffer into a standalone ArrayBuffer safe to transfer. Buffers
 * may be views into Node's shared allocation pool, whose backing
 * ArrayBuffer must never be detached, so an explicit copy is required.
 */
function toTransferableArrayBuffer(buffer) {
    const arrayBuffer = new ArrayBuffer(buffer.byteLength);
    new Uint8Array(arrayBuffer).set(buffer);
    return arrayBuffer;
}
/**
 * Revives options after structured cloning: Buffers become plain
 * Uint8Arrays across the thread boundary, so font registration sources are
 * wrapped back into Buffers.
 */
function reviveOptions(options) {
    const register = options.fonts?.register;
    if (!register || register.length === 0) {
        return options;
    }
    const revived = register.map((registration) => {
        const source = registration.source;
        if (typeof source === 'string' || Buffer.isBuffer(source)) {
            return registration;
        }
        const view = source;
        return {
            ...registration,
            source: Buffer.from(view.buffer, view.byteOffset, view.byteLength),
        };
    });
    return { ...options, fonts: { ...options.fonts, register: revived } };
}
/**
 * Processes one render job and posts the response. Never throws: failures
 * are posted back as an error response for the job.
 */
async function handleJob(port, job) {
    try {
        const input = typeof job.input === 'string'
            ? job.input
            : Buffer.from(job.input.buffer, job.input.byteOffset, job.input.byteLength);
        const options = reviveOptions(job.options);
        const document = await PptxDocument_js_1.PptxDocument.open(input, { logLevel: options.logLevel ?? 'warn' });
        try {
            const result = await document.renderAll({ ...options, slideNumbers: job.slideNumbers });
            const transfers = [];
            const slides = result.slides.map((slide) => {
                const imageData = toTransferableArrayBuffer(slide.imageData);
                transfers.push(imageData);
                return {
                    slideIndex: slide.slideIndex,
                    slideNumber: slide.slideNumber,
                    imageData,
                    width: slide.width,
                    height: slide.height,
                    success: slide.success,
                    ...(slide.errorMessage !== undefined ? { errorMessage: slide.errorMessage } : {}),
                    ...(slide.errorStack !== undefined ? { errorStack: slide.errorStack } : {}),
                    ...(slide.warnings !== undefined ? { warnings: slide.warnings } : {}),
                };
            });
            const response = {
                id: job.id,
                ok: true,
                slides,
                totalSlides: result.totalSlides,
                presentationWarnings: (result.warnings ?? []).filter((warning) => warning.slideNumber === undefined),
                errors: result.errors ?? [],
            };
            port.postMessage(response, transfers);
        }
        finally {
            document.close();
        }
    }
    catch (error) {
        const response = {
            id: job.id,
            ok: false,
            message: error instanceof Error ? error.message : String(error),
            ...(error instanceof Error && error.stack !== undefined ? { stack: error.stack } : {}),
        };
        port.postMessage(response);
    }
}
/**
 * Starts the worker loop on the given port. Jobs are processed strictly in
 * arrival order so concurrent renders on one worker never interleave.
 */
function startWorker(port) {
    let queue = Promise.resolve();
    port.on('message', (job) => {
        queue = queue.then(() => handleJob(port, job));
    });
}
if (node_worker_threads_1.parentPort) {
    startWorker(node_worker_threads_1.parentPort);
}
