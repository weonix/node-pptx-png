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
exports.McpStdioServer = exports.McpInvalidParamsError = exports.JSONRPC_ERRORS = exports.SUPPORTED_PROTOCOL_VERSIONS = exports.LATEST_PROTOCOL_VERSION = void 0;
exports.parseSlideList = parseSlideList;
/**
 * Dependency-free MCP (Model Context Protocol) stdio server.
 *
 * Implements the MCP stdio transport: JSON-RPC 2.0 messages delimited by
 * newlines on stdin/stdout (per the MCP spec, stdio messages are
 * newline-delimited JSON and MUST NOT contain embedded newlines — this is
 * NOT LSP-style Content-Length framing). Diagnostics go to stderr only;
 * stdout carries nothing but protocol messages.
 *
 * Exposed tools:
 * - `render_pptx`: render a deck's slides to image files (png/jpeg/webp/svg)
 * - `pptx_info`:   slide count + native dimensions (EMU and px at 96 DPI)
 * - `export_pdf`:  whole-deck (or slide-subset) vector PDF
 *
 * Start via `node dist/mcp/cli.js` (see {@link module:mcp/cli}); the server
 * exits when its stdin closes, per the MCP shutdown model.
 */
const fs = __importStar(require("node:fs"));
const path = __importStar(require("node:path"));
const PptxDocument_js_1 = require("../core/PptxDocument.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
/** Latest MCP protocol revision this server implements. */
exports.LATEST_PROTOCOL_VERSION = '2025-06-18';
/** Protocol revisions the server accepts (echoed back when requested). */
exports.SUPPORTED_PROTOCOL_VERSIONS = [
    // 2025-03-26 is deliberately absent: that revision REQUIRES JSON-RPC
    // batch support, which this server does not implement (batching was
    // removed again in 2025-06-18).
    '2025-06-18',
    '2024-11-05',
];
/** JSON-RPC 2.0 error codes used by this server. */
exports.JSONRPC_ERRORS = {
    PARSE_ERROR: -32700,
    INVALID_REQUEST: -32600,
    METHOD_NOT_FOUND: -32601,
    INVALID_PARAMS: -32602,
    INTERNAL_ERROR: -32603,
};
/**
 * Thrown by argument validation; mapped to a JSON-RPC -32602 (invalid
 * params) error so misuse is distinguishable from tool execution failures
 * (which are reported in-band with `isError: true` so the calling model
 * can read and correct them).
 */
class McpInvalidParamsError extends Error {
}
exports.McpInvalidParamsError = McpInvalidParamsError;
const VALID_FORMATS = ['png', 'jpeg', 'webp', 'svg'];
const VALID_PRESETS = ['thumb', 'preview', 'hd', '4k'];
/** File extension used for each output format. */
const FORMAT_EXTENSIONS = {
    png: 'png',
    jpeg: 'jpg',
    webp: 'webp',
    svg: 'svg',
};
/**
 * Tool definitions advertised by tools/list. `inputSchema` follows the MCP
 * tools spec (JSON Schema object with properties/required).
 */
const TOOL_DEFINITIONS = [
    {
        name: 'render_pptx',
        title: 'Render PPTX slides to images',
        description: 'Renders slides of a PowerPoint (.pptx) file to per-slide image files ' +
            '(slide-<n>.<ext>) in outputDir. Returns a JSON summary with the written ' +
            'file paths, per-slide dimensions, any per-slide failures, and fidelity ' +
            'warnings (e.g. missing fonts or unsupported elements).',
        inputSchema: {
            type: 'object',
            properties: {
                path: {
                    type: 'string',
                    description: 'Path to the .pptx file (absolute or relative to the server cwd).',
                },
                outputDir: {
                    type: 'string',
                    description: 'Directory to write slide images into (created if missing).',
                },
                format: {
                    type: 'string',
                    enum: ['png', 'jpeg', 'webp', 'svg'],
                    description: 'Output image format. Default: png.',
                },
                width: {
                    type: 'number',
                    description: 'Output width in pixels; height keeps the slide aspect ratio. ' +
                        'Mutually exclusive with preset.',
                },
                preset: {
                    type: 'string',
                    enum: ['thumb', 'preview', 'hd', '4k'],
                    description: 'Named size preset: thumb (256px), preview (640px), hd (1920px), 4k (3840px). ' +
                        'Mutually exclusive with width.',
                },
                slides: {
                    type: 'string',
                    description: 'Slides to render as a 1-based list, e.g. "1-5,8". Default: all slides.',
                },
            },
            required: ['path', 'outputDir'],
        },
    },
    {
        name: 'pptx_info',
        title: 'Inspect a PPTX file',
        description: 'Returns metadata for a PowerPoint (.pptx) file without rendering it: ' +
            'slide count and native slide dimensions in EMU and in pixels at 96 DPI.',
        inputSchema: {
            type: 'object',
            properties: {
                path: {
                    type: 'string',
                    description: 'Path to the .pptx file (absolute or relative to the server cwd).',
                },
            },
            required: ['path'],
        },
    },
    {
        name: 'export_pdf',
        title: 'Export PPTX to vector PDF',
        description: 'Exports a PowerPoint (.pptx) file to a vector PDF (one page per slide, ' +
            'text and shapes stay vector). Returns a JSON summary with the output path.',
        inputSchema: {
            type: 'object',
            properties: {
                path: {
                    type: 'string',
                    description: 'Path to the .pptx file (absolute or relative to the server cwd).',
                },
                outputPath: {
                    type: 'string',
                    description: 'Path of the PDF file to write (parent directory is created if missing).',
                },
                slides: {
                    type: 'string',
                    description: 'Slides to include as a 1-based list, e.g. "1-5,8". Default: all slides.',
                },
            },
            required: ['path', 'outputPath'],
        },
    },
];
/**
 * Parses a slides list like "1-5,8" into 1-based slide numbers.
 *
 * Minimal duplicate of the CLI's private parseSlideList (src/cli.ts does
 * not export it); throws {@link McpInvalidParamsError} instead of the
 * CLI's usage error.
 */
function parseSlideList(value) {
    const slides = [];
    for (const part of value.split(',')) {
        const trimmed = part.trim();
        if (trimmed === '') {
            throw new McpInvalidParamsError(`Invalid slides value "${value}": empty entry`);
        }
        const rangeMatch = /^(\d+)-(\d+)$/.exec(trimmed);
        if (rangeMatch) {
            const from = parseInt(rangeMatch[1] ?? '', 10);
            const to = parseInt(rangeMatch[2] ?? '', 10);
            if (from < 1 || to < from) {
                throw new McpInvalidParamsError(`Invalid slides range "${trimmed}": expected 1-based from<=to`);
            }
            if (to - from + 1 > 10000) {
                throw new McpInvalidParamsError(`Invalid slides range "${trimmed}": range spans more than 10000 slides`);
            }
            for (let slideNumber = from; slideNumber <= to; slideNumber++) {
                slides.push(slideNumber);
            }
        }
        else if (/^\d+$/.test(trimmed)) {
            const slideNumber = parseInt(trimmed, 10);
            if (slideNumber < 1) {
                throw new McpInvalidParamsError(`Invalid slides entry "${trimmed}": slide numbers are 1-based`);
            }
            slides.push(slideNumber);
        }
        else {
            throw new McpInvalidParamsError(`Invalid slides entry "${trimmed}": expected N or N-M (e.g. "1-5,8")`);
        }
    }
    return slides;
}
/** Requires `args[key]` to be a non-empty string. */
function requireString(args, key, tool) {
    const value = args[key];
    if (typeof value !== 'string' || value.length === 0) {
        throw new McpInvalidParamsError(`${tool}: "${key}" is required and must be a non-empty string`);
    }
    return value;
}
/** Returns `args[key]` as an optional string, rejecting other types. */
function optionalString(args, key, tool) {
    const value = args[key];
    if (value === undefined || value === null) {
        return undefined;
    }
    if (typeof value !== 'string') {
        throw new McpInvalidParamsError(`${tool}: "${key}" must be a string`);
    }
    return value;
}
/**
 * Resolves a user-supplied path against the server's cwd and verifies the
 * file exists (execution error, not a validation error, so the model sees
 * a readable message and can correct the path).
 */
function resolveInputPath(value) {
    const resolved = path.resolve(value);
    if (!fs.existsSync(resolved)) {
        throw new Error(`Input file not found: ${resolved}`);
    }
    return resolved;
}
/** Validates render_pptx arguments (throws McpInvalidParamsError). */
function parseRenderPptxArgs(args) {
    const tool = 'render_pptx';
    const inputPath = requireString(args, 'path', tool);
    const outputDir = requireString(args, 'outputDir', tool);
    let format = 'png';
    const rawFormat = optionalString(args, 'format', tool);
    if (rawFormat !== undefined) {
        const normalized = rawFormat.toLowerCase() === 'jpg' ? 'jpeg' : rawFormat.toLowerCase();
        if (!VALID_FORMATS.includes(normalized)) {
            throw new McpInvalidParamsError(`${tool}: invalid format "${rawFormat}". Expected one of: ${VALID_FORMATS.join(', ')}`);
        }
        format = normalized;
    }
    let width;
    if (args['width'] !== undefined && args['width'] !== null) {
        const value = args['width'];
        if (typeof value !== 'number' || !Number.isInteger(value) || value < 1) {
            throw new McpInvalidParamsError(`${tool}: "width" must be a positive integer`);
        }
        width = value;
    }
    let preset;
    const rawPreset = optionalString(args, 'preset', tool);
    if (rawPreset !== undefined) {
        if (!VALID_PRESETS.includes(rawPreset)) {
            throw new McpInvalidParamsError(`${tool}: invalid preset "${rawPreset}". Expected one of: ${VALID_PRESETS.join(', ')}`);
        }
        preset = rawPreset;
    }
    if (width !== undefined && preset !== undefined) {
        throw new McpInvalidParamsError(`${tool}: "width" and "preset" are mutually exclusive`);
    }
    const rawSlides = optionalString(args, 'slides', tool);
    const slides = rawSlides !== undefined ? parseSlideList(rawSlides) : undefined;
    return {
        path: inputPath,
        outputDir,
        format,
        ...(width !== undefined ? { width } : {}),
        ...(preset !== undefined ? { preset } : {}),
        ...(slides !== undefined ? { slides } : {}),
    };
}
/** render_pptx: renders slides to files, returns a JSON summary string. */
async function runRenderPptx(rawArgs) {
    const args = parseRenderPptxArgs(rawArgs);
    const inputPath = resolveInputPath(args.path);
    const outputDir = path.resolve(args.outputDir);
    const extension = FORMAT_EXTENSIONS[args.format];
    const document = await (0, PptxDocument_js_1.openPresentation)(inputPath, { logLevel: 'error' });
    const rendered = [];
    const failed = [];
    const warnings = [];
    try {
        fs.mkdirSync(outputDir, { recursive: true });
        const iterator = document.slides({
            ...(args.slides !== undefined ? { slides: args.slides } : {}),
            ...(args.width !== undefined ? { width: args.width } : {}),
            ...(args.preset !== undefined ? { preset: args.preset } : {}),
            format: args.format,
        });
        for await (const slide of iterator) {
            if (slide.warnings) {
                warnings.push(...slide.warnings);
            }
            if (!slide.success) {
                failed.push({
                    slideNumber: slide.slideNumber,
                    error: slide.errorMessage ?? 'unknown error',
                });
                continue;
            }
            const file = path.join(outputDir, `slide-${slide.slideNumber}.${extension}`);
            fs.writeFileSync(file, slide.imageData);
            rendered.push({
                slideNumber: slide.slideNumber,
                file,
                width: slide.width,
                height: slide.height,
            });
        }
        return JSON.stringify({
            input: inputPath,
            outputDir,
            format: args.format,
            slideCount: document.slideCount,
            rendered,
            ...(failed.length > 0 ? { failed } : {}),
            warnings,
        }, null, 2);
    }
    finally {
        document.close();
    }
}
/** pptx_info: returns deck metadata as a JSON string (no rendering). */
async function runPptxInfo(rawArgs) {
    const inputPath = resolveInputPath(requireString(rawArgs, 'path', 'pptx_info'));
    const document = await (0, PptxDocument_js_1.openPresentation)(inputPath, { logLevel: 'error' });
    try {
        return JSON.stringify({
            path: inputPath,
            slideCount: document.slideCount,
            size: {
                widthEmu: document.size.widthEmu,
                heightEmu: document.size.heightEmu,
                widthPx: Math.round(document.size.widthEmu / UnitConverter_js_1.EMU_PER_PIXEL_96DPI),
                heightPx: Math.round(document.size.heightEmu / UnitConverter_js_1.EMU_PER_PIXEL_96DPI),
            },
        }, null, 2);
    }
    finally {
        document.close();
    }
}
/** export_pdf: writes a vector PDF, returns a JSON summary string. */
async function runExportPdf(rawArgs) {
    const tool = 'export_pdf';
    const inputPath = resolveInputPath(requireString(rawArgs, 'path', tool));
    const outputPath = path.resolve(requireString(rawArgs, 'outputPath', tool));
    const rawSlides = optionalString(rawArgs, 'slides', tool);
    const slides = rawSlides !== undefined ? parseSlideList(rawSlides) : undefined;
    const document = await (0, PptxDocument_js_1.openPresentation)(inputPath, { logLevel: 'error' });
    try {
        const pdfBuffer = await document.exportPdf({
            ...(slides !== undefined ? { slides } : {}),
        });
        fs.mkdirSync(path.dirname(outputPath), { recursive: true });
        fs.writeFileSync(outputPath, pdfBuffer);
        return JSON.stringify({
            input: inputPath,
            outputPath,
            bytes: pdfBuffer.length,
            pages: slides !== undefined ? slides.length : document.slideCount,
        }, null, 2);
    }
    finally {
        document.close();
    }
}
/**
 * Reads this package's version by walking up from the executed script to
 * the nearest package.json (same approach as src/cli.ts — import.meta is
 * unavailable because the CJS transpile cannot parse it). Returns
 * 'unknown' rather than failing.
 */
function readVersion() {
    try {
        // CJS build: __dirname points at dist/cjs/mcp, walking up finds this
        // package. ESM: fall back to the spawned entry script (correct for the
        // node-pptx-png-mcp bin, the primary mode).
        let start;
        if (typeof __dirname !== 'undefined') {
            start = path.join(__dirname, 'server.js');
        }
        else {
            const entry = process.argv[1];
            if (!entry) {
                return 'unknown';
            }
            start = fs.realpathSync(entry);
        }
        let dir = path.dirname(start);
        for (;;) {
            const candidate = path.join(dir, 'package.json');
            if (fs.existsSync(candidate)) {
                const pkg = JSON.parse(fs.readFileSync(candidate, 'utf8'));
                // Only trust this package's own manifest (an in-process embedder's
                // walk would otherwise report the HOST application's version)
                return pkg.name === 'node-pptx-png' ? (pkg.version ?? 'unknown') : 'unknown';
            }
            const parent = path.dirname(dir);
            if (parent === dir) {
                return 'unknown';
            }
            dir = parent;
        }
    }
    catch {
        return 'unknown';
    }
}
/**
 * MCP server over newline-delimited JSON-RPC 2.0 (the MCP stdio
 * transport). Construct and call {@link start}; or, for in-process tests,
 * call {@link handleMessage} directly with decoded JSON-RPC messages.
 */
class McpStdioServer {
    input;
    output;
    onClose;
    serverVersion = readVersion();
    buffered = '';
    /** Serializes message handling so responses keep request order. */
    queue = Promise.resolve();
    constructor(options = {}) {
        this.input = options.input ?? process.stdin;
        this.output = options.output ?? process.stdout;
        this.onClose =
            options.onClose ??
                (() => {
                    process.exit(0);
                });
    }
    /** Starts reading newline-delimited JSON-RPC messages from the input. */
    start() {
        this.input.setEncoding('utf8');
        this.input.on('data', (chunk) => {
            this.buffered += typeof chunk === 'string' ? chunk : chunk.toString('utf8');
            this.drainBuffer();
        });
        this.input.on('end', () => {
            // Flush any final line that arrived without a trailing newline,
            // then close once every queued request has been answered.
            if (this.buffered.trim() !== '') {
                this.enqueueLine(this.buffered);
                this.buffered = '';
            }
            void this.queue.then(() => {
                this.onClose();
            });
        });
    }
    /** Splits buffered input into complete lines and queues each. */
    drainBuffer() {
        for (;;) {
            const newlineIndex = this.buffered.indexOf('\n');
            if (newlineIndex === -1) {
                return;
            }
            const line = this.buffered.slice(0, newlineIndex);
            this.buffered = this.buffered.slice(newlineIndex + 1);
            this.enqueueLine(line);
        }
    }
    /** Queues one raw line for sequential handling. */
    enqueueLine(rawLine) {
        const line = rawLine.trim();
        if (line === '') {
            return;
        }
        this.queue = this.queue.then(async () => {
            const response = await this.handleLine(line);
            if (response !== undefined) {
                this.writeMessage(response);
            }
        });
    }
    /** Parses one line; returns the response message, if any. */
    async handleLine(line) {
        let message;
        try {
            message = JSON.parse(line);
        }
        catch {
            return {
                jsonrpc: '2.0',
                id: null,
                error: { code: exports.JSONRPC_ERRORS.PARSE_ERROR, message: 'Parse error: invalid JSON' },
            };
        }
        return this.handleMessage(message);
    }
    /**
     * Handles one decoded JSON-RPC message. Returns the response for
     * requests, or undefined for notifications (and for client responses,
     * which this server never solicits). Exposed for in-process tests.
     */
    async handleMessage(message) {
        if (typeof message !== 'object' || message === null || Array.isArray(message)) {
            return {
                jsonrpc: '2.0',
                id: null,
                error: {
                    code: exports.JSONRPC_ERRORS.INVALID_REQUEST,
                    message: 'Invalid request: expected a JSON-RPC 2.0 message object',
                },
            };
        }
        const record = message;
        const rawId = record['id'];
        const id = typeof rawId === 'string' || typeof rawId === 'number' ? rawId : null;
        const method = record['method'];
        if (typeof method !== 'string') {
            // A message without a method is either a client response (ignored —
            // this server sends no requests) or malformed.
            if ('result' in record || 'error' in record) {
                return undefined;
            }
            return {
                jsonrpc: '2.0',
                id,
                error: { code: exports.JSONRPC_ERRORS.INVALID_REQUEST, message: 'Invalid request: missing method' },
            };
        }
        // Notifications (no id) never get a response, known or not.
        const isNotification = rawId === undefined || rawId === null;
        if (isNotification) {
            // notifications/initialized, notifications/cancelled, etc.
            return undefined;
        }
        const params = typeof record['params'] === 'object' && record['params'] !== null
            ? record['params']
            : {};
        try {
            switch (method) {
                case 'initialize':
                    return { jsonrpc: '2.0', id, result: this.handleInitialize(params) };
                case 'ping':
                    return { jsonrpc: '2.0', id, result: {} };
                case 'tools/list':
                    return { jsonrpc: '2.0', id, result: { tools: TOOL_DEFINITIONS } };
                case 'tools/call':
                    return { jsonrpc: '2.0', id, result: await this.handleToolCall(params) };
                default:
                    return {
                        jsonrpc: '2.0',
                        id,
                        error: {
                            code: exports.JSONRPC_ERRORS.METHOD_NOT_FOUND,
                            message: `Method not found: ${method}`,
                        },
                    };
            }
        }
        catch (error) {
            if (error instanceof McpInvalidParamsError) {
                return {
                    jsonrpc: '2.0',
                    id,
                    error: { code: exports.JSONRPC_ERRORS.INVALID_PARAMS, message: error.message },
                };
            }
            const messageText = error instanceof Error ? error.message : String(error);
            return {
                jsonrpc: '2.0',
                id,
                error: { code: exports.JSONRPC_ERRORS.INTERNAL_ERROR, message: messageText },
            };
        }
    }
    /** Builds the initialize result (version negotiation + capabilities). */
    handleInitialize(params) {
        const requested = params['protocolVersion'];
        const protocolVersion = typeof requested === 'string' && exports.SUPPORTED_PROTOCOL_VERSIONS.includes(requested)
            ? requested
            : exports.LATEST_PROTOCOL_VERSION;
        return {
            protocolVersion,
            capabilities: { tools: {} },
            serverInfo: {
                name: 'node-pptx-png',
                title: 'node-pptx-png PPTX renderer',
                version: this.serverVersion,
            },
        };
    }
    /**
     * Dispatches tools/call. Unknown tool names and invalid arguments are
     * JSON-RPC errors (thrown); tool execution failures are reported
     * in-band with isError so the calling model can read them.
     */
    async handleToolCall(params) {
        const name = params['name'];
        if (typeof name !== 'string') {
            throw new McpInvalidParamsError('tools/call: "name" is required and must be a string');
        }
        const rawArguments = params['arguments'];
        if (rawArguments !== undefined &&
            (typeof rawArguments !== 'object' || rawArguments === null || Array.isArray(rawArguments))) {
            throw new McpInvalidParamsError('tools/call: "arguments" must be an object');
        }
        const args = (rawArguments ?? {});
        let runTool;
        switch (name) {
            case 'render_pptx':
                runTool = runRenderPptx;
                break;
            case 'pptx_info':
                runTool = runPptxInfo;
                break;
            case 'export_pdf':
                runTool = runExportPdf;
                break;
            default:
                throw new McpInvalidParamsError(`Unknown tool: ${name}. Available tools: ${TOOL_DEFINITIONS.map((t) => t.name).join(', ')}`);
        }
        try {
            const text = await runTool(args);
            return { content: [{ type: 'text', text }] };
        }
        catch (error) {
            if (error instanceof McpInvalidParamsError) {
                throw error;
            }
            const messageText = error instanceof Error ? error.message : String(error);
            return { content: [{ type: 'text', text: `Error: ${messageText}` }], isError: true };
        }
    }
    /** Writes one message as a single line (MCP stdio framing). */
    writeMessage(message) {
        // JSON.stringify escapes any newline inside string values, so the
        // serialized message never contains an embedded newline.
        this.output.write(`${JSON.stringify(message)}\n`);
    }
}
exports.McpStdioServer = McpStdioServer;
