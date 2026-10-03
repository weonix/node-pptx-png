#!/usr/bin/env node
/**
 * node-pptx-png command-line interface.
 *
 * Renders a PPTX deck to per-slide image files and/or a whole-deck vector
 * PDF, with a machine-readable --json summary. Argument parsing is
 * hand-rolled (no dependencies); {@link parseArgs} is exported for tests.
 *
 * Usage:
 *   node-pptx-png <input.pptx> [-o outdir] [--format png|jpeg|webp|svg]
 *     [--width N | --scale N | --preset thumb|preview|hd|4k]
 *     [--slides 1-5,8] [--pdf out.pdf] [--quality 0.9] [--quiet] [--json]
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { openPresentation } from './core/PptxDocument.js';
/** File extension used for each output format. */
const FORMAT_EXTENSIONS = {
    png: 'png',
    jpeg: 'jpg',
    webp: 'webp',
    svg: 'svg',
};
const VALID_FORMATS = ['png', 'jpeg', 'webp', 'svg'];
const VALID_PRESETS = ['thumb', 'preview', 'hd', '4k'];
/**
 * Error thrown by {@link parseArgs} for invalid command lines. The CLI
 * prints the message plus usage and exits with code 2.
 */
export class CliUsageError extends Error {
}
/**
 * Returns the CLI usage text.
 */
export function usage() {
    return `Usage: node-pptx-png <input.pptx> [options]

Renders each slide to <outdir>/slide-<n>.<ext> and/or a vector PDF.

Options:
  -o, --outdir <dir>      Output directory (default: current directory)
  --format <fmt>          Image format: png | jpeg | webp | svg (default: png)
  --width <N>             Output width in pixels (height keeps aspect ratio)
  --scale <N>             Scale factor on the slide's native 96-DPI size
  --preset <name>         Size preset: thumb (256) | preview (640) | hd (1920) | 4k (3840)
  --slides <list>         Slides to render, e.g. "1-5,8" (1-based)
  --pdf <file>            Write a whole-deck vector PDF; without --format,
                          the PDF replaces the per-slide images
  --quality <N>           Quality 0-1 for jpeg/webp (default: 0.9)
  --quiet                 Suppress progress output
  --json                  Print a JSON result summary to stdout
  -h, --help              Show this help
  -v, --version           Print the package version

Exit codes: 0 success, 1 any slide or PDF failed, 2 invalid arguments`;
}
/**
 * Parses a --slides list like "1-5,8" into 1-based slide numbers.
 */
function parseSlideList(value) {
    const slides = [];
    for (const part of value.split(',')) {
        const trimmed = part.trim();
        if (trimmed === '') {
            throw new CliUsageError(`Invalid --slides value "${value}": empty entry`);
        }
        const rangeMatch = /^(\d+)-(\d+)$/.exec(trimmed);
        if (rangeMatch) {
            const from = parseInt(rangeMatch[1] ?? '', 10);
            const to = parseInt(rangeMatch[2] ?? '', 10);
            if (from < 1 || to < from) {
                throw new CliUsageError(`Invalid --slides range "${trimmed}": expected 1-based from<=to`);
            }
            // Bound eager expansion: no deck has anywhere near this many slides,
            // and an unbounded range (e.g. a typo like 1-100000000) would OOM
            if (to - from + 1 > 10000) {
                throw new CliUsageError(`Invalid --slides range "${trimmed}": range spans more than 10000 slides`);
            }
            for (let slideNumber = from; slideNumber <= to; slideNumber++) {
                slides.push(slideNumber);
            }
        }
        else if (/^\d+$/.test(trimmed)) {
            const slideNumber = parseInt(trimmed, 10);
            if (slideNumber < 1) {
                throw new CliUsageError(`Invalid --slides entry "${trimmed}": slide numbers are 1-based`);
            }
            slides.push(slideNumber);
        }
        else {
            throw new CliUsageError(`Invalid --slides entry "${trimmed}": expected N or N-M`);
        }
    }
    return slides;
}
/**
 * Parses CLI arguments (without the node/script prefix) into options.
 * Throws {@link CliUsageError} on unknown flags, missing values, invalid
 * values, or conflicting size options.
 */
export function parseArgs(argv) {
    const options = {
        input: '',
        outDir: '.',
        format: undefined,
        quiet: false,
        json: false,
        help: false,
        version: false,
    };
    const positionals = [];
    let index = 0;
    const next = (flag) => {
        const value = argv[++index];
        if (value === undefined) {
            throw new CliUsageError(`Missing value for ${flag}`);
        }
        return value;
    };
    for (; index < argv.length; index++) {
        const raw = argv[index];
        if (raw === undefined) {
            continue;
        }
        // Support --flag=value for long flags
        let arg = raw;
        let inlineValue;
        if (arg.startsWith('--')) {
            const eq = arg.indexOf('=');
            if (eq !== -1) {
                inlineValue = arg.slice(eq + 1);
                arg = arg.slice(0, eq);
            }
        }
        const value = (flag) => inlineValue ?? next(flag);
        switch (arg) {
            case '-o':
            case '--outdir':
                options.outDir = value(arg);
                break;
            case '--format': {
                const fmt = value(arg).toLowerCase();
                const normalized = fmt === 'jpg' ? 'jpeg' : fmt;
                if (!VALID_FORMATS.includes(normalized)) {
                    throw new CliUsageError(`Invalid --format "${fmt}". Expected one of: ${VALID_FORMATS.join(', ')}`);
                }
                options.format = normalized;
                break;
            }
            case '--width': {
                const width = Number(value(arg));
                if (!Number.isInteger(width) || width < 1) {
                    throw new CliUsageError(`Invalid --width "${String(width)}": expected a positive integer`);
                }
                options.width = width;
                break;
            }
            case '--scale': {
                const scale = Number(value(arg));
                if (!Number.isFinite(scale) || scale <= 0) {
                    throw new CliUsageError(`Invalid --scale "${String(scale)}": expected a positive number`);
                }
                options.scale = scale;
                break;
            }
            case '--preset': {
                const preset = value(arg);
                if (!VALID_PRESETS.includes(preset)) {
                    throw new CliUsageError(`Invalid --preset "${preset}". Expected one of: ${VALID_PRESETS.join(', ')}`);
                }
                options.preset = preset;
                break;
            }
            case '--slides':
                options.slides = parseSlideList(value(arg));
                break;
            case '--pdf':
                options.pdf = value(arg);
                break;
            case '--quality': {
                const quality = Number(value(arg));
                if (!Number.isFinite(quality) || quality < 0 || quality > 1) {
                    throw new CliUsageError(`Invalid --quality "${String(quality)}": expected a number between 0 and 1`);
                }
                options.quality = quality;
                break;
            }
            case '--quiet':
                options.quiet = true;
                break;
            case '--json':
                options.json = true;
                break;
            case '-h':
            case '--help':
                options.help = true;
                break;
            case '-v':
            case '--version':
                options.version = true;
                break;
            default:
                if (arg.startsWith('-')) {
                    throw new CliUsageError(`Unknown option: ${arg}`);
                }
                positionals.push(arg);
        }
    }
    if (options.help || options.version) {
        return options;
    }
    if (positionals.length === 0) {
        throw new CliUsageError('Missing input file. Expected: node-pptx-png <input.pptx> [options]');
    }
    if (positionals.length > 1) {
        throw new CliUsageError(`Unexpected extra arguments: ${positionals.slice(1).join(' ')}`);
    }
    options.input = positionals[0] ?? '';
    const sizeFlags = [
        options.width !== undefined ? '--width' : undefined,
        options.scale !== undefined ? '--scale' : undefined,
        options.preset !== undefined ? '--preset' : undefined,
    ].filter((flag) => flag !== undefined);
    if (sizeFlags.length > 1) {
        throw new CliUsageError(`${sizeFlags.join(' and ')} are mutually exclusive`);
    }
    return options;
}
/**
 * Reads this package's version by walking up from the executed script to
 * the nearest package.json. Returns 'unknown' when it cannot be found
 * (the CLI must not fail over version metadata).
 */
function readVersion() {
    try {
        const entry = process.argv[1];
        if (!entry) {
            return 'unknown';
        }
        let dir = path.dirname(fs.realpathSync(entry));
        for (;;) {
            const candidate = path.join(dir, 'package.json');
            if (fs.existsSync(candidate)) {
                const pkg = JSON.parse(fs.readFileSync(candidate, 'utf8'));
                return pkg.version ?? 'unknown';
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
 * Runs the CLI with the given arguments (excluding the node/script
 * prefix) and returns the process exit code. Human-readable progress goes
 * to stderr; the --json summary is the only stdout output in json mode.
 */
export async function run(argv) {
    let options;
    try {
        options = parseArgs(argv);
    }
    catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        process.stderr.write(`${message}\n\n${usage()}\n`);
        return 2;
    }
    if (options.help) {
        process.stdout.write(`${usage()}\n`);
        return 0;
    }
    if (options.version) {
        process.stdout.write(`${readVersion()}\n`);
        return 0;
    }
    const log = (line) => {
        if (!options.quiet) {
            process.stderr.write(`${line}\n`);
        }
    };
    const writeImages = options.pdf === undefined || options.format !== undefined;
    const format = options.format ?? 'png';
    const extension = FORMAT_EXTENSIONS[format];
    const sizeOptions = {
        ...(options.width !== undefined ? { width: options.width } : {}),
        ...(options.scale !== undefined ? { scale: options.scale } : {}),
        ...(options.preset !== undefined ? { preset: options.preset } : {}),
    };
    const slideSummaries = [];
    const warnings = [];
    let pdfFile = null;
    let failures = 0;
    try {
        const document = await openPresentation(options.input, {
            logLevel: options.quiet || options.json ? 'error' : 'warn',
        });
        try {
            if (writeImages) {
                fs.mkdirSync(options.outDir, { recursive: true });
                const iterator = document.slides({
                    ...(options.slides !== undefined ? { slides: options.slides } : {}),
                    ...sizeOptions,
                    format,
                    ...(options.quality !== undefined ? { quality: options.quality } : {}),
                });
                for await (const slide of iterator) {
                    if (slide.warnings) {
                        warnings.push(...slide.warnings);
                    }
                    if (!slide.success) {
                        failures++;
                        log(`slide ${slide.slideNumber}: FAILED (${slide.errorMessage ?? 'unknown error'})`);
                        slideSummaries.push({
                            slideNumber: slide.slideNumber,
                            file: null,
                            width: slide.width,
                            height: slide.height,
                            success: false,
                            ...(slide.errorMessage !== undefined ? { errorMessage: slide.errorMessage } : {}),
                        });
                        continue;
                    }
                    const file = path.join(options.outDir, `slide-${slide.slideNumber}.${extension}`);
                    fs.writeFileSync(file, slide.imageData);
                    log(`slide ${slide.slideNumber} -> ${file} (${slide.width}x${slide.height})`);
                    slideSummaries.push({
                        slideNumber: slide.slideNumber,
                        file,
                        width: slide.width,
                        height: slide.height,
                        success: true,
                    });
                }
            }
            if (options.pdf !== undefined) {
                try {
                    const pdfBuffer = await document.exportPdf({
                        ...(options.slides !== undefined ? { slides: options.slides } : {}),
                        ...sizeOptions,
                    });
                    const pdfDir = path.dirname(options.pdf);
                    if (pdfDir !== '.') {
                        fs.mkdirSync(pdfDir, { recursive: true });
                    }
                    fs.writeFileSync(options.pdf, pdfBuffer);
                    pdfFile = options.pdf;
                    log(`pdf -> ${options.pdf}`);
                }
                catch (error) {
                    failures++;
                    const message = error instanceof Error ? error.message : String(error);
                    log(`pdf: FAILED (${message})`);
                }
            }
        }
        finally {
            document.close();
        }
    }
    catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        process.stderr.write(`Error: ${message}\n`);
        if (options.json) {
            process.stdout.write(`${JSON.stringify({ success: false, error: message })}\n`);
        }
        return 1;
    }
    if (options.json) {
        const summary = {
            success: failures === 0,
            input: options.input,
            ...(writeImages ? { outDir: options.outDir, format } : {}),
            slides: slideSummaries,
            ...(pdfFile !== null ? { pdf: pdfFile } : {}),
            warnings,
        };
        process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
    }
    return failures === 0 ? 0 : 1;
}
/**
 * Detects direct execution without import.meta (which the CJS transpile
 * cannot parse): the CLI runs when the executed script is this file or
 * the installed bin alias.
 */
function isMainModule() {
    const entry = process.argv[1];
    if (!entry) {
        return false;
    }
    const base = path.basename(entry);
    return base === 'cli.js' || base === 'cli.ts' || base === 'node-pptx-png';
}
if (isMainModule()) {
    void run(process.argv.slice(2)).then((code) => {
        process.exitCode = code;
    });
}
//# sourceMappingURL=cli.js.map