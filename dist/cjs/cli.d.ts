#!/usr/bin/env node
import type { ImageFormat, SizePreset } from './types/options.js';
/**
 * Error thrown by {@link parseArgs} for invalid command lines. The CLI
 * prints the message plus usage and exits with code 2.
 */
export declare class CliUsageError extends Error {
}
/**
 * Parsed command-line options.
 */
export interface CliOptions {
    /** Input .pptx path. Empty only when help/version was requested. */
    input: string;
    /** Output directory for slide images. */
    outDir: string;
    /** Explicit image format, or undefined when --format was not given. */
    format: ImageFormat | undefined;
    /** Explicit width in pixels (mutually exclusive with scale/preset). */
    width?: number;
    /** Scale factor (mutually exclusive with width/preset). */
    scale?: number;
    /** Size preset (mutually exclusive with width/scale). */
    preset?: SizePreset;
    /** 1-based slide numbers to render, or undefined for all slides. */
    slides?: number[];
    /** Output path for the whole-deck vector PDF, when --pdf was given. */
    pdf?: string;
    /** Lossy encoding quality (0-1) for jpeg/webp. */
    quality?: number;
    /** Suppress progress output. */
    quiet: boolean;
    /** Emit a machine-readable JSON summary on stdout. */
    json: boolean;
    /** --help was given. */
    help: boolean;
    /** --version was given. */
    version: boolean;
}
/**
 * Returns the CLI usage text.
 */
export declare function usage(): string;
/**
 * Parses CLI arguments (without the node/script prefix) into options.
 * Throws {@link CliUsageError} on unknown flags, missing values, invalid
 * values, or conflicting size options.
 */
export declare function parseArgs(argv: string[]): CliOptions;
/**
 * Runs the CLI with the given arguments (excluding the node/script
 * prefix) and returns the process exit code. Human-readable progress goes
 * to stderr; the --json summary is the only stdout output in json mode.
 */
export declare function run(argv: string[]): Promise<number>;
