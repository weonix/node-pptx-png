/**
 * PNG optimization utility using Sharp for compression.
 * Provides presets and custom options for optimizing PNG output.
 */
import type { Canvas } from 'skia-canvas';
import type { PngOptimizationPreset, PngOptimizationOptions } from '../types/options.js';
import type { ILogger } from './Logger.js';
/**
 * Preset configurations for PNG optimization.
 *
 * Note: The native skia-canvas encoder is already efficient (similar to zlib level 6).
 * Sharp's main value is palette-based quantization for dramatic size reduction.
 *
 * Lossless presets (fast/balanced/maximum) provide ~2-3% improvement.
 * Palette-based presets (web) provide 60-70% reduction but may affect quality.
 *
 * - 'none': Skip Sharp, use native encoder (fastest)
 * - 'fast': Quick lossless recompression (~2% smaller)
 * - 'balanced': Lossless with adaptive filtering (~2-3% smaller)
 * - 'maximum': Same as balanced (skia-canvas is already efficient)
 * - 'web': Palette quantization (60-70% smaller, may lose quality on photos)
 */
export declare const PNG_PRESETS: Record<PngOptimizationPreset, PngOptimizationOptions>;
/**
 * PNG optimizer using Sharp for compression.
 * Gracefully falls back to native canvas export if Sharp is not available.
 */
export declare class PngOptimizer {
    private sharp;
    private initialized;
    private readonly logger;
    constructor(logger?: ILogger);
    /**
     * Initializes the optimizer by attempting to load Sharp.
     * @returns true if Sharp is available, false otherwise
     */
    initialize(): Promise<boolean>;
    /**
     * Checks if Sharp is available for optimization.
     */
    isAvailable(): boolean;
    /**
     * Optimizes a PNG buffer using Sharp.
     *
     * @param pngBuffer Raw PNG buffer from canvas
     * @param options Optimization options or preset name
     * @returns Optimized PNG buffer
     */
    optimize(pngBuffer: Buffer, options?: PngOptimizationPreset | PngOptimizationOptions): Promise<Buffer>;
    /**
     * Optimizes a canvas directly.
     * Gets the PNG buffer from canvas and optimizes it.
     *
     * @param canvas Skia canvas to optimize
     * @param options Optimization options or preset name
     * @returns Optimized PNG buffer
     */
    optimizeCanvas(canvas: Canvas, options?: PngOptimizationPreset | PngOptimizationOptions): Promise<Buffer>;
    /**
     * Applies Sharp optimization to a PNG buffer.
     */
    private applyOptimization;
    /**
     * Gets compression statistics for a buffer.
     * Useful for benchmarking and debugging.
     *
     * @param original Original PNG buffer
     * @param optimized Optimized PNG buffer
     * @returns Compression statistics
     */
    getCompressionStats(original: Buffer, optimized: Buffer): {
        originalSize: number;
        optimizedSize: number;
        savedBytes: number;
        reductionPercent: number;
    };
}
/**
 * Creates a PNG optimizer instance.
 * Call initialize() before using optimize methods.
 */
export declare function createPngOptimizer(logger?: ILogger): PngOptimizer;
//# sourceMappingURL=PngOptimizer.d.ts.map