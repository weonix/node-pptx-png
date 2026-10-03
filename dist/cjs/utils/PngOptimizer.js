"use strict";
/**
 * PNG optimization utility using Sharp for compression.
 * Provides presets and custom options for optimizing PNG output.
 */
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
exports.PngOptimizer = exports.PNG_PRESETS = void 0;
exports.createPngOptimizer = createPngOptimizer;
const Logger_js_1 = require("./Logger.js");
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
exports.PNG_PRESETS = {
    /** No optimization - use native canvas export (fastest) */
    none: {},
    /** Fast lossless compression with adaptive filtering */
    fast: {
        compressionLevel: 6,
        adaptiveFiltering: true,
    },
    /** Balanced lossless compression - slightly better than fast */
    balanced: {
        compressionLevel: 9,
        adaptiveFiltering: true,
    },
    /**
     * Maximum lossless compression.
     * Note: Identical to 'balanced' because skia-canvas already uses efficient
     * compression. This preset exists for API completeness - users expect a
     * "maximum" option. For significantly smaller files, use 'web' preset instead.
     */
    maximum: {
        compressionLevel: 9,
        adaptiveFiltering: true,
    },
    /**
     * Web-optimized with palette quantization.
     * Provides 60-70% size reduction but may affect quality on:
     * - Photos and gradients (> 256 colors)
     * - Smooth color transitions
     * Best for diagrams, text-heavy slides, and simple graphics.
     */
    web: {
        compressionLevel: 9,
        adaptiveFiltering: true,
        palette: true,
        colors: 256,
        quality: 85,
        dither: 1.0,
    },
};
/**
 * PNG optimizer using Sharp for compression.
 * Gracefully falls back to native canvas export if Sharp is not available.
 */
class PngOptimizer {
    sharp = null;
    initialized = false;
    logger;
    constructor(logger) {
        this.logger = logger ?? (0, Logger_js_1.createLogger)('warn', 'PngOptimizer');
    }
    /**
     * Initializes the optimizer by attempting to load Sharp.
     * @returns true if Sharp is available, false otherwise
     */
    async initialize() {
        if (this.initialized) {
            return this.sharp !== null;
        }
        try {
            // Dynamic import of sharp
            const sharpModule = await Promise.resolve().then(() => __importStar(require('sharp')));
            this.sharp = sharpModule.default;
            this.initialized = true;
            this.logger.debug('Sharp loaded successfully');
            return true;
        }
        catch (error) {
            this.logger.debug('Sharp not available, PNG optimization disabled', {
                error: error instanceof Error ? error.message : String(error),
            });
            this.initialized = true;
            return false;
        }
    }
    /**
     * Checks if Sharp is available for optimization.
     */
    isAvailable() {
        return this.sharp !== null;
    }
    /**
     * Optimizes a PNG buffer using Sharp.
     *
     * @param pngBuffer Raw PNG buffer from canvas
     * @param options Optimization options or preset name
     * @returns Optimized PNG buffer
     */
    async optimize(pngBuffer, options = 'balanced') {
        // Resolve preset to options
        const opts = typeof options === 'string' ? exports.PNG_PRESETS[options] : options;
        // If no optimization requested or Sharp not available, return original
        if (!this.sharp || (typeof options === 'string' && options === 'none')) {
            return pngBuffer;
        }
        try {
            return await this.applyOptimization(pngBuffer, opts);
        }
        catch (error) {
            this.logger.warn('PNG optimization failed, returning original', {
                error: error instanceof Error ? error.message : String(error),
            });
            return pngBuffer;
        }
    }
    /**
     * Optimizes a canvas directly.
     * Gets the PNG buffer from canvas and optimizes it.
     *
     * @param canvas Skia canvas to optimize
     * @param options Optimization options or preset name
     * @returns Optimized PNG buffer
     */
    async optimizeCanvas(canvas, options = 'balanced') {
        // Get raw PNG from canvas
        const pngBuffer = await canvas.toBuffer('png');
        return this.optimize(pngBuffer, options);
    }
    /**
     * Applies Sharp optimization to a PNG buffer.
     */
    async applyOptimization(pngBuffer, options) {
        if (!this.sharp) {
            return pngBuffer;
        }
        const sharpInstance = this.sharp(pngBuffer);
        // If palette mode requested, try it first with fallback
        if (options.palette) {
            try {
                return await sharpInstance
                    .png({
                    palette: true,
                    colors: options.colors ?? 256,
                    quality: options.quality ?? 90,
                    dither: options.dither ?? 1.0,
                    compressionLevel: options.compressionLevel ?? 9,
                })
                    .toBuffer();
            }
            catch (paletteError) {
                // Palette mode failed (likely too many colors), fall back to non-palette
                this.logger.debug('Palette mode failed, falling back to standard compression', {
                    error: paletteError instanceof Error ? paletteError.message : String(paletteError),
                });
                // Create new instance for fallback (Sharp instances are single-use)
                return await this.sharp(pngBuffer)
                    .png({
                    compressionLevel: options.compressionLevel ?? 9,
                    adaptiveFiltering: options.adaptiveFiltering ?? true,
                })
                    .toBuffer();
            }
        }
        // Standard compression without palette
        return await sharpInstance
            .png({
            compressionLevel: options.compressionLevel ?? 6,
            adaptiveFiltering: options.adaptiveFiltering ?? true,
        })
            .toBuffer();
    }
    /**
     * Gets compression statistics for a buffer.
     * Useful for benchmarking and debugging.
     *
     * @param original Original PNG buffer
     * @param optimized Optimized PNG buffer
     * @returns Compression statistics
     */
    getCompressionStats(original, optimized) {
        const originalSize = original.length;
        const optimizedSize = optimized.length;
        const savedBytes = originalSize - optimizedSize;
        const reductionPercent = originalSize > 0 ? (savedBytes / originalSize) * 100 : 0;
        return {
            originalSize,
            optimizedSize,
            savedBytes,
            reductionPercent: Math.round(reductionPercent * 100) / 100,
        };
    }
}
exports.PngOptimizer = PngOptimizer;
/**
 * Creates a PNG optimizer instance.
 * Call initialize() before using optimize methods.
 */
function createPngOptimizer(logger) {
    return new PngOptimizer(logger);
}
