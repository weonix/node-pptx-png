/**
 * VisualComparator - Image comparison with quality metrics
 *
 * Compares two images pixel-by-pixel and calculates industry-standard
 * image quality metrics including SSIM, MSE, and PSNR.
 */
/**
 * Result of comparing two images.
 */
export interface ComparisonResult {
    /**
     * Structural Similarity Index (0-1, higher is better).
     * 0.99 = 99% similar. Industry standard for perceptual quality.
     */
    ssim: number;
    /**
     * Mean Squared Error (lower is better).
     * 0 = identical images.
     */
    mse: number;
    /**
     * Peak Signal-to-Noise Ratio in decibels (higher is better).
     * Typical values: 30-50 dB for good quality.
     */
    psnr: number;
    /**
     * Percentage of pixels that differ (0-100).
     */
    pixelDiffPercent: number;
    /**
     * Visual diff image highlighting differences in red.
     */
    diffImage?: Buffer;
    /**
     * Width of the compared images in pixels.
     */
    width: number;
    /**
     * Height of the compared images in pixels.
     */
    height: number;
}
/**
 * Options for image comparison.
 */
export interface ComparisonOptions {
    /**
     * Generate a visual diff image highlighting differences.
     * Default: true
     */
    generateDiffImage?: boolean;
    /**
     * Threshold for considering pixels different (0-255).
     * Default: 0 (exact match required)
     */
    pixelThreshold?: number;
    /**
     * SSIM window size for local comparison.
     * Default: 11 (standard 11x11 window)
     */
    ssimWindowSize?: number;
}
/**
 * Compares two images and returns similarity metrics.
 *
 * @param baseline - Path to baseline image or Buffer
 * @param rendered - Path to rendered image or Buffer
 * @param options - Comparison options
 * @returns Comparison result with similarity metrics
 */
export declare function compareImages(baseline: string | Buffer, rendered: string | Buffer, options?: ComparisonOptions): Promise<ComparisonResult>;
/**
 * VisualComparator class for object-oriented usage.
 */
export declare class VisualComparator {
    private readonly options;
    constructor(options?: ComparisonOptions);
    /**
     * Compares two images and returns similarity metrics.
     */
    compare(baseline: string | Buffer, rendered: string | Buffer, options?: ComparisonOptions): Promise<ComparisonResult>;
    /**
     * Checks if two images meet a minimum SSIM threshold.
     */
    meetsThreshold(baseline: string | Buffer, rendered: string | Buffer, minSSIM: number): Promise<boolean>;
}
