/**
 * FidelityTester - Main testing orchestration for visual fidelity testing
 *
 * Loads baseline images, renders PPTX slides, compares them,
 * and generates detailed reports with per-slide and overall fidelity scores.
 */
import type { PptxRenderOptions } from '../types/index.js';
/**
 * Report for a single slide comparison.
 */
export interface SlideReport {
    /**
     * One-based slide number.
     */
    slideNumber: number;
    /**
     * Structural Similarity Index (0-1).
     */
    ssim: number;
    /**
     * Mean Squared Error.
     */
    mse: number;
    /**
     * Peak Signal-to-Noise Ratio (dB).
     */
    psnr: number;
    /**
     * Percentage of differing pixels.
     */
    pixelDiffPercent: number;
    /**
     * Path to baseline image.
     */
    baselinePath: string;
    /**
     * Path to rendered image (if saved).
     */
    renderedPath: string;
    /**
     * Path to diff image (if generated).
     */
    diffPath?: string;
    /**
     * Whether this slide passed the fidelity threshold.
     */
    passed: boolean;
    /**
     * Error message if comparison failed.
     */
    error?: string;
}
/**
 * Complete fidelity test report.
 */
export interface FidelityReport {
    /**
     * Path to the tested PPTX file.
     */
    pptxPath: string;
    /**
     * When the test was run.
     */
    timestamp: Date;
    /**
     * Average SSIM across all slides.
     */
    overallFidelity: number;
    /**
     * Per-slide comparison reports.
     */
    slides: SlideReport[];
    /**
     * Whether all slides met the target fidelity.
     */
    passed: boolean;
    /**
     * Target fidelity threshold used.
     */
    targetFidelity: number;
    /**
     * Total number of slides.
     */
    totalSlides: number;
    /**
     * Number of slides that passed.
     */
    passedSlides: number;
    /**
     * Test duration in milliseconds.
     */
    durationMs: number;
}
/**
 * Options for fidelity testing.
 */
export interface FidelityTestOptions {
    /**
     * Path to directory containing baseline images.
     * Images should be named slide-1.png, slide-2.png, etc.
     */
    baselinesDir: string;
    /**
     * Target SSIM score to pass (0-1).
     * Default: 0.95
     */
    targetFidelity?: number;
    /**
     * Output directory for rendered images and diffs.
     * Default: undefined (no output saved)
     */
    outputDir?: string;
    /**
     * Generate diff images for debugging.
     * Default: true
     */
    generateDiffs?: boolean;
    /**
     * Render options for PPTX rendering.
     */
    renderOptions?: PptxRenderOptions;
    /**
     * Pixel difference threshold (0-255).
     * Default: 0
     */
    pixelThreshold?: number;
}
/**
 * Runs fidelity tests for a PPTX presentation.
 *
 * @param pptxPath - Path to the PPTX file
 * @param options - Fidelity test options
 * @returns Fidelity report with per-slide and overall scores
 */
export declare function runFidelityTest(pptxPath: string, options: FidelityTestOptions): Promise<FidelityReport>;
/**
 * Formats a fidelity report for console output.
 */
export declare function formatReport(report: FidelityReport): string;
/**
 * FidelityTester class for object-oriented usage.
 */
export declare class FidelityTester {
    private readonly renderer;
    private readonly defaultOptions;
    constructor(options?: Partial<FidelityTestOptions>);
    /**
     * Runs fidelity tests for a PPTX presentation.
     */
    test(pptxPath: string, options: FidelityTestOptions): Promise<FidelityReport>;
    /**
     * Quick check if a PPTX meets a minimum fidelity threshold.
     */
    passes(pptxPath: string, baselinesDir: string, targetFidelity?: number): Promise<boolean>;
}
