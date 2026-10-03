"use strict";
/**
 * FidelityTester - Main testing orchestration for visual fidelity testing
 *
 * Loads baseline images, renders PPTX slides, compares them,
 * and generates detailed reports with per-slide and overall fidelity scores.
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
exports.FidelityTester = void 0;
exports.runFidelityTest = runFidelityTest;
exports.formatReport = formatReport;
const fs = __importStar(require("node:fs"));
const path = __importStar(require("node:path"));
const PptxImageRenderer_js_1 = require("../core/PptxImageRenderer.js");
const VisualComparator_js_1 = require("./VisualComparator.js");
/**
 * Default fidelity test options.
 */
const DEFAULT_OPTIONS = {
    targetFidelity: 0.95,
    outputDir: '',
    generateDiffs: true,
    renderOptions: {},
    pixelThreshold: 0,
};
/**
 * Runs fidelity tests for a PPTX presentation.
 *
 * @param pptxPath - Path to the PPTX file
 * @param options - Fidelity test options
 * @returns Fidelity report with per-slide and overall scores
 */
async function runFidelityTest(pptxPath, options) {
    const startTime = Date.now();
    const opts = { ...DEFAULT_OPTIONS, ...options };
    // Validate inputs
    if (!fs.existsSync(pptxPath)) {
        throw new Error(`PPTX file not found: ${pptxPath}`);
    }
    if (!fs.existsSync(opts.baselinesDir)) {
        throw new Error(`Baselines directory not found: ${opts.baselinesDir}`);
    }
    // Create output directory if specified
    if (opts.outputDir) {
        fs.mkdirSync(opts.outputDir, { recursive: true });
    }
    // Render the presentation
    const renderer = new PptxImageRenderer_js_1.PptxImageRenderer({ logLevel: opts.renderOptions.logLevel ?? 'warn' });
    const pptxBuffer = fs.readFileSync(pptxPath);
    const renderResult = await renderer.renderPresentation(pptxBuffer, opts.renderOptions);
    // Compare each slide
    const slideReports = [];
    let ssimSum = 0;
    let passedCount = 0;
    for (const slide of renderResult.slides) {
        const slideNumber = slide.slideNumber;
        const baselinePath = path.join(opts.baselinesDir, `slide-${slideNumber}.png`);
        // Check if baseline exists
        if (!fs.existsSync(baselinePath)) {
            slideReports.push({
                slideNumber,
                ssim: 0,
                mse: Infinity,
                psnr: 0,
                pixelDiffPercent: 100,
                baselinePath,
                renderedPath: '',
                passed: false,
                error: `Baseline not found: ${baselinePath}`,
            });
            continue;
        }
        // Save rendered image if output directory specified
        let renderedPath = '';
        if (opts.outputDir && slide.success) {
            renderedPath = path.join(opts.outputDir, `rendered-${slideNumber}.png`);
            fs.writeFileSync(renderedPath, slide.imageData);
        }
        // Handle render failure
        if (!slide.success) {
            slideReports.push({
                slideNumber,
                ssim: 0,
                mse: Infinity,
                psnr: 0,
                pixelDiffPercent: 100,
                baselinePath,
                renderedPath,
                passed: false,
                error: `Render failed: ${slide.errorMessage}`,
            });
            continue;
        }
        // Compare images
        let comparison;
        try {
            comparison = await (0, VisualComparator_js_1.compareImages)(baselinePath, slide.imageData, {
                generateDiffImage: opts.generateDiffs,
                pixelThreshold: opts.pixelThreshold,
            });
        }
        catch (error) {
            slideReports.push({
                slideNumber,
                ssim: 0,
                mse: Infinity,
                psnr: 0,
                pixelDiffPercent: 100,
                baselinePath,
                renderedPath,
                passed: false,
                error: `Comparison failed: ${error instanceof Error ? error.message : String(error)}`,
            });
            continue;
        }
        // Save diff image if generated
        let diffPath;
        if (opts.outputDir && opts.generateDiffs && comparison.diffImage) {
            diffPath = path.join(opts.outputDir, `diff-${slideNumber}.png`);
            fs.writeFileSync(diffPath, comparison.diffImage);
        }
        // Check if passed
        const passed = comparison.ssim >= opts.targetFidelity;
        if (passed) {
            passedCount++;
        }
        ssimSum += comparison.ssim;
        slideReports.push({
            slideNumber,
            ssim: comparison.ssim,
            mse: comparison.mse,
            psnr: comparison.psnr,
            pixelDiffPercent: comparison.pixelDiffPercent,
            baselinePath,
            renderedPath,
            diffPath,
            passed,
        });
    }
    // Calculate overall metrics
    const totalSlides = slideReports.length;
    const overallFidelity = totalSlides > 0 ? ssimSum / totalSlides : 0;
    const allPassed = passedCount === totalSlides && totalSlides > 0;
    const report = {
        pptxPath: path.resolve(pptxPath),
        timestamp: new Date(),
        overallFidelity,
        slides: slideReports,
        passed: allPassed,
        targetFidelity: opts.targetFidelity,
        totalSlides,
        passedSlides: passedCount,
        durationMs: Date.now() - startTime,
    };
    // Save JSON report if output directory specified
    if (opts.outputDir) {
        const reportPath = path.join(opts.outputDir, 'fidelity-report.json');
        fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    }
    return report;
}
/**
 * Formats a fidelity report for console output.
 */
function formatReport(report) {
    const lines = [];
    lines.push('='.repeat(60));
    lines.push('FIDELITY TEST REPORT');
    lines.push('='.repeat(60));
    lines.push('');
    lines.push(`PPTX: ${report.pptxPath}`);
    lines.push(`Timestamp: ${report.timestamp.toISOString()}`);
    lines.push(`Duration: ${report.durationMs}ms`);
    lines.push(`Target Fidelity: ${(report.targetFidelity * 100).toFixed(1)}%`);
    lines.push('');
    lines.push('-'.repeat(60));
    lines.push('SLIDE RESULTS');
    lines.push('-'.repeat(60));
    for (const slide of report.slides) {
        const status = slide.passed ? 'PASS' : 'FAIL';
        const statusIcon = slide.passed ? '[OK]' : '[!!]';
        if (slide.error) {
            lines.push(`${statusIcon} Slide ${slide.slideNumber}: ${status} - ERROR: ${slide.error}`);
        }
        else {
            lines.push(`${statusIcon} Slide ${slide.slideNumber}: ${status} | ` +
                `SSIM: ${(slide.ssim * 100).toFixed(2)}% | ` +
                `MSE: ${slide.mse.toFixed(2)} | ` +
                `PSNR: ${slide.psnr === Infinity ? 'Inf' : slide.psnr.toFixed(2)}dB | ` +
                `Diff: ${slide.pixelDiffPercent.toFixed(2)}%`);
        }
    }
    lines.push('');
    lines.push('-'.repeat(60));
    lines.push('SUMMARY');
    lines.push('-'.repeat(60));
    lines.push(`Overall Fidelity: ${(report.overallFidelity * 100).toFixed(2)}%`);
    lines.push(`Slides Passed: ${report.passedSlides}/${report.totalSlides}`);
    lines.push('');
    if (report.passed) {
        lines.push('RESULT: PASSED');
    }
    else {
        lines.push('RESULT: FAILED');
    }
    lines.push('='.repeat(60));
    return lines.join('\n');
}
/**
 * FidelityTester class for object-oriented usage.
 */
class FidelityTester {
    renderer;
    defaultOptions;
    constructor(options) {
        this.renderer = new PptxImageRenderer_js_1.PptxImageRenderer({ logLevel: options?.renderOptions?.logLevel ?? 'warn' });
        this.defaultOptions = options ?? {};
    }
    /**
     * Runs fidelity tests for a PPTX presentation.
     */
    async test(pptxPath, options) {
        return runFidelityTest(pptxPath, { ...this.defaultOptions, ...options });
    }
    /**
     * Quick check if a PPTX meets a minimum fidelity threshold.
     */
    async passes(pptxPath, baselinesDir, targetFidelity = 0.95) {
        const report = await this.test(pptxPath, {
            baselinesDir,
            targetFidelity,
            generateDiffs: false,
        });
        return report.passed;
    }
}
exports.FidelityTester = FidelityTester;
