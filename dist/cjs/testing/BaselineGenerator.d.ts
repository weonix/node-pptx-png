/**
 * BaselineGenerator - Generate baseline images from PPTX using LibreOffice
 *
 * Uses LibreOffice headless mode to export PPTX slides to high-quality PNG
 * images that serve as reference baselines for fidelity testing.
 */
/**
 * Result of baseline generation.
 */
export interface BaselineGenerationResult {
    /**
     * Path to the source PPTX file.
     */
    pptxPath: string;
    /**
     * Directory where baselines were saved.
     */
    outputDir: string;
    /**
     * Paths to generated baseline images.
     */
    baselinePaths: string[];
    /**
     * Number of slides processed.
     */
    slideCount: number;
    /**
     * Whether generation succeeded.
     */
    success: boolean;
    /**
     * Error message if generation failed.
     */
    error?: string;
    /**
     * Duration in milliseconds.
     */
    durationMs: number;
}
/**
 * Options for baseline generation.
 */
export interface BaselineGeneratorOptions {
    /**
     * Output directory for baseline images.
     * Default: './test/baselines'
     */
    outputDir?: string;
    /**
     * Path to LibreOffice executable.
     * Default: auto-detect
     */
    libreOfficePath?: string;
    /**
     * Image width in pixels.
     * Default: 1920
     */
    width?: number;
    /**
     * Image height in pixels.
     * Default: 1080
     */
    height?: number;
    /**
     * Timeout for LibreOffice process in milliseconds.
     * Default: 60000 (1 minute)
     */
    timeout?: number;
}
/**
 * Finds the LibreOffice executable path for the current platform.
 *
 * Checks the well-known install locations for the current platform and
 * returns the first one that exists on disk.
 *
 * @returns Absolute path to the soffice executable, or null when LibreOffice
 *          is not installed in any of the known locations
 */
export declare function findLibreOfficePath(): string | null;
/**
 * Generates baseline images from a PPTX file using LibreOffice.
 *
 * @param pptxPath - Path to the PPTX file
 * @param options - Generation options
 * @returns Generation result
 */
export declare function generateBaselines(pptxPath: string, options?: BaselineGeneratorOptions): Promise<BaselineGenerationResult>;
/**
 * Generates baselines using PDF intermediate (for better quality).
 * This is an alternative approach that first converts to PDF, then to PNG.
 */
export declare function generateBaselinesViaPdf(pptxPath: string, options?: BaselineGeneratorOptions): Promise<BaselineGenerationResult>;
/**
 * BaselineGenerator class for object-oriented usage.
 */
export declare class BaselineGenerator {
    private readonly options;
    constructor(options?: BaselineGeneratorOptions);
    /**
     * Generates baseline images from a PPTX file.
     */
    generate(pptxPath: string, outputDir?: string): Promise<BaselineGenerationResult>;
    /**
     * Generates baselines using PDF as intermediate format (higher quality).
     */
    generateViaPdf(pptxPath: string, outputDir?: string): Promise<BaselineGenerationResult>;
    /**
     * Checks if LibreOffice is available.
     */
    isLibreOfficeAvailable(): boolean;
    /**
     * Gets the detected LibreOffice path.
     */
    getLibreOfficePath(): string | null;
}
/**
 * Checks if LibreOffice is available on the system.
 */
export declare function isLibreOfficeAvailable(): boolean;
