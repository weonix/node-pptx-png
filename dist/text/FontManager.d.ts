/**
 * Font registration for rendering: decodes embedded PPTX font parts
 * (plain TTF/OTF, EOT-wrapped, or ODTTF-obfuscated) and registers them —
 * along with user-supplied fonts — with the skia-canvas FontLibrary.
 *
 * skia-canvas registers fonts by file path only, so font bytes are written
 * to a session temp directory (fs.mkdtemp) before registration. Registration
 * is process-global (FontLibrary has no per-render scope), so successfully
 * registered fonts stay available for subsequent renders; a content-hash
 * cache prevents re-writing and re-registering identical font bytes.
 */
import type { EmbeddedFontData } from '../core/PptxParser.js';
import type { FontRegistration } from '../types/options.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Detected format of decoded font bytes.
 */
export type FontFormat = 'ttf' | 'otf' | 'ttc';
/**
 * Result of decoding a fntdata part: either usable font bytes or a reason
 * the part cannot be used.
 */
export type FntdataDecodeResult = {
    ok: true;
    data: Buffer;
    format: FontFormat;
    container: 'plain' | 'eot' | 'odttf';
} | {
    ok: false;
    reason: string;
};
/**
 * Minimal FontLibrary surface used by FontManager. Injectable for tests.
 */
export interface FontLibraryLike {
    use(familyName: string, fontPaths?: string | readonly string[]): unknown;
    has(familyName: string): boolean;
}
/**
 * Outcome of registering a set of embedded fonts.
 */
export interface EmbeddedFontRegistrationResult {
    /** Family names successfully registered (as declared by p:font/@typeface) */
    registered: string[];
    /** Fonts (or variants) that could not be registered, with reasons */
    skipped: Array<{
        typeface: string;
        reason: string;
    }>;
}
/**
 * Sniffs the format of font bytes from their leading magic.
 *
 * @param data Candidate font bytes
 * @returns 'ttf' | 'otf' | 'ttc', or undefined when not a recognizable font
 */
export declare function sniffFontFormat(data: Buffer): FontFormat | undefined;
/**
 * Deobfuscates an ODTTF font part: the first 32 bytes are XORed with the
 * 16 GUID bytes from the part name in reverse order, repeated twice.
 * XOR is symmetric, so the same operation obfuscates plain font bytes.
 *
 * @param data Obfuscated font bytes (not modified)
 * @param guidBytes The 16 GUID bytes from the part name
 * @returns A new buffer with the leading bytes deobfuscated
 */
export declare function deobfuscateOdttf(data: Buffer, guidBytes: Buffer): Buffer;
/**
 * Decodes the raw bytes of an embedded font part into loadable font bytes.
 *
 * Supported containers, tried in order:
 * 1. Plain TTF/OTF/TTC (PowerPoint's usual fntdata payload)
 * 2. EOT wrapper (plain or XOR-encrypted payload; PowerPoint 365 writes
 *    MicroType Express compressed EOT, which is detected and reported as
 *    unsupported rather than mis-registered)
 * 3. ODTTF obfuscation per ECMA-376 (first 32 bytes XORed with the GUID
 *    from the part name), when the part name contains a GUID
 *
 * @param data Raw bytes of the font part
 * @param partName Path of the part within the package (used for ODTTF GUIDs)
 * @returns Decoded font bytes, or a skip reason
 */
export declare function decodeFntdata(data: Buffer, partName: string): FntdataDecodeResult;
/**
 * Returns the family names registered through FontManager in this process.
 */
export declare function getRegisteredFontFamilies(): ReadonlySet<string>;
/**
 * Clears the process-global registration bookkeeping.
 * Intended for tests only; does not unregister fonts from skia-canvas.
 *
 * @internal
 */
export declare function resetFontRegistrationStateForTesting(): void;
/**
 * Configuration for FontManager.
 */
export interface FontManagerConfig {
    /** Logger instance */
    logger?: ILogger;
    /** FontLibrary implementation (injectable for tests) */
    fontLibrary?: FontLibraryLike;
    /** Structured warning collector for font registration fidelity events */
    warnings?: WarningCollector;
}
/**
 * Registers embedded and user-supplied fonts with skia-canvas.
 */
export declare class FontManager {
    private readonly logger;
    private readonly fontLibrary;
    private readonly warnings;
    private tempDir;
    constructor(config?: FontManagerConfig);
    /**
     * Registers fonts extracted from a PPTX (see PptxParser.getEmbeddedFonts).
     * Variants that cannot be decoded (e.g., MicroType Express compressed EOT)
     * are skipped with a reason; a font counts as registered when at least one
     * of its variants loads.
     *
     * @param fonts Embedded fonts with raw part bytes
     * @returns Registered family names and skipped fonts with reasons
     */
    registerEmbeddedFonts(fonts: EmbeddedFontData[]): Promise<EmbeddedFontRegistrationResult>;
    /**
     * Registers user-supplied fonts (PptxRenderOptions.fonts.register).
     * String sources are registered by path; Buffer sources are written to the
     * session temp directory first (skia-canvas registers by file path only).
     *
     * @param registrations Fonts to register
     * @returns Family names that registered successfully
     */
    registerFonts(registrations: FontRegistration[]): Promise<string[]>;
    /**
     * Removes the session temp directory holding extracted font files.
     * Safe to call after rendering: skia-canvas reads font files eagerly at
     * registration time. Registered fonts remain available to the process.
     */
    cleanup(): Promise<void>;
    /**
     * Writes font bytes to the session temp directory, creating it on first use.
     */
    private writeTempFont;
}
/**
 * Creates a FontManager.
 */
export declare function createFontManager(config?: FontManagerConfig): FontManager;
//# sourceMappingURL=FontManager.d.ts.map