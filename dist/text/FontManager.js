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
import { FontLibrary } from 'skia-canvas';
import * as fs from 'fs/promises';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'crypto';
import { createLogger } from '../utils/Logger.js';
/** EOT magic number (little-endian at byte offset 34). */
const EOT_MAGIC = 0x504c;
/** Known EOT header versions. */
const EOT_VERSIONS = new Set([0x00010000, 0x00020001, 0x00020002]);
/** EOT flag: font data is MicroType Express (MTX/LZCOMP) compressed. */
const EOT_FLAG_COMPRESSED = 0x00000004;
/** EOT flag: font data is XOR-encrypted with 0x50. */
const EOT_FLAG_XOR = 0x10000000;
/** Number of leading bytes obfuscated in an ODTTF font part. */
const ODTTF_OBFUSCATED_BYTES = 32;
/**
 * Sniffs the format of font bytes from their leading magic.
 *
 * @param data Candidate font bytes
 * @returns 'ttf' | 'otf' | 'ttc', or undefined when not a recognizable font
 */
export function sniffFontFormat(data) {
    if (data.length < 4) {
        return undefined;
    }
    const magic = data.readUInt32BE(0);
    if (magic === 0x00010000 || magic === 0x74727565 /* 'true' (Apple TTF) */) {
        return 'ttf';
    }
    if (magic === 0x4f54544f /* 'OTTO' */) {
        return 'otf';
    }
    if (magic === 0x74746366 /* 'ttcf' */) {
        return 'ttc';
    }
    return undefined;
}
/**
 * Extracts the 16 GUID bytes from a part name containing a GUID
 * (e.g., 'ppt/fonts/{1F5D6DFE-...}.odttf'). Returns undefined when the
 * name carries no GUID.
 */
function extractGuidBytes(partName) {
    const match = partName.match(/([0-9a-fA-F]{8})-?([0-9a-fA-F]{4})-?([0-9a-fA-F]{4})-?([0-9a-fA-F]{4})-?([0-9a-fA-F]{12})/);
    if (!match) {
        return undefined;
    }
    return Buffer.from(match.slice(1).join(''), 'hex');
}
/**
 * Deobfuscates an ODTTF font part: the first 32 bytes are XORed with the
 * 16 GUID bytes from the part name in reverse order, repeated twice.
 * XOR is symmetric, so the same operation obfuscates plain font bytes.
 *
 * @param data Obfuscated font bytes (not modified)
 * @param guidBytes The 16 GUID bytes from the part name
 * @returns A new buffer with the leading bytes deobfuscated
 */
export function deobfuscateOdttf(data, guidBytes) {
    const out = Buffer.from(data);
    const limit = Math.min(ODTTF_OBFUSCATED_BYTES, out.length);
    for (let i = 0; i < limit; i++) {
        const keyByte = guidBytes[guidBytes.length - 1 - (i % guidBytes.length)];
        out[i] = (out[i] ?? 0) ^ (keyByte ?? 0);
    }
    return out;
}
/**
 * Unwraps an Embedded OpenType (EOT) container, returning the enclosed
 * TTF/OTF bytes. Handles the XOR-encrypted variant; MicroType Express
 * compressed payloads are reported as unsupported.
 */
function unwrapEot(data) {
    const eotSize = data.readUInt32LE(0);
    const fontDataSize = data.readUInt32LE(4);
    const version = data.readUInt32LE(8);
    const flags = data.readUInt32LE(12);
    if (!EOT_VERSIONS.has(version)) {
        return { ok: false, reason: `Unknown EOT version 0x${version.toString(16)}` };
    }
    if (eotSize !== data.length || fontDataSize > data.length) {
        return { ok: false, reason: 'Corrupt EOT header (size mismatch)' };
    }
    if ((flags & EOT_FLAG_COMPRESSED) !== 0) {
        return {
            ok: false,
            reason: 'EOT font data is MicroType Express compressed (not supported)',
        };
    }
    // FontData is the last field of the EOT structure
    let fontData = data.subarray(data.length - fontDataSize);
    if ((flags & EOT_FLAG_XOR) !== 0) {
        const decoded = Buffer.from(fontData);
        for (let i = 0; i < decoded.length; i++) {
            decoded[i] = (decoded[i] ?? 0) ^ 0x50;
        }
        fontData = decoded;
    }
    const format = sniffFontFormat(fontData);
    if (!format) {
        return { ok: false, reason: 'EOT font data is not a recognizable TTF/OTF' };
    }
    return { ok: true, data: Buffer.from(fontData), format, container: 'eot' };
}
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
export function decodeFntdata(data, partName) {
    if (data.length < 16) {
        return { ok: false, reason: 'Font part is too small to be a font' };
    }
    // 1. Plain TTF/OTF/TTC
    const plainFormat = sniffFontFormat(data);
    if (plainFormat) {
        return { ok: true, data, format: plainFormat, container: 'plain' };
    }
    // 2. EOT wrapper: magic 0x504C at offset 34, total-size field at offset 0
    if (data.length >= 36 &&
        data.readUInt16LE(34) === EOT_MAGIC &&
        data.readUInt32LE(0) === data.length) {
        return unwrapEot(data);
    }
    // 3. ODTTF obfuscation (requires a GUID in the part name)
    const guidBytes = extractGuidBytes(partName);
    if (guidBytes) {
        const deobfuscated = deobfuscateOdttf(data, guidBytes);
        const format = sniffFontFormat(deobfuscated);
        if (format) {
            return { ok: true, data: deobfuscated, format, container: 'odttf' };
        }
    }
    return { ok: false, reason: 'Unrecognized font container format' };
}
/**
 * Process-global set of family names registered through FontManager
 * (embedded or user-supplied). FontResolver consults this set when checking
 * font availability, so registered families are preferred exactly.
 */
const registeredFamilies = new Set();
/**
 * Process-global cache of registered font content hashes, so identical font
 * bytes (e.g., the same deck rendered twice) are not re-written and
 * re-registered.
 */
const registeredContentHashes = new Set();
/**
 * Returns the family names registered through FontManager in this process.
 */
export function getRegisteredFontFamilies() {
    return registeredFamilies;
}
/**
 * Clears the process-global registration bookkeeping.
 * Intended for tests only; does not unregister fonts from skia-canvas.
 *
 * @internal
 */
export function resetFontRegistrationStateForTesting() {
    registeredFamilies.clear();
    registeredContentHashes.clear();
}
/**
 * Registers embedded and user-supplied fonts with skia-canvas.
 */
export class FontManager {
    logger;
    fontLibrary;
    warnings;
    tempDir = null;
    constructor(config = {}) {
        this.logger = config.logger ?? createLogger('warn', 'FontManager');
        this.fontLibrary = config.fontLibrary ?? FontLibrary;
        this.warnings = config.warnings;
    }
    /**
     * Registers fonts extracted from a PPTX (see PptxParser.getEmbeddedFonts).
     * Variants that cannot be decoded (e.g., MicroType Express compressed EOT)
     * are skipped with a reason; a font counts as registered when at least one
     * of its variants loads.
     *
     * @param fonts Embedded fonts with raw part bytes
     * @returns Registered family names and skipped fonts with reasons
     */
    async registerEmbeddedFonts(fonts) {
        const result = { registered: [], skipped: [] };
        for (const font of fonts) {
            const paths = [];
            const variantReasons = [];
            let alreadyRegistered = false;
            for (const variant of font.variants) {
                const hash = hashContent(font.typeface, variant.data);
                if (registeredContentHashes.has(hash)) {
                    alreadyRegistered = true;
                    continue;
                }
                const decoded = decodeFntdata(variant.data, variant.path);
                if (!decoded.ok) {
                    variantReasons.push(`${variant.style}: ${decoded.reason}`);
                    this.logger.warn('Skipping embedded font variant', {
                        typeface: font.typeface,
                        style: variant.style,
                        reason: decoded.reason,
                    });
                    this.warnings?.push({
                        code: 'embedded-font-unsupported',
                        message: `Embedded font "${font.typeface}" (${variant.style}) could not be used: ${decoded.reason}`,
                        detail: { typeface: font.typeface, style: variant.style, reason: decoded.reason },
                    });
                    continue;
                }
                try {
                    const filePath = await this.writeTempFont(`${sanitizeFileName(font.typeface)}-${variant.style}.${decoded.format}`, decoded.data);
                    paths.push(filePath);
                    registeredContentHashes.add(hash);
                }
                catch (error) {
                    variantReasons.push(`${variant.style}: ${error instanceof Error ? error.message : String(error)}`);
                }
            }
            if (paths.length === 0) {
                if (alreadyRegistered) {
                    // Same bytes were registered earlier in this process
                    registeredFamilies.add(font.typeface);
                    result.registered.push(font.typeface);
                }
                else {
                    result.skipped.push({
                        typeface: font.typeface,
                        reason: variantReasons.join('; ') || 'No variants could be decoded',
                    });
                }
                continue;
            }
            try {
                this.fontLibrary.use(font.typeface, paths);
                registeredFamilies.add(font.typeface);
                result.registered.push(font.typeface);
                this.logger.info('Registered embedded font', {
                    typeface: font.typeface,
                    variants: paths.length,
                });
            }
            catch (error) {
                result.skipped.push({
                    typeface: font.typeface,
                    reason: error instanceof Error ? error.message : String(error),
                });
                this.logger.warn('FontLibrary rejected embedded font', {
                    typeface: font.typeface,
                    error: error instanceof Error ? error.message : String(error),
                });
                this.warnings?.push({
                    code: 'embedded-font-unsupported',
                    message: `Embedded font "${font.typeface}" was rejected by the font library: ${error instanceof Error ? error.message : String(error)}`,
                    detail: { typeface: font.typeface },
                });
            }
        }
        return result;
    }
    /**
     * Registers user-supplied fonts (PptxRenderOptions.fonts.register).
     * String sources are registered by path; Buffer sources are written to the
     * session temp directory first (skia-canvas registers by file path only).
     *
     * @param registrations Fonts to register
     * @returns Family names that registered successfully
     */
    async registerFonts(registrations) {
        const registered = [];
        for (const registration of registrations) {
            try {
                let filePath;
                if (typeof registration.source === 'string') {
                    filePath = registration.source;
                }
                else {
                    const format = sniffFontFormat(registration.source) ?? 'ttf';
                    filePath = await this.writeTempFont(`${sanitizeFileName(registration.family)}-${hashContent(registration.family, registration.source).slice(0, 8)}.${format}`, registration.source);
                }
                this.fontLibrary.use(registration.family, [filePath]);
                registeredFamilies.add(registration.family);
                registered.push(registration.family);
                this.logger.info('Registered user font', { family: registration.family });
            }
            catch (error) {
                this.logger.warn('Failed to register user font', {
                    family: registration.family,
                    error: error instanceof Error ? error.message : String(error),
                });
                this.warnings?.push({
                    code: 'other',
                    message: `Failed to register user font "${registration.family}": ${error instanceof Error ? error.message : String(error)}`,
                    detail: { family: registration.family },
                });
            }
        }
        return registered;
    }
    /**
     * Removes the session temp directory holding extracted font files.
     * Safe to call after rendering: skia-canvas reads font files eagerly at
     * registration time. Registered fonts remain available to the process.
     */
    async cleanup() {
        if (this.tempDir) {
            const dir = this.tempDir;
            this.tempDir = null;
            try {
                await fs.rm(dir, { recursive: true, force: true });
            }
            catch (error) {
                this.logger.warn('Failed to remove font temp directory', {
                    dir,
                    error: error instanceof Error ? error.message : String(error),
                });
            }
        }
    }
    /**
     * Writes font bytes to the session temp directory, creating it on first use.
     */
    async writeTempFont(fileName, data) {
        this.tempDir ??= await fs.mkdtemp(path.join(os.tmpdir(), 'node-pptx-png-fonts-'));
        const filePath = path.join(this.tempDir, fileName);
        await fs.writeFile(filePath, data);
        return filePath;
    }
}
/**
 * Creates a FontManager.
 */
export function createFontManager(config = {}) {
    return new FontManager(config);
}
/** Hashes font bytes (namespaced by family) for the dedupe cache. */
function hashContent(family, data) {
    return createHash('sha256').update(family).update('\0').update(data).digest('hex');
}
/** Makes a font family name safe for use in a file name. */
function sanitizeFileName(name) {
    return name.replace(/[^a-zA-Z0-9._-]+/g, '_');
}
//# sourceMappingURL=FontManager.js.map