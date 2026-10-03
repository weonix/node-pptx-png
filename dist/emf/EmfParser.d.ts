/**
 * Parser for EMF (Enhanced Metafile) binaries per [MS-EMF].
 *
 * Parses the GDI record subset that Office paste content (pasted Excel
 * tables/charts, OLE previews) uses, producing the device-independent
 * MetafileRecord model replayed by EmfRenderer.
 *
 * EMF+ comment records ([MS-EMFPLUS]) are detected but skipped: Office
 * writes EMF+ Dual files that carry a complete GDI fallback stream, so
 * rendering the GDI records reproduces the image. EMF+-only files are
 * flagged so callers can warn about reduced fidelity.
 */
import type { ILogger } from '../utils/Logger.js';
import type { ParsedMetafile } from './types.js';
/**
 * Configuration for EmfParser.
 */
export interface EmfParserConfig {
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Checks whether a buffer looks like an EMF file (EMR_HEADER with the
 * ' EMF' signature at offset 40).
 */
export declare function isEmf(buffer: Buffer): boolean;
/**
 * Parses EMF binaries into the shared metafile record model.
 */
export declare class EmfParser {
    private readonly logger;
    constructor(config?: EmfParserConfig);
    /**
     * Parses an EMF binary.
     *
     * @param buffer The EMF file contents
     * @returns The parsed metafile
     * @throws Error if the buffer is not a valid EMF file
     */
    parse(buffer: Buffer): ParsedMetafile;
    /**
     * Parses a single record, appending to result.records.
     * @returns true when EMR_EOF is reached
     */
    private parseRecord;
    /**
     * Parses EMR_EXTSELECTCLIPRGN. Only RDH_RECTANGLES region data is
     * supported; an empty region with RGN_COPY resets the clip.
     */
    private parseExtSelectClipRgn;
    /**
     * Parses EMR_EXTTEXTOUTW / EMR_EXTTEXTOUTA.
     */
    private parseExtTextOut;
    /**
     * Parses EMR_BITBLT / EMR_STRETCHBLT into a dibDraw record.
     * BitBlt without source bits is commonly a PATCOPY brush fill.
     */
    private parseBltRecord;
    /**
     * Parses EMR_STRETCHDIBITS into a dibDraw record.
     */
    private parseStretchDiBits;
    /**
     * Reads an embedded DIB (BITMAPINFO + bits) referenced by offset/size
     * fields within a record. Returns undefined when no bits are present.
     */
    private readEmbeddedDib;
    /**
     * Parses EMR_COMMENT, detecting EMF+ streams. EMF+ payloads are skipped
     * (Office writes EMF+ Dual with a complete GDI fallback).
     */
    private parseComment;
}
//# sourceMappingURL=EmfParser.d.ts.map