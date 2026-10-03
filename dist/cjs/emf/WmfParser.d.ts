/**
 * Parser for basic WMF (Windows Metafile) binaries per [MS-WMF].
 *
 * Supports the placeable (Aldus) header for physical bounds and the
 * META_ record subset that Office paste/OLE content commonly uses,
 * converting everything into the shared MetafileRecord model replayed
 * by EmfRenderer.
 *
 * WMF records are 16-bit-word based: each record is
 * {RecordSize u32 (in words), RecordFunction u16, params u16[]}.
 */
import type { ILogger } from '../utils/Logger.js';
import type { ParsedMetafile } from './types.js';
/**
 * Configuration for WmfParser.
 */
export interface WmfParserConfig {
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Checks whether a buffer looks like a WMF file (placeable header key, or
 * a standard header with a plausible type/size/version combination).
 */
export declare function isWmf(buffer: Buffer): boolean;
/**
 * Parses basic WMF binaries into the shared metafile record model.
 */
export declare class WmfParser {
    private readonly logger;
    constructor(config?: WmfParserConfig);
    /**
     * Parses a WMF binary.
     *
     * @param buffer The WMF file contents
     * @returns The parsed metafile
     * @throws Error if the buffer is not a valid WMF file
     */
    parse(buffer: Buffer): ParsedMetafile;
    /**
     * Wraps the converted records with a synthesized header and the initial
     * window/viewport mapping.
     */
    private buildParsedMetafile;
    /**
     * Parses a single WMF record into the shared model.
     * WMF parameters follow the GDI convention of reverse order.
     */
    private parseRecord;
    /** Allocates the lowest free WMF object-table slot; returns the 1-based index. */
    private allocateSlot;
    /** Reads a WMF (ANSI) LOGFONT at the given offset. */
    private readLogFont16;
    /** Reads a latin1 string of the given byte length. */
    private readAnsiString;
    /**
     * Reads an inline BITMAPINFO + bits blob (DIBSTRETCHBLT/STRETCHDIB),
     * splitting the palette-bearing header from the pixel data.
     */
    private readDib;
    /** Appends a dibDraw record, attaching the DIB when present. */
    private pushDibDraw;
}
