"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.WmfParser = void 0;
exports.isWmf = isWmf;
const Logger_js_1 = require("../utils/Logger.js");
const types_js_1 = require("./types.js");
/** Placeable WMF header magic number. */
const PLACEABLE_KEY = 0x9ac6cdd7;
/** WMF record function numbers ([MS-WMF] 2.3). */
const META = {
    EOF: 0x0000,
    SETBKCOLOR: 0x0201,
    SETBKMODE: 0x0102,
    SETMAPMODE: 0x0103,
    SETROP2: 0x0104,
    SETPOLYFILLMODE: 0x0106,
    SETSTRETCHBLTMODE: 0x0107,
    SETTEXTCOLOR: 0x0209,
    SETTEXTALIGN: 0x012e,
    SETWINDOWORG: 0x020b,
    SETWINDOWEXT: 0x020c,
    SAVEDC: 0x001e,
    RESTOREDC: 0x0127,
    INTERSECTCLIPRECT: 0x0416,
    MOVETO: 0x0214,
    LINETO: 0x0213,
    RECTANGLE: 0x041b,
    ELLIPSE: 0x0418,
    POLYGON: 0x0324,
    POLYLINE: 0x0325,
    POLYPOLYGON: 0x0538,
    CREATEPENINDIRECT: 0x02fa,
    CREATEBRUSHINDIRECT: 0x02fc,
    CREATEFONTINDIRECT: 0x02fb,
    SELECTOBJECT: 0x012d,
    DELETEOBJECT: 0x01f0,
    TEXTOUT: 0x0521,
    EXTTEXTOUT: 0x0a32,
    DIBSTRETCHBLT: 0x0b41,
    STRETCHDIB: 0x0f43,
};
/** ExtTextOut option flags relevant to the record layout. */
const ETO_OPAQUE = 0x02;
const ETO_CLIPPED = 0x04;
/**
 * Checks whether a buffer looks like a WMF file (placeable header key, or
 * a standard header with a plausible type/size/version combination).
 */
function isWmf(buffer) {
    if (buffer.length >= 4 && buffer.readUInt32LE(0) === PLACEABLE_KEY) {
        return true;
    }
    if (buffer.length >= 18) {
        const type = buffer.readUInt16LE(0);
        const headerSize = buffer.readUInt16LE(2);
        const version = buffer.readUInt16LE(4);
        return ((type === 1 || type === 2) && headerSize === 9 && (version === 0x0100 || version === 0x0300));
    }
    return false;
}
/**
 * Parses basic WMF binaries into the shared metafile record model.
 */
class WmfParser {
    logger;
    constructor(config = {}) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'WmfParser');
    }
    /**
     * Parses a WMF binary.
     *
     * @param buffer The WMF file contents
     * @returns The parsed metafile
     * @throws Error if the buffer is not a valid WMF file
     */
    parse(buffer) {
        if (!isWmf(buffer)) {
            throw new Error('Not a WMF file');
        }
        const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
        let headerStart = 0;
        let bbox;
        let unitsPerInch = 1440;
        if (view.getUint32(0, true) === PLACEABLE_KEY) {
            if (buffer.length < 22 + 18) {
                throw new Error('Truncated placeable WMF');
            }
            bbox = {
                left: view.getInt16(6, true),
                top: view.getInt16(8, true),
                right: view.getInt16(10, true),
                bottom: view.getInt16(12, true),
            };
            const inch = view.getUint16(14, true);
            if (inch > 0) {
                unitsPerInch = inch;
            }
            headerStart = 22;
        }
        const records = [];
        // Object table: WMF objects occupy the lowest free slot; we map slot n
        // to shared-model index n+1 (EMF indices are 1-based)
        const slots = [];
        const state = { truncated: false, sawEof: false };
        let offset = headerStart + 18;
        while (offset + 6 <= buffer.length) {
            const sizeWords = view.getUint32(offset, true);
            const func = view.getUint16(offset + 4, true);
            const sizeBytes = sizeWords * 2;
            if (func === META.EOF) {
                records.push({ type: 'eof' });
                state.sawEof = true;
                break;
            }
            if (sizeWords < 3 || offset + sizeBytes > buffer.length) {
                this.logger.warn('Malformed WMF record; stopping parse', { offset, func, sizeWords });
                state.truncated = true;
                break;
            }
            try {
                this.parseRecord(view, offset, sizeBytes, func, records, slots);
            }
            catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                this.logger.warn('Failed to parse WMF record; skipping', { offset, func, error: message });
            }
            offset += sizeBytes;
        }
        if (!state.sawEof && !state.truncated) {
            state.truncated = true;
            this.logger.warn('WMF stream ended without META_EOF');
        }
        return this.buildParsedMetafile(records, bbox, unitsPerInch, state.truncated);
    }
    /**
     * Wraps the converted records with a synthesized header and the initial
     * window/viewport mapping.
     */
    buildParsedMetafile(records, bbox, unitsPerInch, truncated) {
        // Determine logical bounds: placeable bbox, else the first window
        // org/ext in the stream, else a default square
        let bounds = bbox;
        if (!bounds) {
            let org = { x: 0, y: 0 };
            let ext;
            for (const record of records) {
                if (record.type === 'setWindowOrg') {
                    org = { x: record.x, y: record.y };
                }
                else if (record.type === 'setWindowExt') {
                    ext = { cx: record.cx, cy: record.cy };
                    break;
                }
            }
            bounds = ext
                ? { left: org.x, top: org.y, right: org.x + ext.cx, bottom: org.y + ext.cy }
                : { left: 0, top: 0, right: 1024, bottom: 1024 };
        }
        const logicalWidth = Math.max(1, Math.abs(bounds.right - bounds.left));
        const logicalHeight = Math.max(1, Math.abs(bounds.bottom - bounds.top));
        // Natural pixel size at 96 dpi from the physical size
        const widthPx = Math.max(1, Math.round((logicalWidth / unitsPerInch) * 96));
        const heightPx = Math.max(1, Math.round((logicalHeight / unitsPerInch) * 96));
        const widthMm = (logicalWidth / unitsPerInch) * 25.4;
        const heightMm = (logicalHeight / unitsPerInch) * 25.4;
        // Prepend the initial logical->device mapping; later SETWINDOWORG/EXT
        // records adjust it against the fixed viewport
        const prelude = [
            { type: 'setMapMode', mode: types_js_1.MAP_MODE.ANISOTROPIC },
            { type: 'setViewportOrg', x: 0, y: 0 },
            { type: 'setViewportExt', cx: widthPx, cy: heightPx },
            { type: 'setWindowOrg', x: bounds.left, y: bounds.top },
            { type: 'setWindowExt', cx: bounds.right - bounds.left, cy: bounds.bottom - bounds.top },
        ];
        return {
            header: {
                bounds: { left: 0, top: 0, right: widthPx - 1, bottom: heightPx - 1 },
                frame: {
                    left: 0,
                    top: 0,
                    right: Math.round(widthMm * 100),
                    bottom: Math.round(heightMm * 100),
                },
                devicePx: { cx: widthPx, cy: heightPx },
                deviceMm: { cx: widthMm, cy: heightMm },
            },
            records: [...prelude, ...records],
            source: 'wmf',
            emfPlusDetected: false,
            emfPlusDual: false,
            truncated,
        };
    }
    /**
     * Parses a single WMF record into the shared model.
     * WMF parameters follow the GDI convention of reverse order.
     */
    parseRecord(view, off, size, func, records, slots) {
        const p = off + 6; // params start
        switch (func) {
            case META.SETMAPMODE:
                // The stream's map mode is ignored: the synthesized prelude fixes
                // the logical->device mapping for the whole file
                break;
            case META.SETWINDOWORG:
                records.push({
                    type: 'setWindowOrg',
                    x: view.getInt16(p + 2, true),
                    y: view.getInt16(p, true),
                });
                break;
            case META.SETWINDOWEXT:
                records.push({
                    type: 'setWindowExt',
                    cx: view.getInt16(p + 2, true),
                    cy: view.getInt16(p, true),
                });
                break;
            case META.SETBKCOLOR:
                records.push({ type: 'setBkColor', color: view.getUint32(p, true) });
                break;
            case META.SETBKMODE:
                records.push({ type: 'setBkMode', mode: view.getUint16(p, true) });
                break;
            case META.SETTEXTCOLOR:
                records.push({ type: 'setTextColor', color: view.getUint32(p, true) });
                break;
            case META.SETTEXTALIGN:
                records.push({ type: 'setTextAlign', align: view.getUint16(p, true) });
                break;
            case META.SETPOLYFILLMODE:
                records.push({ type: 'setPolyFillMode', mode: view.getUint16(p, true) });
                break;
            case META.SETROP2:
                records.push({ type: 'setRop2', mode: view.getUint16(p, true) });
                break;
            case META.SETSTRETCHBLTMODE:
                records.push({ type: 'setStretchBltMode', mode: view.getUint16(p, true) });
                break;
            case META.SAVEDC:
                records.push({ type: 'saveDc' });
                break;
            case META.RESTOREDC:
                records.push({ type: 'restoreDc', saved: view.getInt16(p, true) });
                break;
            case META.INTERSECTCLIPRECT:
                records.push({
                    type: 'intersectClipRect',
                    rect: {
                        bottom: view.getInt16(p, true),
                        right: view.getInt16(p + 2, true),
                        top: view.getInt16(p + 4, true),
                        left: view.getInt16(p + 6, true),
                    },
                });
                break;
            case META.MOVETO:
                records.push({
                    type: 'moveTo',
                    x: view.getInt16(p + 2, true),
                    y: view.getInt16(p, true),
                });
                break;
            case META.LINETO:
                records.push({
                    type: 'lineTo',
                    x: view.getInt16(p + 2, true),
                    y: view.getInt16(p, true),
                });
                break;
            case META.RECTANGLE:
                records.push({
                    type: 'rectangle',
                    rect: {
                        bottom: view.getInt16(p, true),
                        right: view.getInt16(p + 2, true),
                        top: view.getInt16(p + 4, true),
                        left: view.getInt16(p + 6, true),
                    },
                });
                break;
            case META.ELLIPSE:
                records.push({
                    type: 'ellipse',
                    rect: {
                        bottom: view.getInt16(p, true),
                        right: view.getInt16(p + 2, true),
                        top: view.getInt16(p + 4, true),
                        left: view.getInt16(p + 6, true),
                    },
                });
                break;
            case META.POLYGON:
            case META.POLYLINE: {
                const count = view.getUint16(p, true);
                const points = [];
                for (let i = 0; i < count; i++) {
                    const pointOff = p + 2 + i * 4;
                    if (pointOff + 4 > off + size) {
                        break;
                    }
                    points.push({
                        x: view.getInt16(pointOff, true),
                        y: view.getInt16(pointOff + 2, true),
                    });
                }
                records.push({ type: func === META.POLYGON ? 'polygon' : 'polyline', points });
                break;
            }
            case META.POLYPOLYGON: {
                const numPolys = view.getUint16(p, true);
                const counts = [];
                for (let i = 0; i < numPolys; i++) {
                    if (p + 2 + i * 2 + 2 > off + size) {
                        break;
                    }
                    counts.push(view.getUint16(p + 2 + i * 2, true));
                }
                const polygons = [];
                let pointOff = p + 2 + numPolys * 2;
                for (const count of counts) {
                    const poly = [];
                    for (let i = 0; i < count; i++) {
                        if (pointOff + 4 > off + size) {
                            break;
                        }
                        poly.push({
                            x: view.getInt16(pointOff, true),
                            y: view.getInt16(pointOff + 2, true),
                        });
                        pointOff += 4;
                    }
                    polygons.push(poly);
                }
                records.push({ type: 'polyPolygon', polygons });
                break;
            }
            case META.CREATEPENINDIRECT: {
                const index = this.allocateSlot(slots);
                records.push({
                    type: 'createPen',
                    index,
                    style: view.getUint16(p, true),
                    width: view.getInt16(p + 2, true),
                    color: view.getUint32(p + 6, true),
                });
                break;
            }
            case META.CREATEBRUSHINDIRECT: {
                const index = this.allocateSlot(slots);
                records.push({
                    type: 'createBrush',
                    index,
                    style: view.getUint16(p, true),
                    color: view.getUint32(p + 2, true),
                });
                break;
            }
            case META.CREATEFONTINDIRECT: {
                const index = this.allocateSlot(slots);
                records.push({ type: 'createFont', index, font: this.readLogFont16(view, p, off + size) });
                break;
            }
            case META.SELECTOBJECT:
                records.push({ type: 'selectObject', index: view.getUint16(p, true) + 1 });
                break;
            case META.DELETEOBJECT: {
                const slot = view.getUint16(p, true);
                if (slot < slots.length) {
                    slots[slot] = false;
                }
                records.push({ type: 'deleteObject', index: slot + 1 });
                break;
            }
            case META.TEXTOUT: {
                const length = view.getUint16(p, true);
                const textEnd = p + 2 + length;
                if (textEnd + 4 > off + size) {
                    break;
                }
                const text = this.readAnsiString(view, p + 2, length);
                const paddedEnd = p + 2 + length + (length % 2);
                records.push({
                    type: 'textOut',
                    y: view.getInt16(paddedEnd, true),
                    x: view.getInt16(paddedEnd + 2, true),
                    text,
                    options: 0,
                });
                break;
            }
            case META.EXTTEXTOUT: {
                const y = view.getInt16(p, true);
                const x = view.getInt16(p + 2, true);
                const length = view.getUint16(p + 4, true);
                const options = view.getUint16(p + 6, true);
                let cursor = p + 8;
                let rect;
                if ((options & (ETO_OPAQUE | ETO_CLIPPED)) !== 0) {
                    rect = {
                        left: view.getInt16(cursor, true),
                        top: view.getInt16(cursor + 2, true),
                        right: view.getInt16(cursor + 4, true),
                        bottom: view.getInt16(cursor + 6, true),
                    };
                    cursor += 8;
                }
                if (cursor + length > off + size) {
                    break;
                }
                const text = this.readAnsiString(view, cursor, length);
                cursor += length + (length % 2);
                let dx;
                if (cursor + length * 2 <= off + size && length > 0) {
                    dx = [];
                    for (let i = 0; i < length; i++) {
                        dx.push(view.getInt16(cursor + i * 2, true));
                    }
                }
                const record = {
                    type: 'textOut',
                    x,
                    y,
                    text,
                    options,
                };
                if (rect) {
                    record.rect = rect;
                }
                if (dx) {
                    record.dx = dx;
                }
                records.push(record);
                break;
            }
            case META.DIBSTRETCHBLT: {
                const rop = view.getUint32(p, true);
                const srcHeight = view.getInt16(p + 4, true);
                const srcWidth = view.getInt16(p + 6, true);
                const srcY = view.getInt16(p + 8, true);
                const srcX = view.getInt16(p + 10, true);
                const destHeight = view.getInt16(p + 12, true);
                const destWidth = view.getInt16(p + 14, true);
                const destY = view.getInt16(p + 16, true);
                const destX = view.getInt16(p + 18, true);
                const dib = this.readDib(view, p + 20, off + size);
                this.pushDibDraw(records, { destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, rop }, dib);
                break;
            }
            case META.STRETCHDIB: {
                const rop = view.getUint32(p, true);
                const srcHeight = view.getInt16(p + 6, true);
                const srcWidth = view.getInt16(p + 8, true);
                const srcY = view.getInt16(p + 10, true);
                const srcX = view.getInt16(p + 12, true);
                const destHeight = view.getInt16(p + 14, true);
                const destWidth = view.getInt16(p + 16, true);
                const destY = view.getInt16(p + 18, true);
                const destX = view.getInt16(p + 20, true);
                const dib = this.readDib(view, p + 22, off + size);
                this.pushDibDraw(records, { destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, rop }, dib);
                break;
            }
            default:
                this.logger.debug('Skipping unsupported WMF record', { func, size });
                break;
        }
    }
    /** Allocates the lowest free WMF object-table slot; returns the 1-based index. */
    allocateSlot(slots) {
        for (let i = 0; i < slots.length; i++) {
            if (!slots[i]) {
                slots[i] = true;
                return i + 1;
            }
        }
        slots.push(true);
        return slots.length;
    }
    /** Reads a WMF (ANSI) LOGFONT at the given offset. */
    readLogFont16(view, off, end) {
        const height = view.getInt16(off, true);
        const weight = view.getInt16(off + 8, true);
        const italic = view.getUint8(off + 10) !== 0;
        const underline = view.getUint8(off + 11) !== 0;
        const strikeOut = view.getUint8(off + 12) !== 0;
        let faceName = '';
        for (let i = 0; i < 32; i++) {
            const at = off + 18 + i;
            if (at >= end) {
                break;
            }
            const code = view.getUint8(at);
            if (code === 0) {
                break;
            }
            faceName += String.fromCharCode(code);
        }
        return { height, weight, italic, underline, strikeOut, faceName };
    }
    /** Reads a latin1 string of the given byte length. */
    readAnsiString(view, off, length) {
        let text = '';
        for (let i = 0; i < length; i++) {
            text += String.fromCharCode(view.getUint8(off + i));
        }
        return text;
    }
    /**
     * Reads an inline BITMAPINFO + bits blob (DIBSTRETCHBLT/STRETCHDIB),
     * splitting the palette-bearing header from the pixel data.
     */
    readDib(view, off, end) {
        if (off + 40 > end) {
            return undefined;
        }
        const headerSize = view.getUint32(off, true);
        if (headerSize < 40 || off + headerSize > end) {
            return undefined;
        }
        const bitCount = view.getUint16(off + 14, true);
        const compression = view.getUint32(off + 16, true);
        const clrUsed = view.getUint32(off + 32, true);
        let paletteEntries = clrUsed;
        if (paletteEntries === 0 && bitCount <= 8) {
            paletteEntries = 1 << bitCount;
        }
        let bmiSize = headerSize + paletteEntries * 4;
        if (compression === 3 && headerSize === 40) {
            bmiSize += 12; // BI_BITFIELDS masks
        }
        if (off + bmiSize >= end) {
            return undefined;
        }
        const base = view.byteOffset + off;
        return {
            bmi: new Uint8Array(view.buffer.slice(base, base + bmiSize)),
            bits: new Uint8Array(view.buffer.slice(base + bmiSize, view.byteOffset + end)),
        };
    }
    /** Appends a dibDraw record, attaching the DIB when present. */
    pushDibDraw(records, geometry, dib) {
        const record = {
            type: 'dibDraw',
            ...geometry,
        };
        if (dib) {
            record.dib = dib;
        }
        records.push(record);
    }
}
exports.WmfParser = WmfParser;
