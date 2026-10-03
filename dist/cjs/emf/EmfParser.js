"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.EmfParser = void 0;
exports.isEmf = isEmf;
const Logger_js_1 = require("../utils/Logger.js");
const types_js_1 = require("./types.js");
/** EMF header signature: ' EMF' as a little-endian u32 at offset 40. */
const EMF_SIGNATURE = 0x464d4520;
/** EMF+ comment identifier: 'EMF+' as a little-endian u32. */
const EMFPLUS_COMMENT_ID = 0x2b464d45;
/** EmfPlusHeader record type ([MS-EMFPLUS] 2.3.3.3). */
const EMFPLUS_HEADER_TYPE = 0x4001;
/** Minimum EMF header size (through szlMillimeters). */
const MIN_HEADER_SIZE = 88;
/**
 * Checks whether a buffer looks like an EMF file (EMR_HEADER with the
 * ' EMF' signature at offset 40).
 */
function isEmf(buffer) {
    return (buffer.length >= 44 &&
        buffer.readUInt32LE(0) === types_js_1.EMR.HEADER &&
        buffer.readUInt32LE(40) === EMF_SIGNATURE);
}
/**
 * Parses EMF binaries into the shared metafile record model.
 */
class EmfParser {
    logger;
    constructor(config = {}) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'EmfParser');
    }
    /**
     * Parses an EMF binary.
     *
     * @param buffer The EMF file contents
     * @returns The parsed metafile
     * @throws Error if the buffer is not a valid EMF file
     */
    parse(buffer) {
        if (buffer.length < MIN_HEADER_SIZE) {
            throw new Error(`EMF buffer too small: ${buffer.length} bytes`);
        }
        const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
        const iType = view.getUint32(0, true);
        const headerSize = view.getUint32(4, true);
        if (iType !== types_js_1.EMR.HEADER || view.getUint32(40, true) !== EMF_SIGNATURE) {
            throw new Error('Not an EMF file: missing EMR_HEADER / signature');
        }
        if (headerSize < MIN_HEADER_SIZE || headerSize > buffer.length) {
            throw new Error(`Invalid EMF header size: ${headerSize}`);
        }
        const result = {
            header: {
                bounds: readRect(view, 8),
                frame: readRect(view, 24),
                devicePx: { cx: view.getUint32(72, true), cy: view.getUint32(76, true) },
                deviceMm: { cx: view.getUint32(80, true), cy: view.getUint32(84, true) },
            },
            records: [],
            source: 'emf',
            emfPlusDetected: false,
            emfPlusDual: false,
            truncated: false,
        };
        let offset = headerSize;
        while (offset + 8 <= buffer.length) {
            const recordType = view.getUint32(offset, true);
            const recordSize = view.getUint32(offset + 4, true);
            if (recordSize < 8 || recordSize % 4 !== 0 || offset + recordSize > buffer.length) {
                this.logger.warn('Malformed EMF record; stopping parse', {
                    offset,
                    recordType,
                    recordSize,
                });
                result.truncated = true;
                break;
            }
            try {
                const done = this.parseRecord(view, offset, recordType, recordSize, result);
                if (done) {
                    return result;
                }
            }
            catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                this.logger.warn('Failed to parse EMF record; skipping', {
                    offset,
                    recordType,
                    error: message,
                });
            }
            offset += recordSize;
        }
        // Fell off the end without EMR_EOF
        result.truncated = true;
        this.logger.warn('EMF stream ended without EMR_EOF');
        return result;
    }
    /**
     * Parses a single record, appending to result.records.
     * @returns true when EMR_EOF is reached
     */
    parseRecord(view, off, type, size, result) {
        const records = result.records;
        switch (type) {
            case types_js_1.EMR.EOF:
                records.push({ type: 'eof' });
                return true;
            case types_js_1.EMR.SETMAPMODE:
                records.push({ type: 'setMapMode', mode: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETWINDOWEXTEX:
                records.push({
                    type: 'setWindowExt',
                    cx: view.getInt32(off + 8, true),
                    cy: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.SETWINDOWORGEX:
                records.push({
                    type: 'setWindowOrg',
                    x: view.getInt32(off + 8, true),
                    y: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.SETVIEWPORTEXTEX:
                records.push({
                    type: 'setViewportExt',
                    cx: view.getInt32(off + 8, true),
                    cy: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.SETVIEWPORTORGEX:
                records.push({
                    type: 'setViewportOrg',
                    x: view.getInt32(off + 8, true),
                    y: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.SETWORLDTRANSFORM:
                records.push({ type: 'setWorldTransform', xform: readXform(view, off + 8) });
                break;
            case types_js_1.EMR.MODIFYWORLDTRANSFORM:
                records.push({
                    type: 'modifyWorldTransform',
                    xform: readXform(view, off + 8),
                    mode: view.getUint32(off + 32, true),
                });
                break;
            case types_js_1.EMR.SAVEDC:
                records.push({ type: 'saveDc' });
                break;
            case types_js_1.EMR.RESTOREDC:
                records.push({ type: 'restoreDc', saved: view.getInt32(off + 8, true) });
                break;
            case types_js_1.EMR.SETBKMODE:
                records.push({ type: 'setBkMode', mode: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETBKCOLOR:
                records.push({ type: 'setBkColor', color: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETTEXTCOLOR:
                records.push({ type: 'setTextColor', color: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETTEXTALIGN:
                records.push({ type: 'setTextAlign', align: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETROP2:
                records.push({ type: 'setRop2', mode: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETSTRETCHBLTMODE:
                records.push({ type: 'setStretchBltMode', mode: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.SETPOLYFILLMODE:
                records.push({ type: 'setPolyFillMode', mode: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.INTERSECTCLIPRECT:
                records.push({ type: 'intersectClipRect', rect: readRect(view, off + 8) });
                break;
            case types_js_1.EMR.EXTSELECTCLIPRGN:
                records.push(this.parseExtSelectClipRgn(view, off, size));
                break;
            case types_js_1.EMR.SELECTCLIPPATH:
                // Path-based clipping not supported; skip gracefully
                this.logger.debug('EMR_SELECTCLIPPATH not supported; skipping');
                break;
            case types_js_1.EMR.CREATEPEN:
                records.push({
                    type: 'createPen',
                    index: view.getUint32(off + 8, true),
                    style: view.getUint32(off + 12, true),
                    width: view.getInt32(off + 16, true),
                    color: view.getUint32(off + 24, true),
                });
                break;
            case types_js_1.EMR.EXTCREATEPEN:
                records.push({
                    type: 'createPen',
                    index: view.getUint32(off + 8, true),
                    style: view.getUint32(off + 28, true),
                    width: view.getInt32(off + 32, true),
                    color: view.getUint32(off + 40, true),
                });
                break;
            case types_js_1.EMR.CREATEBRUSHINDIRECT:
                records.push({
                    type: 'createBrush',
                    index: view.getUint32(off + 8, true),
                    style: view.getUint32(off + 12, true),
                    color: view.getUint32(off + 16, true),
                });
                break;
            case types_js_1.EMR.EXTCREATEFONTINDIRECTW:
                records.push({
                    type: 'createFont',
                    index: view.getUint32(off + 8, true),
                    font: readLogFontW(view, off + 12),
                });
                break;
            case types_js_1.EMR.SELECTOBJECT:
                records.push({ type: 'selectObject', index: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.DELETEOBJECT:
                records.push({ type: 'deleteObject', index: view.getUint32(off + 8, true) });
                break;
            case types_js_1.EMR.MOVETOEX:
                records.push({
                    type: 'moveTo',
                    x: view.getInt32(off + 8, true),
                    y: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.LINETO:
                records.push({
                    type: 'lineTo',
                    x: view.getInt32(off + 8, true),
                    y: view.getInt32(off + 12, true),
                });
                break;
            case types_js_1.EMR.RECTANGLE:
                records.push({ type: 'rectangle', rect: readRect(view, off + 8) });
                break;
            case types_js_1.EMR.ELLIPSE:
                records.push({ type: 'ellipse', rect: readRect(view, off + 8) });
                break;
            case types_js_1.EMR.ROUNDRECT:
                records.push({
                    type: 'roundRect',
                    rect: readRect(view, off + 8),
                    cornerWidth: view.getInt32(off + 24, true),
                    cornerHeight: view.getInt32(off + 28, true),
                });
                break;
            case types_js_1.EMR.POLYGON:
                records.push({ type: 'polygon', points: readPoints32(view, off, size) });
                break;
            case types_js_1.EMR.POLYLINE:
                records.push({ type: 'polyline', points: readPoints32(view, off, size) });
                break;
            case types_js_1.EMR.POLYBEZIER:
                records.push({ type: 'polyBezier', points: readPoints32(view, off, size) });
                break;
            case types_js_1.EMR.POLYBEZIERTO:
                records.push({ type: 'polyBezierTo', points: readPoints32(view, off, size) });
                break;
            case types_js_1.EMR.POLYLINETO:
                records.push({ type: 'polylineTo', points: readPoints32(view, off, size) });
                break;
            case types_js_1.EMR.POLYGON16:
                records.push({ type: 'polygon', points: readPoints16(view, off, size) });
                break;
            case types_js_1.EMR.POLYLINE16:
                records.push({ type: 'polyline', points: readPoints16(view, off, size) });
                break;
            case types_js_1.EMR.POLYBEZIER16:
                records.push({ type: 'polyBezier', points: readPoints16(view, off, size) });
                break;
            case types_js_1.EMR.POLYBEZIERTO16:
                records.push({ type: 'polyBezierTo', points: readPoints16(view, off, size) });
                break;
            case types_js_1.EMR.POLYLINETO16:
                records.push({ type: 'polylineTo', points: readPoints16(view, off, size) });
                break;
            case types_js_1.EMR.POLYPOLYGON:
                records.push({ type: 'polyPolygon', polygons: readPolyPoly(view, off, size, 8) });
                break;
            case types_js_1.EMR.POLYPOLYGON16:
                records.push({ type: 'polyPolygon', polygons: readPolyPoly(view, off, size, 4) });
                break;
            case types_js_1.EMR.POLYPOLYLINE:
                records.push({ type: 'polyPolyline', polylines: readPolyPoly(view, off, size, 8) });
                break;
            case types_js_1.EMR.POLYPOLYLINE16:
                records.push({ type: 'polyPolyline', polylines: readPolyPoly(view, off, size, 4) });
                break;
            case types_js_1.EMR.EXTTEXTOUTA:
            case types_js_1.EMR.EXTTEXTOUTW:
                records.push(this.parseExtTextOut(view, off, size, type === types_js_1.EMR.EXTTEXTOUTW));
                break;
            case types_js_1.EMR.BITBLT:
                records.push(this.parseBltRecord(view, off, size, false));
                break;
            case types_js_1.EMR.STRETCHBLT:
                records.push(this.parseBltRecord(view, off, size, true));
                break;
            case types_js_1.EMR.STRETCHDIBITS:
                records.push(this.parseStretchDiBits(view, off, size));
                break;
            case types_js_1.EMR.COMMENT:
                this.parseComment(view, off, size, result);
                break;
            default:
                // Unsupported record: skip via nSize
                this.logger.debug('Skipping unsupported EMF record', { type, size });
                break;
        }
        return false;
    }
    /**
     * Parses EMR_EXTSELECTCLIPRGN. Only RDH_RECTANGLES region data is
     * supported; an empty region with RGN_COPY resets the clip.
     */
    parseExtSelectClipRgn(view, off, size) {
        const rgnDataSize = view.getUint32(off + 8, true);
        const mode = view.getUint32(off + 12, true);
        const rects = [];
        // RegionDataHeader is 32 bytes: Size, Type, CountRects, RgnSize, Bounds
        if (rgnDataSize >= 32 && off + 16 + 32 <= off + size) {
            const rgnOff = off + 16;
            const countRects = view.getUint32(rgnOff + 8, true);
            let rectOff = rgnOff + 32;
            for (let i = 0; i < countRects; i++) {
                if (rectOff + 16 > off + size) {
                    break;
                }
                rects.push(readRect(view, rectOff));
                rectOff += 16;
            }
        }
        return { type: 'selectClipRgn', mode, rects };
    }
    /**
     * Parses EMR_EXTTEXTOUTW / EMR_EXTTEXTOUTA.
     */
    parseExtTextOut(view, off, size, isWide) {
        // EmrText structure starts at offset 36 (after Bounds, iGraphicsMode, ex/eyScale)
        const x = view.getInt32(off + 36, true);
        const y = view.getInt32(off + 40, true);
        const chars = view.getUint32(off + 44, true);
        const offString = view.getUint32(off + 48, true);
        const options = view.getUint32(off + 52, true);
        const rect = readRect(view, off + 56);
        const offDx = view.getUint32(off + 72, true);
        let text = '';
        const charBytes = isWide ? 2 : 1;
        if (chars > 0 && offString >= 76 && offString + chars * charBytes <= size) {
            const codes = [];
            for (let i = 0; i < chars; i++) {
                codes.push(isWide
                    ? view.getUint16(off + offString + i * 2, true)
                    : view.getUint8(off + offString + i));
            }
            // Chunked conversion: spreading a multi-megabyte record into
            // fromCharCode arguments would exceed V8's argument limit.
            const chunks = [];
            for (let i = 0; i < codes.length; i += 8192) {
                chunks.push(String.fromCharCode(...codes.slice(i, i + 8192)));
            }
            text = chunks.join('');
        }
        let dx;
        // ETO_PDY doubles the array (dx,dy pairs); we use only the x advances
        const stride = (options & 0x2000) !== 0 ? 8 : 4;
        if (chars > 0 && offDx >= 76 && offDx + chars * stride <= size) {
            dx = [];
            for (let i = 0; i < chars; i++) {
                dx.push(view.getInt32(off + offDx + i * stride, true));
            }
        }
        const record = {
            type: 'textOut',
            x,
            y,
            text,
            options,
        };
        if (rect.right >= rect.left && rect.bottom >= rect.top) {
            record.rect = rect;
        }
        if (dx) {
            record.dx = dx;
        }
        return record;
    }
    /**
     * Parses EMR_BITBLT / EMR_STRETCHBLT into a dibDraw record.
     * BitBlt without source bits is commonly a PATCOPY brush fill.
     */
    parseBltRecord(view, off, size, isStretch) {
        const destX = view.getInt32(off + 24, true);
        const destY = view.getInt32(off + 28, true);
        const destWidth = view.getInt32(off + 32, true);
        const destHeight = view.getInt32(off + 36, true);
        const rop = view.getUint32(off + 40, true);
        const srcX = view.getInt32(off + 44, true);
        const srcY = view.getInt32(off + 48, true);
        const dib = this.readEmbeddedDib(view, off, size, off + 84, off + 88, off + 92, off + 96);
        let srcWidth = destWidth;
        let srcHeight = destHeight;
        if (isStretch && size >= 108) {
            srcWidth = view.getInt32(off + 100, true);
            srcHeight = view.getInt32(off + 104, true);
        }
        const record = {
            type: 'dibDraw',
            destX,
            destY,
            destWidth,
            destHeight,
            srcX,
            srcY,
            srcWidth,
            srcHeight,
            rop,
        };
        if (dib) {
            record.dib = dib;
        }
        return record;
    }
    /**
     * Parses EMR_STRETCHDIBITS into a dibDraw record.
     */
    parseStretchDiBits(view, off, size) {
        const destX = view.getInt32(off + 24, true);
        const destY = view.getInt32(off + 28, true);
        const srcX = view.getInt32(off + 32, true);
        const srcY = view.getInt32(off + 36, true);
        const srcWidth = view.getInt32(off + 40, true);
        const srcHeight = view.getInt32(off + 44, true);
        const dib = this.readEmbeddedDib(view, off, size, off + 48, off + 52, off + 56, off + 60);
        const rop = view.getUint32(off + 68, true);
        const destWidth = view.getInt32(off + 72, true);
        const destHeight = view.getInt32(off + 76, true);
        const record = {
            type: 'dibDraw',
            destX,
            destY,
            destWidth,
            destHeight,
            srcX,
            srcY,
            srcWidth,
            srcHeight,
            rop,
        };
        if (dib) {
            record.dib = dib;
        }
        return record;
    }
    /**
     * Reads an embedded DIB (BITMAPINFO + bits) referenced by offset/size
     * fields within a record. Returns undefined when no bits are present.
     */
    readEmbeddedDib(view, recordOff, recordSize, offBmiField, cbBmiField, offBitsField, cbBitsField) {
        const offBmi = view.getUint32(offBmiField, true);
        const cbBmi = view.getUint32(cbBmiField, true);
        const offBits = view.getUint32(offBitsField, true);
        const cbBits = view.getUint32(cbBitsField, true);
        if (cbBmi === 0 || cbBits === 0) {
            return undefined;
        }
        if (offBmi + cbBmi > recordSize || offBits + cbBits > recordSize) {
            this.logger.warn('Embedded DIB extends beyond record; skipping');
            return undefined;
        }
        const base = view.byteOffset + recordOff;
        return {
            bmi: new Uint8Array(view.buffer.slice(base + offBmi, base + offBmi + cbBmi)),
            bits: new Uint8Array(view.buffer.slice(base + offBits, base + offBits + cbBits)),
        };
    }
    /**
     * Parses EMR_COMMENT, detecting EMF+ streams. EMF+ payloads are skipped
     * (Office writes EMF+ Dual with a complete GDI fallback).
     */
    parseComment(view, off, size, result) {
        const dataSize = view.getUint32(off + 8, true);
        if (dataSize < 4 || 12 + dataSize > size) {
            return;
        }
        const commentId = view.getUint32(off + 12, true);
        if (commentId !== EMFPLUS_COMMENT_ID) {
            return;
        }
        result.emfPlusDetected = true;
        // First EMF+ record header follows the comment identifier:
        // Type (u16), Flags (u16). EmfPlusHeader Flags bit 0 = Dual mode.
        if (dataSize >= 8) {
            const plusType = view.getUint16(off + 16, true);
            const plusFlags = view.getUint16(off + 18, true);
            if (plusType === EMFPLUS_HEADER_TYPE && (plusFlags & 0x1) !== 0) {
                result.emfPlusDual = true;
            }
        }
    }
}
exports.EmfParser = EmfParser;
/** Reads a RECTL (4 x s32) at the given offset. */
function readRect(view, off) {
    return {
        left: view.getInt32(off, true),
        top: view.getInt32(off + 4, true),
        right: view.getInt32(off + 8, true),
        bottom: view.getInt32(off + 12, true),
    };
}
/** Reads an XFORM (6 x f32) at the given offset. */
function readXform(view, off) {
    return {
        m11: view.getFloat32(off, true),
        m12: view.getFloat32(off + 4, true),
        m21: view.getFloat32(off + 8, true),
        m22: view.getFloat32(off + 12, true),
        dx: view.getFloat32(off + 16, true),
        dy: view.getFloat32(off + 20, true),
    };
}
/** Reads a LogFont (W) structure at the given offset. */
function readLogFontW(view, off) {
    const height = view.getInt32(off, true);
    const weight = view.getInt32(off + 16, true);
    const italic = view.getUint8(off + 20) !== 0;
    const underline = view.getUint8(off + 21) !== 0;
    const strikeOut = view.getUint8(off + 22) !== 0;
    let faceName = '';
    for (let i = 0; i < 32; i++) {
        const code = view.getUint16(off + 28 + i * 2, true);
        if (code === 0) {
            break;
        }
        faceName += String.fromCharCode(code);
    }
    return { height, weight, italic, underline, strikeOut, faceName };
}
/**
 * Reads the point array of a 32-bit poly record
 * (Bounds at +8, Count at +24, POINTL data at +28).
 */
function readPoints32(view, off, size) {
    const declared = view.getUint32(off + 24, true);
    const count = Math.min(declared, Math.floor((size - 28) / 8));
    const points = [];
    for (let i = 0; i < count; i++) {
        points.push({
            x: view.getInt32(off + 28 + i * 8, true),
            y: view.getInt32(off + 32 + i * 8, true),
        });
    }
    return points;
}
/**
 * Reads the point array of a 16-bit poly record
 * (Bounds at +8, Count at +24, POINTS data at +28).
 */
function readPoints16(view, off, size) {
    const declared = view.getUint32(off + 24, true);
    const count = Math.min(declared, Math.floor((size - 28) / 4));
    const points = [];
    for (let i = 0; i < count; i++) {
        points.push({
            x: view.getInt16(off + 28 + i * 4, true),
            y: view.getInt16(off + 30 + i * 4, true),
        });
    }
    return points;
}
/**
 * Reads a PolyPolygon/PolyPolyline record: NumberOfPolys at +24,
 * total Count at +28, per-poly counts, then packed points.
 * @param pointBytes 8 for POINTL data, 4 for POINTS data
 */
function readPolyPoly(view, off, size, pointBytes) {
    const numPolys = view.getUint32(off + 24, true);
    const countsEnd = 32 + numPolys * 4;
    if (countsEnd > size) {
        return [];
    }
    const counts = [];
    for (let i = 0; i < numPolys; i++) {
        counts.push(view.getUint32(off + 32 + i * 4, true));
    }
    const polys = [];
    let pointOff = off + countsEnd;
    const end = off + size;
    for (const count of counts) {
        const poly = [];
        for (let i = 0; i < count; i++) {
            if (pointOff + pointBytes > end) {
                break;
            }
            if (pointBytes === 8) {
                poly.push({
                    x: view.getInt32(pointOff, true),
                    y: view.getInt32(pointOff + 4, true),
                });
            }
            else {
                poly.push({
                    x: view.getInt16(pointOff, true),
                    y: view.getInt16(pointOff + 2, true),
                });
            }
            pointOff += pointBytes;
        }
        polys.push(poly);
    }
    return polys;
}
