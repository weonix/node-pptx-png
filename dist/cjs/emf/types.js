"use strict";
/**
 * Shared types and GDI constants for the EMF/WMF metafile pipeline.
 *
 * Both EmfParser (MS-EMF binaries) and WmfParser (MS-WMF binaries) produce
 * the same device-independent record model (ParsedMetafile), which
 * EmfRenderer replays onto a skia-canvas Canvas.
 *
 * Ground truth: [MS-EMF] and [MS-WMF] specifications.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.STOCK_OBJECT = exports.STOCK_OBJECT_FLAG = exports.MWT = exports.RGN_MODE = exports.ROP = exports.ETO = exports.TEXT_ALIGN = exports.POLY_FILL_MODE = exports.BRUSH_STYLE = exports.PEN_STYLE = exports.BK_MODE = exports.MAP_MODE = exports.EMR = void 0;
exports.colorRefToCss = colorRefToCss;
/** EMF record type numbers ([MS-EMF] 2.1.1). */
exports.EMR = {
    HEADER: 1,
    POLYBEZIER: 2,
    POLYGON: 3,
    POLYLINE: 4,
    POLYBEZIERTO: 5,
    POLYLINETO: 6,
    POLYPOLYLINE: 7,
    POLYPOLYGON: 8,
    SETWINDOWEXTEX: 9,
    SETWINDOWORGEX: 10,
    SETVIEWPORTEXTEX: 11,
    SETVIEWPORTORGEX: 12,
    EOF: 14,
    SETMAPMODE: 17,
    SETBKMODE: 18,
    SETPOLYFILLMODE: 19,
    SETROP2: 20,
    SETSTRETCHBLTMODE: 21,
    SETTEXTALIGN: 22,
    SETTEXTCOLOR: 24,
    SETBKCOLOR: 25,
    MOVETOEX: 27,
    INTERSECTCLIPRECT: 30,
    SAVEDC: 33,
    RESTOREDC: 34,
    SETWORLDTRANSFORM: 35,
    MODIFYWORLDTRANSFORM: 36,
    SELECTOBJECT: 37,
    CREATEPEN: 38,
    CREATEBRUSHINDIRECT: 39,
    DELETEOBJECT: 40,
    ELLIPSE: 42,
    RECTANGLE: 43,
    ROUNDRECT: 44,
    LINETO: 54,
    SELECTCLIPPATH: 67,
    COMMENT: 70,
    EXTSELECTCLIPRGN: 75,
    BITBLT: 76,
    STRETCHBLT: 77,
    STRETCHDIBITS: 81,
    EXTCREATEFONTINDIRECTW: 82,
    EXTTEXTOUTA: 83,
    EXTTEXTOUTW: 84,
    POLYBEZIER16: 85,
    POLYGON16: 86,
    POLYLINE16: 87,
    POLYBEZIERTO16: 88,
    POLYLINETO16: 89,
    POLYPOLYLINE16: 90,
    POLYPOLYGON16: 91,
    EXTCREATEPEN: 95,
};
/** GDI mapping modes. */
exports.MAP_MODE = {
    TEXT: 1,
    LOMETRIC: 2,
    HIMETRIC: 3,
    LOENGLISH: 4,
    HIENGLISH: 5,
    TWIPS: 6,
    ISOTROPIC: 7,
    ANISOTROPIC: 8,
};
/** GDI background mix modes. */
exports.BK_MODE = {
    TRANSPARENT: 1,
    OPAQUE: 2,
};
/** GDI pen styles (low nibble of PenStyle). */
exports.PEN_STYLE = {
    SOLID: 0,
    DASH: 1,
    DOT: 2,
    DASHDOT: 3,
    DASHDOTDOT: 4,
    NULL: 5,
    INSIDEFRAME: 6,
    /** Mask for the style bits within an ExtCreatePen PenStyle value */
    STYLE_MASK: 0x0f,
};
/** GDI brush styles. */
exports.BRUSH_STYLE = {
    SOLID: 0,
    NULL: 1,
    HATCHED: 2,
    PATTERN: 3,
};
/** GDI polygon fill modes. */
exports.POLY_FILL_MODE = {
    ALTERNATE: 1,
    WINDING: 2,
};
/** TextAlignmentMode flags for SetTextAlign. */
exports.TEXT_ALIGN = {
    UPDATECP: 0x01,
    RIGHT: 0x02,
    CENTER: 0x06,
    HORIZONTAL_MASK: 0x06,
    BOTTOM: 0x08,
    BASELINE: 0x18,
    VERTICAL_MASK: 0x18,
};
/** ExtTextOut option flags. */
exports.ETO = {
    OPAQUE: 0x02,
    CLIPPED: 0x04,
    PDY: 0x2000,
};
/** Ternary raster operations used by Office metafiles. */
exports.ROP = {
    SRCCOPY: 0x00cc0020,
    PATCOPY: 0x00f00021,
    BLACKNESS: 0x00000042,
    WHITENESS: 0x00ff0062,
};
/** Region combine modes for ExtSelectClipRgn. */
exports.RGN_MODE = {
    AND: 1,
    OR: 2,
    XOR: 3,
    DIFF: 4,
    COPY: 5,
};
/** ModifyWorldTransform modes. */
exports.MWT = {
    IDENTITY: 1,
    LEFTMULTIPLY: 2,
    RIGHTMULTIPLY: 3,
};
/** Flag bit marking a SelectObject index as a stock object. */
exports.STOCK_OBJECT_FLAG = 0x80000000;
/** Stock object identifiers (low bits of a flagged SelectObject index). */
exports.STOCK_OBJECT = {
    WHITE_BRUSH: 0,
    LTGRAY_BRUSH: 1,
    GRAY_BRUSH: 2,
    DKGRAY_BRUSH: 3,
    BLACK_BRUSH: 4,
    NULL_BRUSH: 5,
    WHITE_PEN: 6,
    BLACK_PEN: 7,
    NULL_PEN: 8,
    OEM_FIXED_FONT: 10,
    ANSI_FIXED_FONT: 11,
    ANSI_VAR_FONT: 12,
    SYSTEM_FONT: 13,
    DEVICE_DEFAULT_FONT: 14,
    SYSTEM_FIXED_FONT: 16,
    DEFAULT_GUI_FONT: 17,
    DC_BRUSH: 18,
    DC_PEN: 19,
};
/**
 * Converts a GDI COLORREF (0x00BBGGRR) to a CSS rgb() string.
 */
function colorRefToCss(colorRef) {
    const r = colorRef & 0xff;
    const g = (colorRef >> 8) & 0xff;
    const b = (colorRef >> 16) & 0xff;
    return `rgb(${r},${g},${b})`;
}
