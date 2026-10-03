/**
 * Shared types and GDI constants for the EMF/WMF metafile pipeline.
 *
 * Both EmfParser (MS-EMF binaries) and WmfParser (MS-WMF binaries) produce
 * the same device-independent record model (ParsedMetafile), which
 * EmfRenderer replays onto a skia-canvas Canvas.
 *
 * Ground truth: [MS-EMF] and [MS-WMF] specifications.
 */
/** A 2D point in logical (metafile) units. */
export interface MetafilePoint {
    x: number;
    y: number;
}
/** A rectangle in logical or device units (GDI RECTL semantics). */
export interface MetafileRect {
    left: number;
    top: number;
    right: number;
    bottom: number;
}
/** GDI XFORM world-transform matrix (row-major 2x2 plus translation). */
export interface MetafileXform {
    m11: number;
    m12: number;
    m21: number;
    m22: number;
    dx: number;
    dy: number;
}
/**
 * Font attributes from a GDI LOGFONT.
 * Height follows LOGFONT semantics: negative = em height in logical units,
 * positive = cell height (em + internal leading).
 */
export interface MetafileLogFont {
    height: number;
    weight: number;
    italic: boolean;
    underline: boolean;
    strikeOut: boolean;
    faceName: string;
}
/** An embedded device-independent bitmap (BITMAPINFO + pixel data). */
export interface MetafileDib {
    /** BITMAPINFO bytes (BITMAPINFOHEADER or later, plus palette/masks) */
    bmi: Uint8Array;
    /** Raw pixel data */
    bits: Uint8Array;
}
/**
 * Device-independent record model shared by the EMF and WMF parsers.
 * Colors are raw GDI COLORREF values (0x00BBGGRR).
 */
export type MetafileRecord = {
    type: 'setMapMode';
    mode: number;
} | {
    type: 'setWindowExt';
    cx: number;
    cy: number;
} | {
    type: 'setWindowOrg';
    x: number;
    y: number;
} | {
    type: 'setViewportExt';
    cx: number;
    cy: number;
} | {
    type: 'setViewportOrg';
    x: number;
    y: number;
} | {
    type: 'setWorldTransform';
    xform: MetafileXform;
} | {
    type: 'modifyWorldTransform';
    mode: number;
    xform: MetafileXform;
} | {
    type: 'saveDc';
} | {
    type: 'restoreDc';
    saved: number;
} | {
    type: 'setBkMode';
    mode: number;
} | {
    type: 'setBkColor';
    color: number;
} | {
    type: 'setTextColor';
    color: number;
} | {
    type: 'setTextAlign';
    align: number;
} | {
    type: 'setRop2';
    mode: number;
} | {
    type: 'setStretchBltMode';
    mode: number;
} | {
    type: 'setPolyFillMode';
    mode: number;
} | {
    type: 'intersectClipRect';
    rect: MetafileRect;
} | {
    type: 'selectClipRgn';
    mode: number;
    rects: MetafileRect[];
} | {
    type: 'createPen';
    index: number;
    style: number;
    width: number;
    color: number;
} | {
    type: 'createBrush';
    index: number;
    style: number;
    color: number;
} | {
    type: 'createFont';
    index: number;
    font: MetafileLogFont;
} | {
    type: 'selectObject';
    index: number;
} | {
    type: 'deleteObject';
    index: number;
} | {
    type: 'moveTo';
    x: number;
    y: number;
} | {
    type: 'lineTo';
    x: number;
    y: number;
} | {
    type: 'rectangle';
    rect: MetafileRect;
} | {
    type: 'ellipse';
    rect: MetafileRect;
} | {
    type: 'roundRect';
    rect: MetafileRect;
    cornerWidth: number;
    cornerHeight: number;
} | {
    type: 'polygon';
    points: MetafilePoint[];
} | {
    type: 'polyline';
    points: MetafilePoint[];
} | {
    type: 'polylineTo';
    points: MetafilePoint[];
} | {
    type: 'polyPolygon';
    polygons: MetafilePoint[][];
} | {
    type: 'polyPolyline';
    polylines: MetafilePoint[][];
} | {
    type: 'polyBezier';
    points: MetafilePoint[];
} | {
    type: 'polyBezierTo';
    points: MetafilePoint[];
} | {
    type: 'textOut';
    x: number;
    y: number;
    text: string;
    options: number;
    rect?: MetafileRect;
    /** Per-character advance widths in logical units (ExtTextOut dx array) */
    dx?: number[];
} | {
    type: 'dibDraw';
    destX: number;
    destY: number;
    destWidth: number;
    destHeight: number;
    srcX: number;
    srcY: number;
    srcWidth: number;
    srcHeight: number;
    /** Ternary raster operation (e.g. SRCCOPY, PATCOPY) */
    rop: number;
    dib?: MetafileDib;
} | {
    type: 'eof';
};
/** Metafile header information normalized across EMF and WMF. */
export interface MetafileHeader {
    /** Bounds of the drawing in device units (inclusive-inclusive) */
    bounds: MetafileRect;
    /** Picture frame in 0.01 millimeter units */
    frame: MetafileRect;
    /** Reference device size in pixels */
    devicePx: {
        cx: number;
        cy: number;
    };
    /** Reference device size in millimeters */
    deviceMm: {
        cx: number;
        cy: number;
    };
}
/** A fully parsed metafile ready for replay by EmfRenderer. */
export interface ParsedMetafile {
    header: MetafileHeader;
    records: MetafileRecord[];
    /** Which binary format the records came from */
    source: 'emf' | 'wmf';
    /** True when EMF+ comment records were present in the stream */
    emfPlusDetected: boolean;
    /** True when the EMF+ header declared Dual mode (full GDI fallback present) */
    emfPlusDual: boolean;
    /** True when parsing stopped early due to a malformed/truncated record */
    truncated: boolean;
}
/** EMF record type numbers ([MS-EMF] 2.1.1). */
export declare const EMR: {
    readonly HEADER: 1;
    readonly POLYBEZIER: 2;
    readonly POLYGON: 3;
    readonly POLYLINE: 4;
    readonly POLYBEZIERTO: 5;
    readonly POLYLINETO: 6;
    readonly POLYPOLYLINE: 7;
    readonly POLYPOLYGON: 8;
    readonly SETWINDOWEXTEX: 9;
    readonly SETWINDOWORGEX: 10;
    readonly SETVIEWPORTEXTEX: 11;
    readonly SETVIEWPORTORGEX: 12;
    readonly EOF: 14;
    readonly SETMAPMODE: 17;
    readonly SETBKMODE: 18;
    readonly SETPOLYFILLMODE: 19;
    readonly SETROP2: 20;
    readonly SETSTRETCHBLTMODE: 21;
    readonly SETTEXTALIGN: 22;
    readonly SETTEXTCOLOR: 24;
    readonly SETBKCOLOR: 25;
    readonly MOVETOEX: 27;
    readonly INTERSECTCLIPRECT: 30;
    readonly SAVEDC: 33;
    readonly RESTOREDC: 34;
    readonly SETWORLDTRANSFORM: 35;
    readonly MODIFYWORLDTRANSFORM: 36;
    readonly SELECTOBJECT: 37;
    readonly CREATEPEN: 38;
    readonly CREATEBRUSHINDIRECT: 39;
    readonly DELETEOBJECT: 40;
    readonly ELLIPSE: 42;
    readonly RECTANGLE: 43;
    readonly ROUNDRECT: 44;
    readonly LINETO: 54;
    readonly SELECTCLIPPATH: 67;
    readonly COMMENT: 70;
    readonly EXTSELECTCLIPRGN: 75;
    readonly BITBLT: 76;
    readonly STRETCHBLT: 77;
    readonly STRETCHDIBITS: 81;
    readonly EXTCREATEFONTINDIRECTW: 82;
    readonly EXTTEXTOUTA: 83;
    readonly EXTTEXTOUTW: 84;
    readonly POLYBEZIER16: 85;
    readonly POLYGON16: 86;
    readonly POLYLINE16: 87;
    readonly POLYBEZIERTO16: 88;
    readonly POLYLINETO16: 89;
    readonly POLYPOLYLINE16: 90;
    readonly POLYPOLYGON16: 91;
    readonly EXTCREATEPEN: 95;
};
/** GDI mapping modes. */
export declare const MAP_MODE: {
    readonly TEXT: 1;
    readonly LOMETRIC: 2;
    readonly HIMETRIC: 3;
    readonly LOENGLISH: 4;
    readonly HIENGLISH: 5;
    readonly TWIPS: 6;
    readonly ISOTROPIC: 7;
    readonly ANISOTROPIC: 8;
};
/** GDI background mix modes. */
export declare const BK_MODE: {
    readonly TRANSPARENT: 1;
    readonly OPAQUE: 2;
};
/** GDI pen styles (low nibble of PenStyle). */
export declare const PEN_STYLE: {
    readonly SOLID: 0;
    readonly DASH: 1;
    readonly DOT: 2;
    readonly DASHDOT: 3;
    readonly DASHDOTDOT: 4;
    readonly NULL: 5;
    readonly INSIDEFRAME: 6;
    /** Mask for the style bits within an ExtCreatePen PenStyle value */
    readonly STYLE_MASK: 15;
};
/** GDI brush styles. */
export declare const BRUSH_STYLE: {
    readonly SOLID: 0;
    readonly NULL: 1;
    readonly HATCHED: 2;
    readonly PATTERN: 3;
};
/** GDI polygon fill modes. */
export declare const POLY_FILL_MODE: {
    readonly ALTERNATE: 1;
    readonly WINDING: 2;
};
/** TextAlignmentMode flags for SetTextAlign. */
export declare const TEXT_ALIGN: {
    readonly UPDATECP: 1;
    readonly RIGHT: 2;
    readonly CENTER: 6;
    readonly HORIZONTAL_MASK: 6;
    readonly BOTTOM: 8;
    readonly BASELINE: 24;
    readonly VERTICAL_MASK: 24;
};
/** ExtTextOut option flags. */
export declare const ETO: {
    readonly OPAQUE: 2;
    readonly CLIPPED: 4;
    readonly PDY: 8192;
};
/** Ternary raster operations used by Office metafiles. */
export declare const ROP: {
    readonly SRCCOPY: 13369376;
    readonly PATCOPY: 15728673;
    readonly BLACKNESS: 66;
    readonly WHITENESS: 16711778;
};
/** Region combine modes for ExtSelectClipRgn. */
export declare const RGN_MODE: {
    readonly AND: 1;
    readonly OR: 2;
    readonly XOR: 3;
    readonly DIFF: 4;
    readonly COPY: 5;
};
/** ModifyWorldTransform modes. */
export declare const MWT: {
    readonly IDENTITY: 1;
    readonly LEFTMULTIPLY: 2;
    readonly RIGHTMULTIPLY: 3;
};
/** Flag bit marking a SelectObject index as a stock object. */
export declare const STOCK_OBJECT_FLAG = 2147483648;
/** Stock object identifiers (low bits of a flagged SelectObject index). */
export declare const STOCK_OBJECT: {
    readonly WHITE_BRUSH: 0;
    readonly LTGRAY_BRUSH: 1;
    readonly GRAY_BRUSH: 2;
    readonly DKGRAY_BRUSH: 3;
    readonly BLACK_BRUSH: 4;
    readonly NULL_BRUSH: 5;
    readonly WHITE_PEN: 6;
    readonly BLACK_PEN: 7;
    readonly NULL_PEN: 8;
    readonly OEM_FIXED_FONT: 10;
    readonly ANSI_FIXED_FONT: 11;
    readonly ANSI_VAR_FONT: 12;
    readonly SYSTEM_FONT: 13;
    readonly DEVICE_DEFAULT_FONT: 14;
    readonly SYSTEM_FIXED_FONT: 16;
    readonly DEFAULT_GUI_FONT: 17;
    readonly DC_BRUSH: 18;
    readonly DC_PEN: 19;
};
/**
 * Converts a GDI COLORREF (0x00BBGGRR) to a CSS rgb() string.
 */
export declare function colorRefToCss(colorRef: number): string;
//# sourceMappingURL=types.d.ts.map