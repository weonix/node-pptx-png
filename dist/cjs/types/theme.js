"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_THEME = exports.DEFAULT_FONT_SCHEME = exports.DEFAULT_OFFICE_COLORS = exports.DEFAULT_COLOR_MAP = void 0;
/**
 * Default color map used when a master declares no p:clrMap.
 * Matches the standard Office mapping (and the library's previous
 * hardcoded bg1->lt1 / tx1->dk1 behavior).
 */
exports.DEFAULT_COLOR_MAP = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
};
/**
 * Default Office theme colors.
 */
exports.DEFAULT_OFFICE_COLORS = {
    dark1: { r: 0, g: 0, b: 0, a: 255 },
    light1: { r: 255, g: 255, b: 255, a: 255 },
    dark2: { r: 68, g: 84, b: 106, a: 255 },
    light2: { r: 231, g: 230, b: 230, a: 255 },
    accent1: { r: 68, g: 114, b: 196, a: 255 },
    accent2: { r: 237, g: 125, b: 49, a: 255 },
    accent3: { r: 165, g: 165, b: 165, a: 255 },
    accent4: { r: 255, g: 192, b: 0, a: 255 },
    accent5: { r: 91, g: 155, b: 213, a: 255 },
    accent6: { r: 112, g: 173, b: 71, a: 255 },
    hyperlink: { r: 5, g: 99, b: 193, a: 255 },
    followedHyperlink: { r: 149, g: 79, b: 114, a: 255 },
};
/**
 * Default font scheme.
 */
exports.DEFAULT_FONT_SCHEME = {
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
};
/**
 * Default resolved theme.
 */
exports.DEFAULT_THEME = {
    colors: exports.DEFAULT_OFFICE_COLORS,
    fonts: exports.DEFAULT_FONT_SCHEME,
    backgroundFillStyles: [
        { r: 255, g: 255, b: 255, a: 255 },
        { r: 231, g: 230, b: 230, a: 255 },
        { r: 68, g: 84, b: 106, a: 255 },
    ],
};
