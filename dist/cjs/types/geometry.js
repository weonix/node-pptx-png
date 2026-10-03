"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Colors = exports.IDENTITY_TRANSFORM = void 0;
/**
 * Identity transform (no transformation).
 */
exports.IDENTITY_TRANSFORM = {
    a: 1,
    b: 0,
    c: 0,
    d: 1,
    e: 0,
    f: 0,
};
/**
 * Common RGBA colors.
 */
exports.Colors = {
    transparent: { r: 0, g: 0, b: 0, a: 0 },
    black: { r: 0, g: 0, b: 0, a: 255 },
    white: { r: 255, g: 255, b: 255, a: 255 },
    red: { r: 255, g: 0, b: 0, a: 255 },
    green: { r: 0, g: 128, b: 0, a: 255 },
    blue: { r: 0, g: 0, b: 255, a: 255 },
    gray: { r: 128, g: 128, b: 128, a: 255 },
    lightGray: { r: 211, g: 211, b: 211, a: 255 },
    darkGray: { r: 169, g: 169, b: 169, a: 255 },
};
