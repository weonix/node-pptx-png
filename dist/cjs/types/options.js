"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_RENDER_OPTIONS = exports.SIZE_PRESET_WIDTHS = void 0;
exports.mergeRenderOptions = mergeRenderOptions;
/**
 * Target widths in pixels for each size preset.
 */
exports.SIZE_PRESET_WIDTHS = {
    thumb: 256,
    preview: 640,
    hd: 1920,
    '4k': 3840,
};
/**
 * Default rendering options.
 */
exports.DEFAULT_RENDER_OPTIONS = {
    width: 1920,
    height: undefined,
    scale: undefined,
    preset: undefined,
    slideNumbers: undefined,
    format: 'png',
    quality: undefined,
    jpegQuality: 90,
    gpu: 'auto',
    textContrast: undefined,
    textGamma: undefined,
    backgroundColor: undefined,
    logLevel: 'warn',
    debugMode: false,
    pngOptimization: 'none',
    fonts: undefined,
};
function mergeRenderOptions(options = {}) {
    return {
        width: options.width ?? exports.DEFAULT_RENDER_OPTIONS.width,
        height: options.height ?? exports.DEFAULT_RENDER_OPTIONS.height,
        scale: options.scale ?? exports.DEFAULT_RENDER_OPTIONS.scale,
        preset: options.preset ?? exports.DEFAULT_RENDER_OPTIONS.preset,
        slideNumbers: options.slideNumbers ?? exports.DEFAULT_RENDER_OPTIONS.slideNumbers,
        format: options.format ?? exports.DEFAULT_RENDER_OPTIONS.format,
        quality: options.quality ?? exports.DEFAULT_RENDER_OPTIONS.quality,
        jpegQuality: options.jpegQuality ?? exports.DEFAULT_RENDER_OPTIONS.jpegQuality,
        gpu: options.gpu ?? exports.DEFAULT_RENDER_OPTIONS.gpu,
        textContrast: options.textContrast ?? exports.DEFAULT_RENDER_OPTIONS.textContrast,
        textGamma: options.textGamma ?? exports.DEFAULT_RENDER_OPTIONS.textGamma,
        backgroundColor: options.backgroundColor ?? exports.DEFAULT_RENDER_OPTIONS.backgroundColor,
        logLevel: options.logLevel ?? exports.DEFAULT_RENDER_OPTIONS.logLevel,
        debugMode: options.debugMode ?? exports.DEFAULT_RENDER_OPTIONS.debugMode,
        pngOptimization: options.pngOptimization ?? exports.DEFAULT_RENDER_OPTIONS.pngOptimization,
        fonts: options.fonts ?? exports.DEFAULT_RENDER_OPTIONS.fonts,
    };
}
