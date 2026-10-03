/**
 * Target widths in pixels for each size preset.
 */
export const SIZE_PRESET_WIDTHS = {
    thumb: 256,
    preview: 640,
    hd: 1920,
    '4k': 3840,
};
/**
 * Default rendering options.
 */
export const DEFAULT_RENDER_OPTIONS = {
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
export function mergeRenderOptions(options = {}) {
    return {
        width: options.width ?? DEFAULT_RENDER_OPTIONS.width,
        height: options.height ?? DEFAULT_RENDER_OPTIONS.height,
        scale: options.scale ?? DEFAULT_RENDER_OPTIONS.scale,
        preset: options.preset ?? DEFAULT_RENDER_OPTIONS.preset,
        slideNumbers: options.slideNumbers ?? DEFAULT_RENDER_OPTIONS.slideNumbers,
        format: options.format ?? DEFAULT_RENDER_OPTIONS.format,
        quality: options.quality ?? DEFAULT_RENDER_OPTIONS.quality,
        jpegQuality: options.jpegQuality ?? DEFAULT_RENDER_OPTIONS.jpegQuality,
        gpu: options.gpu ?? DEFAULT_RENDER_OPTIONS.gpu,
        textContrast: options.textContrast ?? DEFAULT_RENDER_OPTIONS.textContrast,
        textGamma: options.textGamma ?? DEFAULT_RENDER_OPTIONS.textGamma,
        backgroundColor: options.backgroundColor ?? DEFAULT_RENDER_OPTIONS.backgroundColor,
        logLevel: options.logLevel ?? DEFAULT_RENDER_OPTIONS.logLevel,
        debugMode: options.debugMode ?? DEFAULT_RENDER_OPTIONS.debugMode,
        pngOptimization: options.pngOptimization ?? DEFAULT_RENDER_OPTIONS.pngOptimization,
        fonts: options.fonts ?? DEFAULT_RENDER_OPTIONS.fonts,
    };
}
//# sourceMappingURL=options.js.map