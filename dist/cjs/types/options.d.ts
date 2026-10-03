/**
 * Output image format options.
 *
 * - 'png': lossless raster (image/png); the only format that honors
 *   `pngOptimization`
 * - 'jpeg': lossy raster (image/jpeg); honors `quality`
 * - 'webp': lossy raster (image/webp); honors `quality`
 * - 'svg': vector output (image/svg+xml); `imageData` contains the UTF-8
 *   SVG document bytes
 */
export type ImageFormat = 'png' | 'jpeg' | 'webp' | 'svg';
/**
 * Logging level for the renderer.
 */
export type LogLevel = 'debug' | 'info' | 'warn' | 'error' | 'silent';
/**
 * PNG optimization preset names.
 *
 * Note: The native skia-canvas encoder is already efficient (similar to zlib level 6).
 * Sharp's main value is palette-based quantization for dramatic size reduction.
 *
 * - 'none': No optimization (native canvas export, fastest)
 * - 'fast': Quick lossless recompression (~1-2% smaller)
 * - 'balanced': Lossless with adaptive filtering (~2-3% smaller)
 * - 'maximum': Same as balanced (skia-canvas is already efficient for lossless)
 * - 'web': Palette quantization (60-70% smaller, may lose quality on photos/gradients)
 */
export type PngOptimizationPreset = 'none' | 'fast' | 'balanced' | 'maximum' | 'web';
/**
 * Named output-size presets.
 *
 * Each preset selects a target width in pixels; height is auto-calculated
 * from the slide aspect ratio.
 *
 * - 'thumb': 256px wide
 * - 'preview': 640px wide
 * - 'hd': 1920px wide
 * - '4k': 3840px wide
 */
export type SizePreset = 'thumb' | 'preview' | 'hd' | '4k';
/**
 * Target widths in pixels for each size preset.
 */
export declare const SIZE_PRESET_WIDTHS: Record<SizePreset, number>;
/**
 * Custom PNG optimization options.
 * Use these for fine-grained control over compression settings.
 */
export interface PngOptimizationOptions {
    /**
     * PNG compression level (0-9).
     * 0 = fastest/largest, 9 = slowest/smallest.
     * @default 6
     */
    compressionLevel?: number;
    /**
     * Use adaptive row filtering for better compression.
     * Slightly slower but can improve compression ratio.
     * @default true
     */
    adaptiveFiltering?: boolean;
    /**
     * Convert to indexed/palette PNG.
     * Significant size reduction for images with limited colors.
     * May cause quality loss for photos or gradients.
     * @default false
     */
    palette?: boolean;
    /**
     * Maximum colors for palette mode (2-256).
     * Only used when palette is true.
     * @default 256
     */
    colors?: number;
    /**
     * Quality threshold for palette quantization (1-100).
     * Lower values = more aggressive compression.
     * Only used when palette is true.
     * @default 90
     */
    quality?: number;
    /**
     * Floyd-Steinberg dithering strength (0.0-1.0).
     * Higher values reduce banding in palette mode.
     * Only used when palette is true.
     * @default 1.0
     */
    dither?: number;
}
/**
 * A user-supplied font file to register before rendering.
 */
export interface FontRegistration {
    /**
     * Family name to register the font under. Text runs referencing this
     * family (directly or via theme fonts) will use the registered file.
     */
    family: string;
    /**
     * Path to a font file (TTF/OTF/TTC/WOFF) or a Buffer with the font bytes.
     * Buffers are written to a session temp directory before registration
     * because skia-canvas registers fonts by file path.
     */
    source: string | Buffer;
    /**
     * Advisory font weight (e.g., 400, 700, 'bold'). skia-canvas derives the
     * actual weight from the font file itself; this field documents intent and
     * is reserved for forward compatibility.
     */
    weight?: number | string;
    /**
     * Advisory font style (e.g., 'normal', 'italic'). skia-canvas derives the
     * actual style from the font file itself; this field documents intent and
     * is reserved for forward compatibility.
     */
    style?: string;
}
/**
 * Font handling options.
 */
export interface FontOptions {
    /**
     * Extract and register fonts embedded in the PPTX (p:embeddedFontLst)
     * before rendering, so text renders with the deck's own fonts.
     * @default true
     */
    useEmbeddedFonts?: boolean;
    /**
     * User-supplied font files to register before rendering.
     */
    register?: FontRegistration[];
    /**
     * Custom fallback chains, merged over the built-in chains. Keys are the
     * requested family names (after theme resolution); values are ordered
     * fallback family lists (e.g., `{ Calibri: ['Carlito', 'Arial'] }`).
     * A custom chain replaces the built-in chain for that family.
     */
    fallbacks?: Record<string, string[]>;
}
/**
 * Options for rendering PPTX presentations to images.
 */
export interface PptxRenderOptions {
    /**
     * Target width in pixels.
     * Height will be auto-calculated based on slide aspect ratio.
     * @default 1920
     */
    width?: number;
    /**
     * Target height in pixels.
     * If omitted, calculated from width and slide aspect ratio.
     */
    height?: number;
    /**
     * Scale factor applied to the slide's native size at 96 DPI
     * (e.g., 1.0 renders a 16:9 slide at 1280x720, 0.5 at 640x360).
     * Only used when neither width nor height is specified.
     * Precedence: width/height > scale > default width.
     */
    scale?: number;
    /**
     * Named output-size preset ('thumb' 256px, 'preview' 640px, 'hd' 1920px,
     * '4k' 3840px wide; height auto-calculated from the slide aspect ratio).
     * Precedence: explicit width/height > scale > preset > default width.
     * Conflicts are resolved leniently: the preset is ignored when width,
     * height, or scale is present.
     */
    preset?: SizePreset;
    /**
     * Slides to render, as 1-based slide numbers (e.g., [1, 2, 3]).
     * Only used by renderPresentation; if omitted, all slides are rendered.
     * Invalid slide numbers produce per-slide failure results rather than throwing.
     */
    slideNumbers?: number[];
    /**
     * Output image format.
     * @default 'png'
     */
    format?: ImageFormat;
    /**
     * Encoding quality (0.0-1.0) for lossy formats. Applies when format is
     * 'jpeg' or 'webp'. Takes precedence over the legacy `jpegQuality`.
     * @default 0.9 (via jpegQuality's default of 90)
     */
    quality?: number;
    /**
     * JPEG quality (1-100). Only applicable when format is 'jpeg'.
     * @deprecated Use `quality` (0.0-1.0) instead, which applies to both
     *   'jpeg' and 'webp'. When both are set, `quality` wins.
     * @default 90
     */
    jpegQuality?: number;
    /**
     * GPU rendering toggle for the underlying skia-canvas surface.
     *
     * - true: request GPU rendering (falls back to CPU if unavailable)
     * - false: force CPU rendering
     * - 'auto': leave the skia-canvas default untouched
     *
     * The mode actually used is recorded in the debug log.
     * @default 'auto'
     */
    gpu?: boolean | 'auto';
    /**
     * Additional text stroke contrast (skia-canvas textContrast, 0-1).
     * PowerPoint's rasterizer draws text with heavier stems than skia's
     * default; textContrast: 1 with textGamma: 1 closely matches PowerPoint
     * exports. @default 0 (skia default)
     */
    textContrast?: number;
    /**
     * Gamma for blending letterform edges (skia-canvas textGamma).
     * @default 1.4 (skia default)
     */
    textGamma?: number;
    /**
     * Override slide background color (hex string, e.g., '#FFFFFF').
     * If set, this replaces the slide's background.
     */
    backgroundColor?: string;
    /**
     * Logging level for diagnostic output.
     * @default 'warn'
     */
    logLevel?: LogLevel;
    /**
     * Enable debug mode to draw bounding boxes and element IDs.
     * @default false
     */
    debugMode?: boolean;
    /**
     * PNG optimization settings.
     * Can be a preset name or custom options object.
     * Requires Sharp to be installed for optimization.
     * If Sharp is not available, falls back to native canvas export.
     * @default 'none'
     */
    pngOptimization?: PngOptimizationPreset | PngOptimizationOptions;
    /**
     * Font handling: embedded font extraction, user font registration, and
     * custom fallback chains.
     */
    fonts?: FontOptions;
}
/**
 * Default rendering options.
 */
export declare const DEFAULT_RENDER_OPTIONS: Required<Omit<PptxRenderOptions, 'height' | 'scale' | 'preset' | 'slideNumbers' | 'backgroundColor' | 'textContrast' | 'textGamma' | 'pngOptimization' | 'fonts' | 'quality'>> & {
    height: undefined;
    scale: undefined;
    preset: undefined;
    slideNumbers: undefined;
    backgroundColor: undefined;
    textContrast: undefined;
    textGamma: undefined;
    pngOptimization: PngOptimizationPreset;
    fonts: undefined;
    quality: undefined;
};
/**
 * Resolved render options after merging with defaults.
 * Unlike Required<PptxRenderOptions>, this type accurately represents
 * that height and backgroundColor remain optional (undefined) even after resolution.
 */
export interface ResolvedRenderOptions {
    /** Target width in pixels. */
    width: number;
    /** Target height in pixels. Undefined means auto-calculate from aspect ratio. */
    height: number | undefined;
    /** Scale factor on the slide's native 96-DPI size. Undefined means not requested. */
    scale: number | undefined;
    /** Named output-size preset. Undefined means not requested. */
    preset: SizePreset | undefined;
    /** 1-based slide numbers to render. Undefined means all slides. */
    slideNumbers: number[] | undefined;
    /** Output image format. */
    format: ImageFormat;
    /** Encoding quality (0.0-1.0) for jpeg/webp. Undefined defers to jpegQuality. */
    quality: number | undefined;
    /** JPEG quality (1-100). Legacy alias; `quality` wins when both are set. */
    jpegQuality: number;
    /** GPU rendering toggle. 'auto' leaves the skia-canvas default untouched. */
    gpu: boolean | 'auto';
    /** Text stroke contrast (skia textContrast). Undefined = skia default. */
    textContrast: number | undefined;
    /** Letterform edge gamma (skia textGamma). Undefined = skia default. */
    textGamma: number | undefined;
    /** Override background color. Undefined means use slide's background. */
    backgroundColor: string | undefined;
    /** Logging level. */
    logLevel: LogLevel;
    /** Debug mode for bounding boxes and element IDs. */
    debugMode: boolean;
    /** PNG optimization settings. */
    pngOptimization: PngOptimizationPreset | PngOptimizationOptions;
    /** Font handling options. Undefined means all-default font behavior. */
    fonts: FontOptions | undefined;
}
/**
 * Merges user-provided render options with the defaults.
 *
 * Unlike a plain object spread, keys whose value is explicitly `undefined`
 * do not override the corresponding default (e.g. `{ width: undefined }`
 * behaves identically to omitting `width`), preventing silently broken
 * output such as NaN canvas dimensions.
 */
/**
 * Options for openPresentation().
 */
export interface OpenOptions {
    /**
     * Logging level for diagnostic output.
     * @default 'warn'
     */
    logLevel?: LogLevel;
}
/**
 * Progress event emitted by PptxDocument.slides() after each slide.
 */
export interface SlideProgressEvent {
    /** Number of slides emitted so far, including the current one. */
    done: number;
    /** Total number of slides selected for this iteration. */
    total: number;
    /** 1-based number of the slide that was just rendered. */
    slideNumber: number;
}
/**
 * Slide selection for PptxDocument.slides(): an explicit list of 1-based
 * slide numbers (emitted in the given order), or an inclusive range.
 */
export type SlideSelection = number[] | {
    from?: number;
    to?: number;
};
/**
 * Options for PptxDocument.slides().
 */
export interface SlidesOptions extends PptxRenderOptions {
    /**
     * Slides to emit. List form: 1-based slide numbers, emitted in the given
     * order; invalid numbers yield per-slide failure results (matching the
     * slideNumbers render option). Range form: inclusive `from`/`to` bounds,
     * clamped to the deck (`from` defaults to 1, `to` to the slide count).
     * If omitted, all slides are emitted in deck order.
     */
    slides?: SlideSelection;
    /**
     * Abort signal checked before each slide. When aborted, iteration throws
     * the signal's abort reason (a DOMException named 'AbortError' by default).
     */
    signal?: AbortSignal;
    /**
     * Called after each slide finishes rendering (whether it succeeded or
     * failed), immediately before the result is yielded.
     */
    onProgress?: (event: SlideProgressEvent) => void;
}
/**
 * Options for PptxDocument.exportPdf().
 *
 * Size-related render options (width/height/scale/preset) set the PDF page
 * size, and backgroundColor/fonts/debugMode affect page content. Raster
 * output options (format, quality, jpegQuality, pngOptimization) are
 * ignored: the PDF is exported as vector content.
 */
export interface ExportPdfOptions extends PptxRenderOptions {
    /**
     * Slides to include, as an explicit list of 1-based slide numbers
     * (pages appear in the given order) or an inclusive `from`/`to` range
     * clamped to the deck. If omitted, `slideNumbers` is honored as a list;
     * otherwise all slides are included in deck order.
     */
    slides?: SlideSelection;
}
export declare function mergeRenderOptions(options?: PptxRenderOptions): ResolvedRenderOptions;
