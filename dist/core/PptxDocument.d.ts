/// <reference lib="esnext.disposable" preserve="true" />
import type { ExportPdfOptions, OpenOptions, PptxRenderOptions, SlidesOptions, SlideRenderResult, PresentationRenderResult } from '../types/index.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * An opened PPTX presentation: parse once, render many.
 *
 * Obtain an instance via {@link openPresentation} (or the static
 * {@link PptxDocument.open}). The underlying archive is opened and the
 * presentation part parsed exactly once; the resolved theme, parsed XML
 * parts, relationships, and table styles are cached on the document and
 * reused across every subsequent render call.
 *
 * Slide numbers on this API are **1-based**, consistent with the
 * `slideNumbers` render option (note that the legacy
 * `PptxImageRenderer.renderSlide` takes a 0-based index instead).
 *
 * Call {@link close} when done, or use `await using` / `using` — the
 * document implements both `Symbol.asyncDispose` and `Symbol.dispose`.
 *
 * @example
 * ```typescript
 * await using doc = await openPresentation('./deck.pptx');
 * for await (const slide of doc.slides({ preset: 'thumb' })) {
 *   await fs.writeFile(`slide-${slide.slideNumber}.png`, slide.imageData);
 * }
 * ```
 */
export declare class PptxDocument {
    /** Number of slides in the presentation. */
    readonly slideCount: number;
    /** Native slide size in EMU (914400 EMU = 1 inch). */
    readonly size: {
        readonly widthEmu: number;
        readonly heightEmu: number;
    };
    private readonly parser;
    private readonly logger;
    /** Resolved presentation-default theme, cached after the first render
     *  (promise-cached so concurrent first calls resolve it only once). */
    private themePromise;
    /** Per-master resolved themes (promise-cached, keyed by master part
     *  path). Multi-master decks reference different themes; each slide's
     *  colors resolve against the theme of its own master chain. */
    private readonly themeByMasterPath;
    /** Embedded-font extraction + registration, performed at most once per
     *  document (registration is global to the process font library). */
    private embeddedFontsPromise;
    /** Presentation-level warnings (font preparation and other once-per-open
     *  events); slideNumber stays undefined on these. Surfaced through the
     *  renderAll/renderPresentation aggregate. */
    private readonly docWarnings;
    private closed;
    private constructor();
    /**
     * Opens a PPTX presentation from a file path or Buffer and parses its
     * presentation part. Rejects if the input is not a readable PPTX.
     *
     * Most callers should use the {@link openPresentation} convenience
     * function instead.
     *
     * @param input File path or Buffer containing PPTX data
     * @param options Open options (log level)
     * @param logger Advanced: existing logger to use instead of creating one
     *   from `options.logLevel` (used internally by PptxImageRenderer)
     */
    static open(input: Buffer | string, options?: OpenOptions, logger?: ILogger): Promise<PptxDocument>;
    /**
     * Renders a single slide by its 1-based slide number.
     *
     * Bad per-call input (out-of-range slide number, invalid scale or preset)
     * resolves to a result with `success: false` rather than rejecting.
     *
     * @param slideNumber 1-based slide number (consistent with the
     *   `slideNumbers` render option; the legacy renderSlide index is 0-based)
     * @param options Optional rendering options
     */
    slide(slideNumber: number, options?: PptxRenderOptions): Promise<SlideRenderResult>;
    /**
     * Streams rendered slides as an async generator, in order.
     *
     * Selection: `options.slides` chooses which slides to emit (list form
     * preserves the given order; range form emits ascending). When omitted,
     * `options.slideNumbers` is honored as a list for consistency with
     * renderAll; otherwise all slides are emitted in deck order.
     *
     * `options.signal` is checked before each slide; on abort, iteration
     * throws the signal's abort reason (a DOMException named 'AbortError' by
     * default). `options.onProgress` fires after each slide renders,
     * immediately before its result is yielded.
     *
     * The parse, theme resolution, font preparation, and renderer setup are
     * all shared across the iteration — no per-slide re-work.
     *
     * Each yielded result carries its own structured `warnings` (omitted when
     * the slide rendered cleanly). Presentation-level warnings (e.g. embedded
     * fonts) are not attributed to any slide; use {@link renderAll} or
     * `renderPresentation` to obtain the aggregate that includes them.
     */
    slides(options?: SlidesOptions): AsyncGenerator<SlideRenderResult, void, undefined>;
    /**
     * Renders all slides (or those selected via `options.slideNumbers`) and
     * buffers the results, matching `renderPresentation` semantics exactly:
     * invalid slide numbers produce per-slide failure results, and render-time
     * errors resolve to a presentation-level error result instead of rejecting.
     *
     * When fidelity warnings occur, the result's `warnings` aggregates
     * presentation-level warnings (slideNumber undefined, e.g. embedded-font
     * registration) and all per-slide warnings, deduplicated by
     * code + message + slideNumber; each slide result also carries its own
     * `warnings`. A clean render omits the field entirely.
     */
    renderAll(options?: PptxRenderOptions): Promise<PresentationRenderResult>;
    /**
     * Exports the selected slides as a single multi-page vector PDF.
     *
     * Each slide is drawn onto its own page of one shared skia-canvas
     * surface (`canvas.newPage()`), so text and shapes stay vector content
     * rather than being rasterized. Page size follows the size options
     * (explicit width/height > scale > preset > default width); raster
     * output options (format, quality, pngOptimization) are ignored.
     *
     * Unlike the raster render methods, this rejects on failure: an invalid
     * slide selection, bad options, or a slide that fails to draw all reject
     * the returned promise (a partial PDF is never produced). Fidelity
     * warnings are logged but not returned; render the slides individually
     * to collect structured warnings.
     *
     * @param options Slide selection plus standard render options
     * @returns PDF bytes for the selected slides, one page per slide
     */
    exportPdf(options?: ExportPdfOptions): Promise<Buffer>;
    /**
     * Closes the document and releases the underlying archive and caches.
     * Idempotent: calling close() again is a no-op. Any render call after
     * close() throws a clear error.
     */
    close(): void;
    /**
     * `await using` support: disposing the document closes it.
     */
    [Symbol.asyncDispose](): Promise<void>;
    /**
     * `using` support: disposing the document closes it.
     */
    [Symbol.dispose](): void;
    /**
     * Throws if the document has been closed.
     */
    private ensureNotClosed;
    /**
     * Whether a 1-based slide number refers to a slide in this deck.
     */
    private isValidSlideNumber;
    /**
     * Builds a failure result for an invalid slide number or bad per-call
     * options, mirroring the failure-result shape of successful render paths.
     */
    private invalidSlideResult;
    /**
     * Expands a slides() selection into an ordered list of 1-based slide
     * numbers. List form is returned as-is (order preserved); range form is
     * clamped to the deck and expanded ascending; undefined means all slides.
     */
    private resolveSlideSelection;
    /**
     * Creates the per-call render session: merges options, applies the size
     * precedence (explicit width/height > scale > preset > default), prepares
     * fonts, resolves the theme (cached on the document), and builds the
     * SlideRenderer that every slide in the call will share.
     *
     * Throws on invalid scale or preset values.
     */
    private createRenderSession;
    /**
     * Applies the scale and preset options to produce concrete pixel
     * dimensions from the slide's native size at 96 DPI.
     * Precedence: explicit width/height > scale > preset > default width.
     * The preset is ignored (lenient conflict handling) when width, height,
     * or scale is present.
     */
    private resolveTargetSize;
    /**
     * Prepares fonts before rendering: sets custom fallback chains, registers
     * user-supplied fonts, and extracts + registers fonts embedded in the
     * PPTX (unless fonts.useEmbeddedFonts is false). Embedded-font extraction
     * runs at most once per document; registration is process-global, so
     * later calls reuse it.
     *
     * Font problems never fail a render; they degrade to the fallback chains.
     */
    private prepareFonts;
    /**
     * Extracts and registers fonts embedded in the PPTX. Never rejects; font
     * problems are logged as warnings so renders degrade gracefully.
     */
    private registerEmbeddedFonts;
    /**
     * Resolves the theme once per document and caches it.
     */
    private resolveThemeOnce;
    /**
     * Resolves the theme for a slide from its master chain. Each master has
     * its own theme relationship, and multi-master decks can carry themes
     * with entirely different color schemes — a slide's scheme colors
     * (accent1, tx2, ...) must resolve against ITS master's theme, not a
     * single presentation-wide one. Falls back to the presentation-default
     * theme when the master is unknown or its theme cannot be loaded.
     * Results are promise-cached per master path.
     */
    private resolveThemeForSlide;
    /**
     * Returns the session renderer for a resolved theme, creating and caching
     * one per distinct theme. The default theme's renderer is seeded at
     * session creation, so single-theme decks always reuse it.
     */
    private rendererForTheme;
    /**
     * Renders one slide (0-based index) with an established session, loading
     * the slide's layout/master inheritance chain. Never rejects: failures
     * resolve to a result with `success: false`.
     */
    private renderSlideWithSession;
    /**
     * Loads a slide's layout/master inheritance chain. Load failures never
     * reject: they degrade to an absent layout or master and push a
     * 'layout-load-failed' warning to the given collector.
     */
    private loadInheritanceChain;
}
/**
 * Opens a PPTX presentation for repeated rendering: the archive is opened
 * and parsed once, and the resolved theme and parsed parts are reused
 * across every render call on the returned {@link PptxDocument}.
 *
 * Rejects if the input is not a readable PPTX. Call
 * {@link PptxDocument.close} when done (or use `await using`).
 *
 * @param input File path or Buffer containing PPTX data
 * @param options Open options (log level)
 */
export declare function openPresentation(input: Buffer | string, options?: OpenOptions): Promise<PptxDocument>;
//# sourceMappingURL=PptxDocument.d.ts.map