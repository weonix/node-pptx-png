import { mergeRenderOptions, SIZE_PRESET_WIDTHS } from '../types/options.js';
import { EMU_PER_PIXEL_96DPI } from './UnitConverter.js';
import { PptxParser } from './PptxParser.js';
import { ThemeResolver } from '../theme/ThemeResolver.js';
import { SlideRenderer } from '../rendering/SlideRenderer.js';
import { FontManager } from '../text/FontManager.js';
import { setCustomFontFallbacks } from '../text/FontResolver.js';
import { createLogger } from '../utils/Logger.js';
import { WarningCollector, dedupeWarnings } from '../utils/WarningCollector.js';
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
export class PptxDocument {
    /** Number of slides in the presentation. */
    slideCount;
    /** Native slide size in EMU (914400 EMU = 1 inch). */
    size;
    parser;
    logger;
    /** Resolved presentation-default theme, cached after the first render
     *  (promise-cached so concurrent first calls resolve it only once). */
    themePromise = null;
    /** Per-master resolved themes (promise-cached, keyed by master part
     *  path). Multi-master decks reference different themes; each slide's
     *  colors resolve against the theme of its own master chain. */
    themeByMasterPath = new Map();
    /** Embedded-font extraction + registration, performed at most once per
     *  document (registration is global to the process font library). */
    embeddedFontsPromise = null;
    /** Presentation-level warnings (font preparation and other once-per-open
     *  events); slideNumber stays undefined on these. Surfaced through the
     *  renderAll/renderPresentation aggregate. */
    docWarnings = new WarningCollector();
    closed = false;
    constructor(parser, presentation, logger) {
        this.parser = parser;
        this.logger = logger;
        this.slideCount = presentation.slideCount;
        this.size = Object.freeze({
            widthEmu: presentation.slideWidth,
            heightEmu: presentation.slideHeight,
        });
    }
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
    static async open(input, options = {}, logger) {
        const docLogger = logger ?? createLogger(options.logLevel ?? 'warn', 'PptxDocument');
        const parser = new PptxParser(docLogger.child('Parser'));
        await parser.open(input);
        try {
            const presentation = await parser.getPresentation();
            return new PptxDocument(parser, presentation, docLogger);
        }
        catch (error) {
            parser.close();
            throw error;
        }
    }
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
    async slide(slideNumber, options = {}) {
        this.ensureNotClosed();
        if (!this.isValidSlideNumber(slideNumber)) {
            return this.invalidSlideResult(slideNumber, `Slide number ${slideNumber} out of range (1-${this.slideCount})`);
        }
        let session;
        try {
            session = await this.createRenderSession(options);
        }
        catch (error) {
            return this.invalidSlideResult(slideNumber, error instanceof Error ? error.message : String(error));
        }
        return this.renderSlideWithSession(session, slideNumber - 1);
    }
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
    async *slides(options = {}) {
        this.ensureNotClosed();
        const { slides, signal, onProgress, ...renderOptions } = options;
        signal?.throwIfAborted();
        const slideNumbers = this.resolveSlideSelection(slides ?? renderOptions.slideNumbers);
        const session = await this.createRenderSession(renderOptions);
        const total = slideNumbers.length;
        let done = 0;
        for (const slideNumber of slideNumbers) {
            this.ensureNotClosed();
            signal?.throwIfAborted();
            const result = this.isValidSlideNumber(slideNumber)
                ? await this.renderSlideWithSession(session, slideNumber - 1)
                : this.invalidSlideResult(slideNumber, `Slide number ${slideNumber} out of range (1-${this.slideCount})`);
            done++;
            onProgress?.({ done, total, slideNumber });
            yield result;
        }
    }
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
    async renderAll(options = {}) {
        this.ensureNotClosed();
        try {
            const session = await this.createRenderSession(options);
            this.logger.info('Rendering presentation', {
                slideCount: this.slideCount,
                slideWidth: this.size.widthEmu,
                slideHeight: this.size.heightEmu,
            });
            const slideNumbers = session.options.slideNumbers ?? Array.from({ length: this.slideCount }, (_, i) => i + 1);
            const slides = [];
            const errors = [];
            let successfulSlides = 0;
            for (const slideNumber of slideNumbers) {
                if (!this.isValidSlideNumber(slideNumber)) {
                    const message = `Slide number ${slideNumber} out of range (1-${this.slideCount})`;
                    this.logger.warn(message);
                    slides.push(this.invalidSlideResult(slideNumber, message));
                    errors.push({ level: 'slide', slideIndex: slideNumber - 1, message });
                    continue;
                }
                const result = await this.renderSlideWithSession(session, slideNumber - 1);
                slides.push(result);
                if (result.success) {
                    successfulSlides++;
                }
                else {
                    errors.push({
                        level: 'slide',
                        slideIndex: result.slideIndex,
                        message: result.errorMessage ?? 'Unknown slide rendering error',
                        stack: result.errorStack,
                    });
                }
            }
            // Aggregate: presentation-level warnings (slideNumber undefined)
            // first, then per-slide warnings, deduped by code+message+slideNumber
            const warnings = dedupeWarnings([
                ...this.docWarnings.list(),
                ...session.sessionWarnings.list(),
                ...slides.flatMap((slide) => slide.warnings ?? []),
            ]);
            return {
                slides,
                totalSlides: this.slideCount,
                successfulSlides,
                allSuccessful: successfulSlides === slides.length && errors.length === 0,
                ...(errors.length > 0 ? { errors } : {}),
                ...(warnings.length > 0 ? { warnings } : {}),
            };
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const stack = error instanceof Error ? error.stack : undefined;
            this.logger.error('Failed to render presentation', { error: message });
            return {
                slides: [],
                totalSlides: 0,
                successfulSlides: 0,
                allSuccessful: false,
                errors: [{ level: 'presentation', message, stack }],
            };
        }
    }
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
    async exportPdf(options = {}) {
        this.ensureNotClosed();
        const { slides, ...renderOptions } = options;
        const session = await this.createRenderSession(renderOptions);
        const slideNumbers = this.resolveSlideSelection(slides ?? renderOptions.slideNumbers);
        if (slideNumbers.length === 0) {
            throw new Error('exportPdf: no slides selected');
        }
        for (const slideNumber of slideNumbers) {
            if (!this.isValidSlideNumber(slideNumber)) {
                throw new Error(`exportPdf: slide number ${slideNumber} out of range (1-${this.slideCount})`);
            }
        }
        const { width, height } = session.renderer.getTargetDimensions(this.size.widthEmu, this.size.heightEmu);
        this.logger.info('Exporting PDF', { pages: slideNumbers.length, width, height });
        // One canvas, one page per slide: the canvas starts with zero pages and
        // newPage() appends one per slide, so the PDF page count matches exactly
        const canvas = session.renderer.createCanvas(width, height);
        for (const slideNumber of slideNumbers) {
            this.ensureNotClosed();
            const slideIndex = slideNumber - 1;
            const ctx = canvas.newPage(width, height);
            const slideData = await this.parser.getSlide(slideIndex);
            // Warnings are collected per slide for logging parity with the raster
            // path, but a Buffer result has no channel to surface them on
            const warnings = new WarningCollector({ slideNumber });
            const chain = await this.loadInheritanceChain(slideData, slideIndex, warnings);
            const theme = await this.resolveThemeForSlide(chain.masterPath);
            const renderer = this.rendererForTheme(session, theme);
            await renderer.renderSlideToContext(canvas, ctx, this.parser, slideData, this.size.widthEmu, this.size.heightEmu, chain.layoutNode, chain.masterNode, chain.layoutPath, chain.masterPath, warnings);
        }
        return canvas.toBuffer('pdf');
    }
    /**
     * Closes the document and releases the underlying archive and caches.
     * Idempotent: calling close() again is a no-op. Any render call after
     * close() throws a clear error.
     */
    close() {
        if (this.closed) {
            return;
        }
        this.closed = true;
        this.parser.close();
        this.logger.debug('PptxDocument closed');
    }
    /**
     * `await using` support: disposing the document closes it.
     */
    [Symbol.asyncDispose]() {
        this.close();
        return Promise.resolve();
    }
    /**
     * `using` support: disposing the document closes it.
     */
    [Symbol.dispose]() {
        this.close();
    }
    /**
     * Throws if the document has been closed.
     */
    ensureNotClosed() {
        if (this.closed) {
            throw new Error('PptxDocument is closed. Open a new document with openPresentation().');
        }
    }
    /**
     * Whether a 1-based slide number refers to a slide in this deck.
     */
    isValidSlideNumber(slideNumber) {
        return Number.isInteger(slideNumber) && slideNumber >= 1 && slideNumber <= this.slideCount;
    }
    /**
     * Builds a failure result for an invalid slide number or bad per-call
     * options, mirroring the failure-result shape of successful render paths.
     */
    invalidSlideResult(slideNumber, message) {
        return {
            slideIndex: slideNumber - 1,
            slideNumber,
            imageData: Buffer.alloc(0),
            width: 0,
            height: 0,
            success: false,
            errorMessage: message,
        };
    }
    /**
     * Expands a slides() selection into an ordered list of 1-based slide
     * numbers. List form is returned as-is (order preserved); range form is
     * clamped to the deck and expanded ascending; undefined means all slides.
     */
    resolveSlideSelection(selection) {
        if (Array.isArray(selection)) {
            return [...selection];
        }
        const from = Math.max(1, Math.floor(selection?.from ?? 1));
        const to = Math.min(this.slideCount, Math.floor(selection?.to ?? this.slideCount));
        const slideNumbers = [];
        for (let slideNumber = from; slideNumber <= to; slideNumber++) {
            slideNumbers.push(slideNumber);
        }
        return slideNumbers;
    }
    /**
     * Creates the per-call render session: merges options, applies the size
     * precedence (explicit width/height > scale > preset > default), prepares
     * fonts, resolves the theme (cached on the document), and builds the
     * SlideRenderer that every slide in the call will share.
     *
     * Throws on invalid scale or preset values.
     */
    async createRenderSession(options) {
        const merged = mergeRenderOptions(options);
        const hasExplicitSize = options.width !== undefined || options.height !== undefined;
        const resolved = this.resolveTargetSize(merged, hasExplicitSize);
        // Register embedded/user fonts and set custom fallbacks before any
        // text is measured or rendered. Per-call font problems collect into the
        // session so they do not leak into later renders' aggregates.
        const sessionWarnings = new WarningCollector();
        await this.prepareFonts(resolved.fonts, sessionWarnings);
        const theme = await this.resolveThemeOnce();
        const renderer = new SlideRenderer(theme, resolved, this.logger.child('Slide'));
        const renderersByTheme = new Map([[theme, renderer]]);
        return { options: resolved, renderer, renderersByTheme, sessionWarnings };
    }
    /**
     * Applies the scale and preset options to produce concrete pixel
     * dimensions from the slide's native size at 96 DPI.
     * Precedence: explicit width/height > scale > preset > default width.
     * The preset is ignored (lenient conflict handling) when width, height,
     * or scale is present.
     */
    resolveTargetSize(options, hasExplicitSize) {
        const { scale, preset } = options;
        if (hasExplicitSize) {
            return options;
        }
        if (scale !== undefined) {
            if (!Number.isFinite(scale) || scale <= 0) {
                throw new Error(`Invalid scale option: ${scale}. Expected a positive finite number.`);
            }
            return {
                ...options,
                width: Math.max(1, Math.round((this.size.widthEmu / EMU_PER_PIXEL_96DPI) * scale)),
                height: Math.max(1, Math.round((this.size.heightEmu / EMU_PER_PIXEL_96DPI) * scale)),
            };
        }
        if (preset !== undefined) {
            // Runtime lookup tolerates unknown strings from untyped callers
            const presetWidth = SIZE_PRESET_WIDTHS[preset];
            if (presetWidth === undefined) {
                throw new Error(`Invalid preset option: ${String(preset)}. ` +
                    `Expected one of: ${Object.keys(SIZE_PRESET_WIDTHS).join(', ')}.`);
            }
            // Height stays undefined so it is auto-calculated from the aspect ratio
            return { ...options, width: presetWidth, height: undefined };
        }
        return options;
    }
    /**
     * Prepares fonts before rendering: sets custom fallback chains, registers
     * user-supplied fonts, and extracts + registers fonts embedded in the
     * PPTX (unless fonts.useEmbeddedFonts is false). Embedded-font extraction
     * runs at most once per document; registration is process-global, so
     * later calls reuse it.
     *
     * Font problems never fail a render; they degrade to the fallback chains.
     */
    async prepareFonts(fonts, sessionWarnings) {
        // Always set (or clear) session fallbacks so a previous render's custom
        // chains do not leak into this one.
        // KNOWN LIMITATION: the fallback table is process-global, so concurrent
        // render calls (or two open documents) using DIFFERENT fonts.fallbacks
        // race — the last call's chains win for all in-flight renders. Scoping
        // it per render requires threading fallbacks into every FontResolver
        // construction; tracked as an API v2 follow-up.
        setCustomFontFallbacks(fonts?.fallbacks);
        // Font preparation happens once per open (not per slide), so its
        // warnings carry no slide attribution: they go to the document-level
        // collector and surface in the presentation aggregate.
        const fontManager = new FontManager({
            logger: this.logger.child('Fonts'),
            warnings: this.docWarnings,
        });
        if (fonts?.register && fonts.register.length > 0) {
            try {
                const registered = await fontManager.registerFonts(fonts.register);
                this.logger.info('User fonts registered', { families: registered });
            }
            catch (error) {
                this.logger.warn('User font registration failed', {
                    error: error instanceof Error ? error.message : String(error),
                });
                sessionWarnings.push({
                    code: 'other',
                    message: `User font registration failed: ${error instanceof Error ? error.message : String(error)}`,
                });
            }
        }
        if (fonts?.useEmbeddedFonts !== false) {
            this.embeddedFontsPromise ??= this.registerEmbeddedFonts(fontManager);
            await this.embeddedFontsPromise;
        }
    }
    /**
     * Extracts and registers fonts embedded in the PPTX. Never rejects; font
     * problems are logged as warnings so renders degrade gracefully.
     */
    async registerEmbeddedFonts(fontManager) {
        try {
            const embedded = await this.parser.getEmbeddedFonts();
            if (embedded.length > 0) {
                const result = await fontManager.registerEmbeddedFonts(embedded);
                this.logger.info('Embedded fonts registered', {
                    registered: result.registered,
                    skipped: result.skipped.map((s) => `${s.typeface}: ${s.reason}`),
                });
            }
        }
        catch (error) {
            this.logger.warn('Embedded font extraction failed', {
                error: error instanceof Error ? error.message : String(error),
            });
            this.docWarnings.push({
                code: 'embedded-font-unsupported',
                message: `Embedded font extraction failed: ${error instanceof Error ? error.message : String(error)}`,
            });
        }
    }
    /**
     * Resolves the theme once per document and caches it.
     */
    resolveThemeOnce() {
        this.themePromise ??= new ThemeResolver(this.logger.child('Theme')).resolveTheme(this.parser);
        return this.themePromise;
    }
    /**
     * Resolves the theme for a slide from its master chain. Each master has
     * its own theme relationship, and multi-master decks can carry themes
     * with entirely different color schemes — a slide's scheme colors
     * (accent1, tx2, ...) must resolve against ITS master's theme, not a
     * single presentation-wide one. Falls back to the presentation-default
     * theme when the master is unknown or its theme cannot be loaded.
     * Results are promise-cached per master path.
     */
    resolveThemeForSlide(masterPath) {
        if (!masterPath) {
            return this.resolveThemeOnce();
        }
        let themePromise = this.themeByMasterPath.get(masterPath);
        if (!themePromise) {
            themePromise = this.parser
                .getThemeForMaster(masterPath)
                .then((themeData) => themeData
                ? new ThemeResolver(this.logger.child('Theme')).parseTheme(themeData)
                : this.resolveThemeOnce())
                .catch((error) => {
                this.logger.warn('Failed to resolve master theme, using default', {
                    masterPath,
                    error: error instanceof Error ? error.message : String(error),
                });
                return this.resolveThemeOnce();
            });
            this.themeByMasterPath.set(masterPath, themePromise);
        }
        return themePromise;
    }
    /**
     * Returns the session renderer for a resolved theme, creating and caching
     * one per distinct theme. The default theme's renderer is seeded at
     * session creation, so single-theme decks always reuse it.
     */
    rendererForTheme(session, theme) {
        let renderer = session.renderersByTheme.get(theme);
        if (!renderer) {
            renderer = new SlideRenderer(theme, session.options, this.logger.child('Slide'));
            session.renderersByTheme.set(theme, renderer);
        }
        return renderer;
    }
    /**
     * Renders one slide (0-based index) with an established session, loading
     * the slide's layout/master inheritance chain. Never rejects: failures
     * resolve to a result with `success: false`.
     */
    async renderSlideWithSession(session, slideIndex) {
        // Per-slide collector: every warning pushed during this render is
        // attributed to this slide's 1-based number
        const warnings = new WarningCollector({ slideNumber: slideIndex + 1 });
        try {
            // Get slide data
            const slideData = await this.parser.getSlide(slideIndex);
            // Get layout and master for inheritance chain
            const chain = await this.loadInheritanceChain(slideData, slideIndex, warnings);
            // Render with the theme of this slide's master chain (multi-master
            // decks reference different themes with different color schemes)
            const theme = await this.resolveThemeForSlide(chain.masterPath);
            const renderer = this.rendererForTheme(session, theme);
            // Render the slide
            const output = await renderer.renderSlide(this.parser, slideData, this.size.widthEmu, this.size.heightEmu, chain.layoutNode, chain.masterNode, chain.layoutPath, chain.masterPath, warnings);
            const collected = warnings.list();
            return {
                slideIndex,
                slideNumber: slideIndex + 1,
                imageData: output.imageData,
                width: output.width,
                height: output.height,
                success: output.success,
                errorMessage: output.error,
                ...(collected.length > 0 ? { warnings: collected } : {}),
            };
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const stack = error instanceof Error ? error.stack : undefined;
            this.logger.error('Failed to render slide', {
                slideIndex,
                error: message,
            });
            const collected = warnings.list();
            return {
                slideIndex,
                slideNumber: slideIndex + 1,
                imageData: Buffer.alloc(0),
                width: 0,
                height: 0,
                success: false,
                errorMessage: message,
                errorStack: stack,
                ...(collected.length > 0 ? { warnings: collected } : {}),
            };
        }
    }
    /**
     * Loads a slide's layout/master inheritance chain. Load failures never
     * reject: they degrade to an absent layout or master and push a
     * 'layout-load-failed' warning to the given collector.
     */
    async loadInheritanceChain(slideData, slideIndex, warnings) {
        let layoutNode = undefined;
        let layoutPath = undefined;
        let masterNode = undefined;
        let masterPath = undefined;
        if (slideData.layoutRelId) {
            try {
                const layoutData = await this.parser.getSlideLayout(slideData.path, slideData.layoutRelId);
                layoutNode = layoutData.content;
                layoutPath = layoutData.path;
                if (layoutData.masterRelId) {
                    try {
                        const masterData = await this.parser.getSlideMaster(layoutData.path, layoutData.masterRelId);
                        masterNode = masterData.content;
                        masterPath = masterData.path;
                    }
                    catch (error) {
                        this.logger.warn('Failed to load slide master', {
                            slideIndex,
                            error: error instanceof Error ? error.message : String(error),
                        });
                        warnings.push({
                            code: 'layout-load-failed',
                            message: `Failed to load slide master: ${error instanceof Error ? error.message : String(error)}`,
                        });
                    }
                }
            }
            catch (error) {
                this.logger.warn('Failed to load slide layout', {
                    slideIndex,
                    error: error instanceof Error ? error.message : String(error),
                });
                warnings.push({
                    code: 'layout-load-failed',
                    message: `Failed to load slide layout: ${error instanceof Error ? error.message : String(error)}`,
                });
            }
        }
        return { layoutNode, masterNode, layoutPath, masterPath };
    }
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
export async function openPresentation(input, options = {}) {
    return PptxDocument.open(input, options);
}
//# sourceMappingURL=PptxDocument.js.map