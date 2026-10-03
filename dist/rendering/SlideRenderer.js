import { Canvas } from 'skia-canvas';
import { mergeRenderOptions, Colors } from '../types/index.js';
import { getXmlChild, getXmlAttr, getOrderedChildren, hasXmlChild } from '../core/PptxParser.js';
import { UnitConverter } from '../core/UnitConverter.js';
import { resolveEffectiveColorMap } from '../theme/ColorResolver.js';
import { SHAPE_ELEMENT_TYPES } from '../core/constants.js';
import { BackgroundRenderer } from './BackgroundRenderer.js';
import { ShapeRenderer } from './ShapeRenderer.js';
import { GroupShapeRenderer } from './GroupShapeRenderer.js';
import { AlternateContentRenderer } from './AlternateContentRenderer.js';
import { DiagramRenderer, DIAGRAM_GRAPHIC_DATA_URI } from './DiagramRenderer.js';
import { ChartRenderer } from './ChartRenderer.js';
import { ChartParser } from '../parsers/ChartParser.js';
import { TableRenderer } from './TableRenderer.js';
import { TableStyleParser } from '../parsers/TableStyleParser.js';
import { createLogger } from '../utils/Logger.js';
import { PngOptimizer } from '../utils/PngOptimizer.js';
/**
 * Maps each top-level spTree element to the non-visual-properties wrapper
 * that can carry its placeholder marker (nv*Pr > p:nvPr > p:ph).
 */
const NV_PR_WRAPPER_BY_TAG = {
    'p:sp': 'p:nvSpPr',
    'p:pic': 'p:nvPicPr',
    'p:graphicFrame': 'p:nvGraphicFramePr',
    'p:cxnSp': 'p:nvCxnSpPr',
    'p:grpSp': 'p:nvGrpSpPr',
};
/**
 * a:graphicData URI of an embedded OLE object (ECMA-376 §19.3.2.4), e.g. an
 * Excel worksheet range pasted as an embedded object.
 */
const OLE_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole';
/**
 * Reads an xsd:boolean XML attribute ("0"/"false"/"1"/"true"), returning
 * `defaultValue` when the attribute is absent.
 */
function parseXmlBoolAttr(node, attr, defaultValue) {
    const raw = getXmlAttr(node, attr);
    if (raw === undefined)
        return defaultValue;
    return raw !== '0' && raw !== 'false';
}
/**
 * Whether a top-level spTree element is a placeholder (carries p:ph inside
 * its nv*Pr > p:nvPr). A bare self-closed `<p:ph/>` parses to an empty
 * string, so presence is checked rather than truthiness.
 */
function isPlaceholderElement(tagName, node) {
    const wrapperTag = NV_PR_WRAPPER_BY_TAG[tagName];
    if (!wrapperTag)
        return false;
    const nvPr = getXmlChild(getXmlChild(node, wrapperTag), 'p:nvPr');
    return hasXmlChild(nvPr, 'p:ph');
}
/**
 * Renders individual slides to images.
 */
export class SlideRenderer {
    logger;
    theme;
    options;
    backgroundRenderer;
    unitConverter;
    chartParser;
    alternateContentRenderer;
    pngOptimizer;
    pngOptimizerInitialized = false;
    /** Table styles parsed once per parser (multi-slide renders reuse it) */
    tableStylesCache = new WeakMap();
    /**
     * Ordered-parse results cached per parser+part path. Master/layout parts
     * are re-rendered for every slide that uses them; the parser already
     * caches the raw XML string, and this cache additionally avoids repeating
     * the preserveOrder parse of identical part content across slides.
     */
    orderedDocCache = new WeakMap();
    constructor(theme, options = {}, logger) {
        this.logger = logger ?? createLogger('warn', 'SlideRenderer');
        this.theme = theme;
        // mergeRenderOptions ignores explicitly-undefined keys so they cannot
        // clobber defaults (e.g. { width: undefined }).
        this.options = mergeRenderOptions(options);
        this.backgroundRenderer = new BackgroundRenderer(theme, this.logger.child('Background'));
        this.unitConverter = new UnitConverter();
        this.chartParser = new ChartParser({ theme, logger: this.logger.child('ChartParser') });
        this.alternateContentRenderer = new AlternateContentRenderer({
            logger: this.logger.child('AltContent'),
        });
        this.pngOptimizer = new PngOptimizer(this.logger.child('PngOptimizer'));
    }
    /**
     * Renders a slide to an image buffer.
     *
     * @param parser PPTX parser instance
     * @param slideData Slide data including path and content
     * @param slideWidthEmu Slide width in EMU
     * @param slideHeightEmu Slide height in EMU
     * @param layoutNode Optional layout XML node
     * @param masterNode Optional master XML node
     * @param layoutPath Optional path to the layout file for relationship resolution
     * @param masterPath Optional path to the master file for relationship resolution
     * @param warnings Optional structured warning collector; fidelity events
     *   from this render pass (missing fonts, unsupported elements, ...) are
     *   pushed to it in addition to being logged
     */
    async renderSlide(parser, slideData, slideWidthEmu, slideHeightEmu, layoutNode, masterNode, layoutPath, masterPath, warnings) {
        try {
            // Calculate target dimensions
            const { targetWidth, targetHeight } = this.calculateDimensions(slideWidthEmu, slideHeightEmu);
            // Create canvas (applies the gpu option)
            const canvas = this.createCanvas(targetWidth, targetHeight);
            const ctx = canvas.getContext('2d');
            await this.renderSlideToContext(canvas, ctx, parser, slideData, slideWidthEmu, slideHeightEmu, layoutNode, masterNode, layoutPath, masterPath, warnings);
            // Export to image
            const imageData = await this.exportCanvas(canvas);
            return {
                imageData,
                width: targetWidth,
                height: targetHeight,
                success: true,
            };
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to render slide', {
                index: slideData.index,
                error: message,
            });
            return {
                imageData: Buffer.alloc(0),
                width: 0,
                height: 0,
                success: false,
                error: message,
            };
        }
    }
    /**
     * Renders a slide onto a caller-provided canvas context. This is the
     * drawing path shared by {@link renderSlide} (which creates its own
     * single-page canvas and exports it) and multi-page exports such as
     * PptxDocument.exportPdf (which draw each slide onto a `canvas.newPage()`
     * context of one shared canvas).
     *
     * The context is assumed to cover the pixel dimensions reported by
     * {@link getTargetDimensions} for this slide size. Unlike renderSlide,
     * this method throws on failure instead of returning an error result.
     *
     * @param canvas Canvas the context belongs to (used for background sizing)
     * @param ctx 2D context to draw the slide onto
     * @param parser PPTX parser instance
     * @param slideData Slide data including path and content
     * @param slideWidthEmu Slide width in EMU
     * @param slideHeightEmu Slide height in EMU
     * @param layoutNode Optional layout XML node
     * @param masterNode Optional master XML node
     * @param layoutPath Optional path to the layout file for relationship resolution
     * @param masterPath Optional path to the master file for relationship resolution
     * @param warnings Optional structured warning collector for this render pass
     */
    async renderSlideToContext(canvas, ctx, parser, slideData, slideWidthEmu, slideHeightEmu, layoutNode, masterNode, layoutPath, masterPath, warnings) {
        const { targetWidth, targetHeight, scaleX, scaleY } = this.calculateDimensions(slideWidthEmu, slideHeightEmu);
        this.logger.info('Rendering slide', {
            index: slideData.index,
            width: targetWidth,
            height: targetHeight,
            scaleX: scaleX.toFixed(4),
            scaleY: scaleY.toFixed(4),
        });
        // Render background (async to support picture backgrounds)
        await this.backgroundRenderer.renderBackgroundAsync(ctx, canvas, slideData.content, parser, slideData.path, layoutNode, layoutPath, masterNode, masterPath, this.options.backgroundColor);
        // Load table styles for a:tableStyleId resolution, cached per parser
        // so multi-slide renders parse tableStyles.xml once
        const effectiveColorMap = resolveEffectiveColorMap(masterNode, layoutNode, slideData.content);
        let tableStyles = this.tableStylesCache.get(parser);
        if (!tableStyles) {
            tableStyles = await new TableStyleParser({
                logger: this.logger.child('TableStyles'),
                theme: this.theme,
                colorMap: effectiveColorMap,
            }).load(parser);
            this.tableStylesCache.set(parser, tableStyles);
        }
        // Presentation root for the p:defaultTextStyle inheritance layer
        // (readXml caches the part, so this is cheap per slide)
        const presentationNode = (await parser.getPresentation()).content;
        const stateParams = {
            parser,
            slideData,
            scaleX,
            scaleY,
            effectiveColorMap,
            tableStyles,
            presentationNode,
            layoutNode,
            masterNode,
            warnings,
        };
        // Inherited master/layout shapes render UNDER the slide's own shapes:
        // background < master spTree < layout spTree < slide spTree. Their
        // placeholder shapes are excluded — those only donate inheritance to the
        // slide's placeholders and never draw directly.
        //
        // Visibility follows the showMasterSp cascade (ECMA-376 §19.3.1.38 p:sld,
        // §19.3.1.39 p:sldLayout — each gates whether "shapes on the slide
        // master" are shown one level up the design chain, where the layout acts
        // as the slide's immediate master): layout shapes show on the slide iff
        // p:sld@showMasterSp != 0 (PowerPoint's "Hide Background Graphics"
        // checkbox), and master shapes show iff additionally
        // p:sldLayout@showMasterSp != 0. This matches PowerPoint's OM
        // (Slide.DisplayMasterShapes / CustomLayout.DisplayMasterShapes) and
        // Apache POI's XSLFSheet.getFollowMasterGraphics cascade.
        const slideShowsInherited = parseXmlBoolAttr(slideData.content, 'showMasterSp', true);
        if (slideShowsInherited) {
            const layoutShowsMaster = layoutNode === undefined || parseXmlBoolAttr(layoutNode, 'showMasterSp', true);
            if (masterNode && masterPath && layoutShowsMaster) {
                const masterState = this.createRenderState(stateParams, masterPath, 'p:sldMaster');
                await this.renderShapeTreeOrdered(ctx, masterState);
            }
            if (layoutNode && layoutPath) {
                const layoutState = this.createRenderState(stateParams, layoutPath, 'p:sldLayout');
                await this.renderShapeTreeOrdered(ctx, layoutState);
            }
        }
        // Render the slide's own shape tree (Phase 2+5) on top.
        // Use ordered parsing to preserve z-order of interleaved elements
        const renderState = this.createRenderState(stateParams, slideData.path, 'p:sld');
        await this.renderShapeTreeOrdered(ctx, renderState);
    }
    /**
     * Builds the per-part render state used by {@link renderShapeTreeOrdered}.
     * The same slide-scoped context (theme, color map, inheritance nodes,
     * slide number) is shared by all parts of one slide render; only the
     * source part path differs so relationships (images, charts, diagrams)
     * resolve against the right .rels file.
     */
    createRenderState(params, sourcePath, rootTag) {
        const { parser, slideData, scaleX, scaleY, effectiveColorMap, tableStyles, presentationNode, layoutNode, masterNode, warnings, } = params;
        return {
            parser,
            slidePath: sourcePath,
            rootTag,
            skipPlaceholders: rootTag !== 'p:sld',
            scaleX,
            scaleY,
            warnings,
            shapeRenderer: new ShapeRenderer({
                theme: this.theme,
                scaleX,
                scaleY,
                parser,
                sourcePath,
                slideNode: slideData.content,
                layoutNode,
                masterNode,
                presentationNode,
                slideNumber: slideData.index + 1,
                logger: this.logger.child('Shape'),
                warnings,
            }),
            groupShapeRenderer: new GroupShapeRenderer({
                scaleX,
                scaleY,
                logger: this.logger.child('GroupShape'),
            }),
            tableRenderer: new TableRenderer({
                theme: this.theme,
                colorMap: effectiveColorMap,
                scaleX,
                scaleY,
                tableStyles,
                slideNumber: slideData.index + 1,
                logger: this.logger.child('Table'),
                warnings,
            }),
            diagramRenderer: new DiagramRenderer({
                theme: this.theme,
                parser,
                slidePath: sourcePath,
                slideNode: slideData.content,
                layoutNode,
                masterNode,
                logger: this.logger.child('Diagram'),
                warnings,
            }),
            chartRenderer: new ChartRenderer({
                logger: this.logger.child('Chart'),
                warnings,
            }),
        };
    }
    /**
     * Creates a canvas of the given pixel size, applying the `gpu` render
     * option: an explicit boolean sets `canvas.gpu`, while 'auto' leaves the
     * skia-canvas default untouched. The mode in effect is recorded in the
     * debug log. Multi-page exports (PDF) call this once and then draw each
     * slide onto a `canvas.newPage()` context.
     */
    createCanvas(width, height) {
        const textOptions = this.options.textContrast !== undefined || this.options.textGamma !== undefined
            ? {
                ...(this.options.textContrast !== undefined
                    ? { textContrast: this.options.textContrast }
                    : {}),
                ...(this.options.textGamma !== undefined ? { textGamma: this.options.textGamma } : {}),
            }
            : undefined;
        const canvas = textOptions ? new Canvas(width, height, textOptions) : new Canvas(width, height);
        if (typeof this.options.gpu === 'boolean') {
            canvas.gpu = this.options.gpu;
        }
        this.logger.debug('Canvas created', {
            width,
            height,
            gpuRequested: this.options.gpu,
            gpuEnabled: canvas.gpu,
        });
        return canvas;
    }
    /**
     * Computes the output pixel dimensions these options produce for a slide
     * of the given native EMU size (explicit width/height, or width with
     * height derived from the aspect ratio).
     */
    getTargetDimensions(slideWidthEmu, slideHeightEmu) {
        const { targetWidth, targetHeight } = this.calculateDimensions(slideWidthEmu, slideHeightEmu);
        return { width: targetWidth, height: targetHeight };
    }
    /**
     * Calculates target dimensions maintaining aspect ratio.
     */
    calculateDimensions(slideWidthEmu, slideHeightEmu) {
        // Convert EMU to pixels at 96 DPI
        const nativeWidth = this.unitConverter.emuToPixels(slideWidthEmu);
        const nativeHeight = this.unitConverter.emuToPixels(slideHeightEmu);
        const aspectRatio = nativeWidth / nativeHeight;
        let targetWidth;
        let targetHeight;
        if (this.options.height !== undefined) {
            // Both width and height specified
            targetWidth = this.options.width;
            targetHeight = this.options.height;
        }
        else {
            // Only width specified, calculate height from aspect ratio
            targetWidth = this.options.width;
            targetHeight = Math.round(targetWidth / aspectRatio);
        }
        // Calculate scale factors
        const scaleX = targetWidth / nativeWidth;
        const scaleY = targetHeight / nativeHeight;
        return { targetWidth, targetHeight, scaleX, scaleY };
    }
    /**
     * Exports the canvas to an image buffer in the configured format.
     *
     * Lossy formats (jpeg, webp) use the `quality` option (0-1), falling back
     * to the legacy `jpegQuality` (1-100). SVG returns the UTF-8 document
     * bytes. `pngOptimization` only applies to PNG output.
     */
    async exportCanvas(canvas) {
        const format = this.options.format;
        if (format === 'jpeg' || format === 'webp') {
            return canvas.toBuffer(format, { quality: this.resolveQuality() });
        }
        if (format === 'svg') {
            return canvas.toBuffer('svg');
        }
        // Get raw PNG from canvas
        const pngBuffer = await canvas.toBuffer('png');
        // Apply PNG optimization if configured
        const optimization = this.options.pngOptimization;
        if (optimization && optimization !== 'none') {
            // Initialize optimizer on first use (lazy loading of Sharp)
            if (!this.pngOptimizerInitialized) {
                await this.pngOptimizer.initialize();
                this.pngOptimizerInitialized = true;
            }
            if (this.pngOptimizer.isAvailable()) {
                const optimized = await this.pngOptimizer.optimize(pngBuffer, optimization);
                // Log compression stats in debug mode
                if (this.options.debugMode) {
                    const stats = this.pngOptimizer.getCompressionStats(pngBuffer, optimized);
                    this.logger.debug('PNG optimization', {
                        preset: typeof optimization === 'string' ? optimization : 'custom',
                        originalSize: stats.originalSize,
                        optimizedSize: stats.optimizedSize,
                        reduction: `${stats.reductionPercent}%`,
                    });
                }
                return optimized;
            }
        }
        return pngBuffer;
    }
    /**
     * Resolves the effective lossy-encoding quality (0-1): the `quality`
     * option when set (clamped to [0, 1]), otherwise the legacy `jpegQuality`
     * (1-100) scaled down.
     */
    resolveQuality() {
        if (this.options.quality !== undefined) {
            return Math.min(1, Math.max(0, this.options.quality));
        }
        return Math.min(1, Math.max(0, this.options.jpegQuality / 100));
    }
    /**
     * Converts EMU coordinates to pixel coordinates.
     */
    emuToPixels(context, emuX, emuY) {
        return {
            x: context.unitConverter.emuToPixels(emuX) * context.scaleX,
            y: context.unitConverter.emuToPixels(emuY) * context.scaleY,
        };
    }
    /**
     * Converts EMU dimensions to pixel dimensions.
     */
    emuToPixelSize(context, emuWidth, emuHeight) {
        return {
            width: context.unitConverter.emuToPixels(emuWidth) * context.scaleX,
            height: context.unitConverter.emuToPixels(emuHeight) * context.scaleY,
        };
    }
    /**
     * Gets the background color for contrast calculations.
     */
    getBackgroundColor(slideNode, layoutNode, masterNode) {
        return (this.backgroundRenderer.getBackgroundColor(slideNode, layoutNode, masterNode) ?? Colors.white);
    }
    /**
     * Renders all shapes in the shape tree using ordered parsing.
     * Uses the raw XML and parses with preserveOrder to maintain correct z-order.
     * Elements that appear later in the XML are rendered on top.
     */
    async renderShapeTreeOrdered(ctx, state) {
        // Get the part XML with preserved element order (cached per parser+path
        // so master/layout parts shared by many slides parse once)
        const orderedSlide = await this.readXmlOrderedCached(state.parser, state.slidePath);
        // Navigate to <root> > p:cSld > p:spTree in the ordered structure
        const spTreeChildren = this.navigateToSpTree(orderedSlide, state.rootTag);
        if (!spTreeChildren) {
            this.logger.debug('No shape tree found in part', { path: state.slidePath });
            return;
        }
        // Get all elements in document order using shared element type constants
        const orderedElements = getOrderedChildren(spTreeChildren, SHAPE_ELEMENT_TYPES);
        // Count elements for logging
        let shapeCount = 0;
        let connectionCount = 0;
        let pictureCount = 0;
        let groupCount = 0;
        let graphicFrameCount = 0;
        let alternateContentCount = 0;
        // Render each element in document order
        for (const { tagName, node } of orderedElements) {
            // Master/layout placeholders never draw directly on the slide — they
            // only donate inheritance to the slide's instantiated placeholders
            if (state.skipPlaceholders && isPlaceholderElement(tagName, node)) {
                continue;
            }
            try {
                switch (tagName) {
                    case 'p:sp':
                        await state.shapeRenderer.renderShape(ctx, node);
                        shapeCount++;
                        break;
                    case 'p:cxnSp':
                        state.shapeRenderer.renderConnectionShape(ctx, node);
                        connectionCount++;
                        break;
                    case 'p:pic':
                        // Render picture elements (Phase 4)
                        await state.shapeRenderer.renderPicture(ctx, node);
                        pictureCount++;
                        break;
                    case 'p:grpSp':
                        // Render group shapes (Phase 5)
                        await this.renderGroupShape(ctx, node, state);
                        groupCount++;
                        break;
                    case 'p:graphicFrame':
                        // Render graphic frames - charts, tables (Phase 5)
                        await this.renderGraphicFrame(ctx, node, state);
                        graphicFrameCount++;
                        break;
                    case 'mc:AlternateContent':
                        // Render alternate content - SmartArt fallbacks (Phase 5)
                        await this.renderAlternateContent(ctx, node, state);
                        alternateContentCount++;
                        break;
                }
            }
            catch (error) {
                this.logger.warn(`Failed to render ${tagName}`, {
                    error: error instanceof Error ? error.message : String(error),
                });
                state.warnings?.push({
                    code: 'other',
                    message: `Failed to render ${tagName}: ${error instanceof Error ? error.message : String(error)}`,
                    detail: { elementType: tagName },
                });
            }
        }
        this.logger.debug('Rendered shape tree', {
            shapes: shapeCount,
            connections: connectionCount,
            pictures: pictureCount,
            groups: groupCount,
            graphicFrames: graphicFrameCount,
            alternateContent: alternateContentCount,
        });
    }
    /**
     * Reads a part with preserved element order, caching the parsed result per
     * parser+path. The parser itself caches the raw XML string but re-runs the
     * preserveOrder parse on every readXmlOrdered call; master/layout parts are
     * rendered once per slide that uses them, so this cache makes those
     * repeat renders parse-free. Keyed by parser instance (like the table
     * styles cache) so a new document never sees stale entries.
     */
    async readXmlOrderedCached(parser, path) {
        // Only master/layout parts are read repeatedly (once per slide using
        // them); slide parts render exactly once, so caching them would only
        // grow memory linearly with deck size — and would serve stale content
        // if the parser instance is re-open()ed on another file.
        const reusablePart = path.includes('slideMasters/') || path.includes('slideLayouts/');
        if (!reusablePart) {
            return parser.readXmlOrdered(path);
        }
        let byPath = this.orderedDocCache.get(parser);
        if (!byPath) {
            byPath = new Map();
            this.orderedDocCache.set(parser, byPath);
        }
        const cached = byPath.get(path);
        if (cached) {
            return cached;
        }
        const parsed = await parser.readXmlOrdered(path);
        byPath.set(path, parsed);
        return parsed;
    }
    /**
     * Navigates the ordered XML structure to find the spTree children.
     * Returns the array of children in document order, or undefined if not found.
     *
     * @param orderedSlide Ordered parse of a slide, layout, or master part
     * @param rootTag Root element of the part ('p:sld', 'p:sldLayout', 'p:sldMaster')
     */
    navigateToSpTree(orderedSlide, rootTag) {
        // orderedSlide is an array that may contain XML declaration and the root element
        // Format: [{ '?xml': [...], ':@': {...} }, { 'p:sld': [...], ':@': {...} }]
        if (!orderedSlide || !Array.isArray(orderedSlide) || orderedSlide.length === 0) {
            return undefined;
        }
        // Find the root element (skip XML declaration and other elements)
        let sldChildren;
        for (const element of orderedSlide) {
            if (element?.[rootTag]) {
                sldChildren = element[rootTag];
                break;
            }
        }
        if (!sldChildren || !Array.isArray(sldChildren)) {
            return undefined;
        }
        // Find p:cSld in the children
        for (const child of sldChildren) {
            const cSldChildren = child['p:cSld'];
            if (cSldChildren && Array.isArray(cSldChildren)) {
                // Find p:spTree in p:cSld children
                for (const cSldChild of cSldChildren) {
                    const spTreeChildren = cSldChild['p:spTree'];
                    if (spTreeChildren && Array.isArray(spTreeChildren)) {
                        return spTreeChildren;
                    }
                }
            }
        }
        return undefined;
    }
    /**
     * Renders a group shape (p:grpSp) element.
     * @param ctx Canvas 2D context
     * @param grpSpNode Group shape XML node
     * @param state Internal render state
     * @param parentGroupTransform Optional parent group transform for nested groups
     */
    async renderGroupShape(ctx, grpSpNode, state, parentGroupTransform) {
        // Define the render callback for child elements
        const renderChild = async (childCtx, tagName, node, groupTransform) => {
            // Placeholders never draw from master/layout parts, at any nesting depth
            if (state.skipPlaceholders && isPlaceholderElement(tagName, node)) {
                return;
            }
            switch (tagName) {
                case 'p:sp':
                    this.renderShapeInGroup(childCtx, node, state, groupTransform);
                    break;
                case 'p:cxnSp':
                    this.renderConnectionShapeInGroup(childCtx, node, state, groupTransform);
                    break;
                case 'p:pic':
                    await this.renderPictureInGroup(childCtx, node, state, groupTransform);
                    break;
                case 'p:grpSp':
                    // Recursive group rendering
                    await this.renderGroupShape(childCtx, node, state, groupTransform);
                    break;
                case 'p:graphicFrame':
                    await this.renderGraphicFrame(childCtx, node, state, groupTransform);
                    break;
                case 'mc:AlternateContent':
                    await this.renderAlternateContent(childCtx, node, state, groupTransform);
                    break;
                default:
                    this.logger.debug('Unknown child element in group', { tagName });
            }
        };
        await state.groupShapeRenderer.renderGroupShape(ctx, grpSpNode, renderChild, parentGroupTransform);
    }
    /**
     * Renders a shape within a group, applying the group transform.
     */
    renderShapeInGroup(ctx, spNode, state, groupTransform) {
        // Get shape properties
        const spPr = getXmlChild(spNode, 'p:spPr');
        const shapeTransform = state.groupShapeRenderer.parseShapeTransform(spPr);
        if (!shapeTransform) {
            this.logger.debug('Shape in group has no transform, skipping');
            return;
        }
        // Apply group transform to get final pixel transform
        const pixelTransform = state.groupShapeRenderer.toPixelTransform(shapeTransform, groupTransform);
        // Use the ShapeRenderer's existing shape rendering with the modified transform
        // We render the shape by creating a temporary context state
        state.shapeRenderer.renderShapeWithTransform(ctx, spNode, pixelTransform);
    }
    /**
     * Renders a connection shape within a group.
     */
    renderConnectionShapeInGroup(ctx, cxnSpNode, state, groupTransform) {
        // Get shape properties
        const spPr = getXmlChild(cxnSpNode, 'p:spPr');
        const shapeTransform = state.groupShapeRenderer.parseShapeTransform(spPr);
        if (!shapeTransform) {
            this.logger.debug('Connection shape in group has no transform, skipping');
            return;
        }
        // Apply group transform
        const pixelTransform = state.groupShapeRenderer.toPixelTransform(shapeTransform, groupTransform);
        // Render the connection shape with the transformed coordinates
        state.shapeRenderer.renderConnectionShapeWithTransform(ctx, cxnSpNode, pixelTransform);
    }
    /**
     * Renders a picture within a group.
     */
    async renderPictureInGroup(ctx, picNode, state, groupTransform) {
        // Get shape properties
        const spPr = getXmlChild(picNode, 'p:spPr');
        const shapeTransform = state.groupShapeRenderer.parseShapeTransform(spPr);
        if (!shapeTransform) {
            this.logger.debug('Picture in group has no transform, skipping');
            return;
        }
        // Apply group transform
        const pixelTransform = state.groupShapeRenderer.toPixelTransform(shapeTransform, groupTransform);
        // Render the picture with the transformed coordinates
        await state.shapeRenderer.renderPictureWithTransform(ctx, picNode, pixelTransform);
    }
    /**
     * Renders a graphic frame element (p:graphicFrame).
     * Handles charts, tables, and other embedded content.
     * @param groupTransform Optional group transform when the frame is a group child
     */
    async renderGraphicFrame(ctx, graphicFrameNode, state, groupTransform) {
        // Get transform (position and size)
        const xfrm = getXmlChild(graphicFrameNode, 'p:xfrm');
        if (!xfrm) {
            this.logger.debug('GraphicFrame has no transform');
            return;
        }
        const off = getXmlChild(xfrm, 'a:off');
        const ext = getXmlChild(xfrm, 'a:ext');
        if (!off || !ext) {
            this.logger.debug('GraphicFrame missing offset or extent');
            return;
        }
        const x = parseInt(getXmlAttr(off, 'x') ?? '0', 10);
        const y = parseInt(getXmlAttr(off, 'y') ?? '0', 10);
        const cx = parseInt(getXmlAttr(ext, 'cx') ?? '0', 10);
        const cy = parseInt(getXmlAttr(ext, 'cy') ?? '0', 10);
        // Coordinates of a group child are expressed in the group's child
        // coordinate space; map them into slide space via the group transform.
        // Frame rotation is not supported for charts/tables, so only the
        // position and size of the transformed frame are used.
        let frameTransform = {
            x,
            y,
            width: cx,
            height: cy,
            rotation: 0,
            flipH: false,
            flipV: false,
        };
        if (groupTransform) {
            frameTransform = state.groupShapeRenderer.transformChildToSlide(frameTransform, groupTransform);
        }
        // Convert to pixels
        const bounds = {
            x: this.unitConverter.emuToPixels(frameTransform.x) * state.scaleX,
            y: this.unitConverter.emuToPixels(frameTransform.y) * state.scaleY,
            width: this.unitConverter.emuToPixels(frameTransform.width) * state.scaleX,
            height: this.unitConverter.emuToPixels(frameTransform.height) * state.scaleY,
        };
        if (bounds.width <= 0 || bounds.height <= 0) {
            this.logger.debug('GraphicFrame has invalid bounds');
            return;
        }
        // Get the graphic data
        const graphic = getXmlChild(graphicFrameNode, 'a:graphic');
        const graphicData = graphic ? getXmlChild(graphic, 'a:graphicData') : undefined;
        if (!graphicData) {
            this.logger.debug('GraphicFrame has no graphic data');
            return;
        }
        const uri = getXmlAttr(graphicData, 'uri');
        this.logger.debug('GraphicFrame URI', { uri });
        // Check if this is a chart
        if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
            await this.renderChartFrame(ctx, graphicData, bounds, state);
        }
        // Check if this is a table
        else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            this.renderTableFrame(ctx, graphicData, bounds, state);
        }
        // Check if this is a SmartArt diagram (rendered via its pre-baked drawing part)
        else if (uri === DIAGRAM_GRAPHIC_DATA_URI) {
            await state.diagramRenderer.renderDiagram(ctx, graphicData, bounds, { cx, cy });
        }
        // Embedded OLE object (Excel worksheet, Word document, ...): render the
        // p:pic snapshot PowerPoint stores alongside the embed reference.
        else if (uri === OLE_GRAPHIC_DATA_URI) {
            const pic = this.findOleSnapshotPic(graphicData);
            if (pic) {
                if (groupTransform) {
                    await this.renderPictureInGroup(ctx, pic, state, groupTransform);
                }
                else {
                    await state.shapeRenderer.renderPicture(ctx, pic);
                }
            }
            else {
                this.logger.warn('OLE object has no picture snapshot; content skipped');
                state.warnings?.push({
                    code: 'unsupported-element',
                    message: 'OLE object has no picture snapshot; content skipped',
                    detail: { uri },
                });
            }
        }
        // Other types
        else if (uri) {
            this.logger.warn('Unhandled graphic type', { uri });
            state.warnings?.push({
                code: 'unsupported-element',
                message: `Unhandled graphic frame type "${uri}"; content skipped`,
                detail: { uri },
            });
        }
    }
    /**
     * Finds the p:pic snapshot of an embedded OLE object (ECMA-376
     * §19.3.2.4). PowerPoint stores the p:oleObj either directly under
     * a:graphicData or wrapped in mc:AlternateContent whose mc:Choice
     * requires VML rendering (Requires="v"); the mc:Fallback branch carries
     * the same p:oleObj with a p:pic snapshot image of the object's last
     * rendered state, which is exactly what PowerPoint's own image export
     * shows. The pic carries its own slide-space xfrm.
     */
    findOleSnapshotPic(graphicData) {
        const candidates = [getXmlChild(graphicData, 'p:oleObj')];
        const alternate = getXmlChild(graphicData, 'mc:AlternateContent');
        if (alternate) {
            // Fallback first: the Choice branch requires capabilities (VML) this
            // renderer does not provide, and its oleObj carries no snapshot.
            for (const branch of ['mc:Fallback', 'mc:Choice']) {
                const branchNode = getXmlChild(alternate, branch);
                if (branchNode)
                    candidates.push(getXmlChild(branchNode, 'p:oleObj'));
            }
        }
        for (const oleObj of candidates) {
            const pic = oleObj ? getXmlChild(oleObj, 'p:pic') : undefined;
            if (pic)
                return pic;
        }
        return undefined;
    }
    /**
     * Renders a chart from a graphic frame.
     */
    async renderChartFrame(ctx, graphicData, bounds, state) {
        // Get the chart reference
        const chartRef = getXmlChild(graphicData, 'c:chart');
        if (!chartRef) {
            this.logger.debug('No chart reference in graphic data');
            return;
        }
        const chartRelId = getXmlAttr(chartRef, 'r:id');
        if (!chartRelId) {
            this.logger.debug('Chart has no relationship ID');
            return;
        }
        try {
            // Resolve the chart path from the source part's relationships. The
            // .rels path is derived generically (dir/_rels/name.rels) so charts
            // referenced from layout/master parts resolve as well as slide ones.
            const lastSlash = state.slidePath.lastIndexOf('/');
            const slideRelsPath = `${state.slidePath.slice(0, lastSlash + 1)}_rels/` +
                `${state.slidePath.slice(lastSlash + 1)}.rels`;
            const chartTarget = await state.parser.getRelationshipTarget(slideRelsPath, chartRelId);
            if (!chartTarget) {
                this.logger.warn('Could not resolve chart relationship', { id: chartRelId });
                state.warnings?.push({
                    code: 'chart-failed',
                    message: `Could not resolve chart relationship "${chartRelId}"; chart skipped`,
                    detail: { relationshipId: chartRelId },
                });
                return;
            }
            const chartPath = state.parser.resolvePath(state.slidePath, chartTarget);
            this.logger.debug('Rendering chart', { path: chartPath });
            // Parse and render the chart
            const chartData = await this.chartParser.parseChart(state.parser, chartPath);
            if (chartData) {
                state.chartRenderer.renderChart(ctx, chartData, bounds);
            }
        }
        catch (error) {
            this.logger.warn('Failed to render chart', {
                error: error instanceof Error ? error.message : String(error),
            });
            state.warnings?.push({
                code: 'chart-failed',
                message: `Failed to render chart: ${error instanceof Error ? error.message : String(error)}`,
            });
        }
    }
    /**
     * Renders a table from a graphic frame.
     */
    renderTableFrame(ctx, graphicData, bounds, state) {
        // Get the table node (a:tbl)
        const tableNode = getXmlChild(graphicData, 'a:tbl');
        if (!tableNode) {
            this.logger.debug('No table node in graphic data');
            return;
        }
        try {
            this.logger.debug('Rendering table', { bounds });
            state.tableRenderer.renderTable(ctx, tableNode, bounds);
        }
        catch (error) {
            this.logger.warn('Failed to render table', {
                error: error instanceof Error ? error.message : String(error),
            });
            state.warnings?.push({
                code: 'other',
                message: `Failed to render table: ${error instanceof Error ? error.message : String(error)}`,
                detail: { elementType: 'a:tbl' },
            });
        }
    }
    /**
     * Renders mc:AlternateContent elements.
     * Used for SmartArt, diagrams, and other content with fallback representations.
     * @param groupTransform Optional group transform when the content is a group child
     */
    async renderAlternateContent(ctx, alternateContentNode, state, groupTransform) {
        // Define the render callback for child elements.
        // When the AlternateContent is a group child, its fallback content lives in
        // the group's child coordinate space, so the group transform must be applied.
        const renderChild = async (childCtx, tagName, node) => {
            // Placeholders never draw from master/layout parts, at any nesting depth
            if (state.skipPlaceholders && isPlaceholderElement(tagName, node)) {
                return;
            }
            switch (tagName) {
                case 'p:sp':
                    if (groupTransform) {
                        this.renderShapeInGroup(childCtx, node, state, groupTransform);
                    }
                    else {
                        await state.shapeRenderer.renderShape(childCtx, node);
                    }
                    break;
                case 'p:cxnSp':
                    if (groupTransform) {
                        this.renderConnectionShapeInGroup(childCtx, node, state, groupTransform);
                    }
                    else {
                        state.shapeRenderer.renderConnectionShape(childCtx, node);
                    }
                    break;
                case 'p:pic':
                    if (groupTransform) {
                        await this.renderPictureInGroup(childCtx, node, state, groupTransform);
                    }
                    else {
                        await state.shapeRenderer.renderPicture(childCtx, node);
                    }
                    break;
                case 'p:grpSp':
                    await this.renderGroupShape(childCtx, node, state, groupTransform);
                    break;
                case 'p:graphicFrame':
                    await this.renderGraphicFrame(childCtx, node, state, groupTransform);
                    break;
                default:
                    this.logger.debug('Unknown child element in AlternateContent', { tagName });
            }
        };
        await this.alternateContentRenderer.renderAlternateContent(ctx, alternateContentNode, renderChild);
    }
}
//# sourceMappingURL=SlideRenderer.js.map