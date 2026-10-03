import { Canvas, type CanvasRenderingContext2D } from 'skia-canvas';
import type { ResolvedTheme, PptxRenderOptions, Rgba } from '../types/index.js';
import type { PptxParser, PptxXmlNode, SlideData } from '../core/PptxParser.js';
import { UnitConverter } from '../core/UnitConverter.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Context for rendering a single slide.
 */
export interface SlideRenderContext {
    /** Canvas to render to */
    canvas: Canvas;
    /** 2D rendering context */
    ctx: CanvasRenderingContext2D;
    /** Resolved theme */
    theme: ResolvedTheme;
    /** Slide width in EMU */
    slideWidthEmu: number;
    /** Slide height in EMU */
    slideHeightEmu: number;
    /** Target width in pixels */
    targetWidth: number;
    /** Target height in pixels */
    targetHeight: number;
    /** Horizontal scale factor (EMU to pixels) */
    scaleX: number;
    /** Vertical scale factor (EMU to pixels) */
    scaleY: number;
    /** Unit converter */
    unitConverter: UnitConverter;
    /** PPTX parser for accessing resources */
    parser: PptxParser;
    /** Slide path for relationship resolution */
    slidePath: string;
    /** Current shape's fill color (for text contrast) */
    shapeFillColor?: Rgba;
    /** Debug mode enabled */
    debugMode: boolean;
}
/**
 * Result of rendering a slide.
 */
export interface SlideRenderOutput {
    /** Rendered image buffer (PNG or JPEG) */
    imageData: Buffer;
    /** Rendered width in pixels */
    width: number;
    /** Rendered height in pixels */
    height: number;
    /** Whether rendering succeeded */
    success: boolean;
    /** Error message if failed */
    error?: string;
}
/**
 * Renders individual slides to images.
 */
export declare class SlideRenderer {
    private readonly logger;
    private readonly theme;
    private readonly options;
    private readonly backgroundRenderer;
    private readonly unitConverter;
    private readonly chartParser;
    private readonly alternateContentRenderer;
    private readonly pngOptimizer;
    private pngOptimizerInitialized;
    /** Table styles parsed once per parser (multi-slide renders reuse it) */
    private readonly tableStylesCache;
    /**
     * Ordered-parse results cached per parser+part path. Master/layout parts
     * are re-rendered for every slide that uses them; the parser already
     * caches the raw XML string, and this cache additionally avoids repeating
     * the preserveOrder parse of identical part content across slides.
     */
    private readonly orderedDocCache;
    constructor(theme: ResolvedTheme, options?: PptxRenderOptions, logger?: ILogger);
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
    renderSlide(parser: PptxParser, slideData: SlideData, slideWidthEmu: number, slideHeightEmu: number, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode, layoutPath?: string, masterPath?: string, warnings?: WarningCollector): Promise<SlideRenderOutput>;
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
    renderSlideToContext(canvas: Canvas, ctx: CanvasRenderingContext2D, parser: PptxParser, slideData: SlideData, slideWidthEmu: number, slideHeightEmu: number, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode, layoutPath?: string, masterPath?: string, warnings?: WarningCollector): Promise<void>;
    /**
     * Builds the per-part render state used by {@link renderShapeTreeOrdered}.
     * The same slide-scoped context (theme, color map, inheritance nodes,
     * slide number) is shared by all parts of one slide render; only the
     * source part path differs so relationships (images, charts, diagrams)
     * resolve against the right .rels file.
     */
    private createRenderState;
    /**
     * Creates a canvas of the given pixel size, applying the `gpu` render
     * option: an explicit boolean sets `canvas.gpu`, while 'auto' leaves the
     * skia-canvas default untouched. The mode in effect is recorded in the
     * debug log. Multi-page exports (PDF) call this once and then draw each
     * slide onto a `canvas.newPage()` context.
     */
    createCanvas(width: number, height: number): Canvas;
    /**
     * Computes the output pixel dimensions these options produce for a slide
     * of the given native EMU size (explicit width/height, or width with
     * height derived from the aspect ratio).
     */
    getTargetDimensions(slideWidthEmu: number, slideHeightEmu: number): {
        width: number;
        height: number;
    };
    /**
     * Calculates target dimensions maintaining aspect ratio.
     */
    private calculateDimensions;
    /**
     * Exports the canvas to an image buffer in the configured format.
     *
     * Lossy formats (jpeg, webp) use the `quality` option (0-1), falling back
     * to the legacy `jpegQuality` (1-100). SVG returns the UTF-8 document
     * bytes. `pngOptimization` only applies to PNG output.
     */
    private exportCanvas;
    /**
     * Resolves the effective lossy-encoding quality (0-1): the `quality`
     * option when set (clamped to [0, 1]), otherwise the legacy `jpegQuality`
     * (1-100) scaled down.
     */
    private resolveQuality;
    /**
     * Converts EMU coordinates to pixel coordinates.
     */
    emuToPixels(context: SlideRenderContext, emuX: number, emuY: number): {
        x: number;
        y: number;
    };
    /**
     * Converts EMU dimensions to pixel dimensions.
     */
    emuToPixelSize(context: SlideRenderContext, emuWidth: number, emuHeight: number): {
        width: number;
        height: number;
    };
    /**
     * Gets the background color for contrast calculations.
     */
    getBackgroundColor(slideNode: PptxXmlNode, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode): Rgba;
    /**
     * Renders all shapes in the shape tree using ordered parsing.
     * Uses the raw XML and parses with preserveOrder to maintain correct z-order.
     * Elements that appear later in the XML are rendered on top.
     */
    private renderShapeTreeOrdered;
    /**
     * Reads a part with preserved element order, caching the parsed result per
     * parser+path. The parser itself caches the raw XML string but re-runs the
     * preserveOrder parse on every readXmlOrdered call; master/layout parts are
     * rendered once per slide that uses them, so this cache makes those
     * repeat renders parse-free. Keyed by parser instance (like the table
     * styles cache) so a new document never sees stale entries.
     */
    private readXmlOrderedCached;
    /**
     * Navigates the ordered XML structure to find the spTree children.
     * Returns the array of children in document order, or undefined if not found.
     *
     * @param orderedSlide Ordered parse of a slide, layout, or master part
     * @param rootTag Root element of the part ('p:sld', 'p:sldLayout', 'p:sldMaster')
     */
    private navigateToSpTree;
    /**
     * Renders a group shape (p:grpSp) element.
     * @param ctx Canvas 2D context
     * @param grpSpNode Group shape XML node
     * @param state Internal render state
     * @param parentGroupTransform Optional parent group transform for nested groups
     */
    private renderGroupShape;
    /**
     * Renders a shape within a group, applying the group transform.
     */
    private renderShapeInGroup;
    /**
     * Renders a connection shape within a group.
     */
    private renderConnectionShapeInGroup;
    /**
     * Renders a picture within a group.
     */
    private renderPictureInGroup;
    /**
     * Renders a graphic frame element (p:graphicFrame).
     * Handles charts, tables, and other embedded content.
     * @param groupTransform Optional group transform when the frame is a group child
     */
    private renderGraphicFrame;
    /**
     * Finds the p:pic snapshot of an embedded OLE object (ECMA-376
     * §19.3.2.4). PowerPoint stores the p:oleObj either directly under
     * a:graphicData or wrapped in mc:AlternateContent whose mc:Choice
     * requires VML rendering (Requires="v"); the mc:Fallback branch carries
     * the same p:oleObj with a p:pic snapshot image of the object's last
     * rendered state, which is exactly what PowerPoint's own image export
     * shows. The pic carries its own slide-space xfrm.
     */
    private findOleSnapshotPic;
    /**
     * Renders a chart from a graphic frame.
     */
    private renderChartFrame;
    /**
     * Renders a table from a graphic frame.
     */
    private renderTableFrame;
    /**
     * Renders mc:AlternateContent elements.
     * Used for SmartArt, diagrams, and other content with fallback representations.
     * @param groupTransform Optional group transform when the content is a group child
     */
    private renderAlternateContent;
}
