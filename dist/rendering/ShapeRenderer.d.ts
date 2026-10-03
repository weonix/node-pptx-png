/**
 * Main shape rendering orchestration.
 * Coordinates geometry calculation, transforms, fills, strokes, and text.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { ResolvedTheme } from '../types/theme.js';
import type { Rgba } from '../types/geometry.js';
import type { Fill, Stroke, TextBody } from '../types/elements.js';
import type { PptxParser, PptxXmlNode } from '../core/PptxParser.js';
import { type PixelTransform } from '../geometry/TransformCalculator.js';
import { type GeometryDefinition, type GeometrySpace } from '../geometry/GeometryEngine.js';
import { FillRenderer } from './FillRenderer.js';
import { StrokeRenderer } from './StrokeRenderer.js';
import { TextRenderer } from './TextRenderer.js';
import { TextParser } from '../parsers/TextParser.js';
import { EffectParser, type OuterShadowEffect, type ParsedEffectList } from '../parsers/EffectParser.js';
import { EffectRenderer } from './EffectRenderer.js';
import { ImageRenderer } from './ImageRenderer.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Configuration for ShapeRenderer.
 */
export interface ShapeRendererConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Horizontal scale factor (EMU to pixels) */
    scaleX: number;
    /** Vertical scale factor (EMU to pixels) */
    scaleY: number;
    /** PPTX parser for accessing media (required for picture shapes) */
    parser?: PptxParser;
    /** Source file path for relationship resolution (e.g., ppt/slides/slide1.xml) */
    sourcePath?: string;
    /** Slide layout node for placeholder resolution */
    layoutNode?: PptxXmlNode;
    /** Slide master node for placeholder resolution fallback */
    masterNode?: PptxXmlNode;
    /**
     * Presentation root node (p:presentation) whose p:defaultTextStyle forms
     * the lowest-precedence layer of the placeholder text-style chain
     */
    presentationNode?: PptxXmlNode;
    /** Slide node (p:sld) for color map override resolution */
    slideNode?: PptxXmlNode;
    /**
     * 1-based number of the slide being rendered; forwarded to the text
     * parser so a:fld type="slidenum" fields resolve to the actual number
     */
    slideNumber?: number;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector for fidelity events (per-slide) */
    warnings?: WarningCollector;
}
/**
 * Shape style resolved from a p:style element (ECMA-376 §19.3.1.46).
 * Each ref indexes into the theme's style matrix with the ref's child color
 * substituted for phClr placeholders.
 */
export interface ParsedShapeStyle {
    /** Fill resolved from a:fillRef (undefined when idx is 0/1000 or absent) */
    fill?: Fill;
    /** Stroke resolved from a:lnRef */
    stroke?: Stroke;
    /** Default text color resolved from the a:fontRef child color */
    fontColor?: Rgba;
    /** Font collection referenced by a:fontRef (major/minor/none) */
    fontRefIdx?: string;
    /** Outer shadow resolved from a:effectRef into the theme effectStyleLst */
    outerShadow?: OuterShadowEffect;
}
/**
 * Custom geometry parsed from a:custGeom, ready for evaluation.
 */
export interface ParsedCustomGeometry {
    /** The geometry definition (guides, paths, text rect) */
    def: GeometryDefinition;
    /**
     * Guide coordinate space (the shape's EMU extents). Undefined when the
     * shape has no a:xfrm extents; the pixel bounds are used instead.
     */
    space?: GeometrySpace;
}
/**
 * Parsed shape data ready for rendering.
 */
export interface ParsedShape {
    /** Shape ID */
    id: string;
    /** Shape name */
    name?: string;
    /** Transform in pixels */
    transform: PixelTransform;
    /** Preset geometry name or 'custGeom' for custom geometry */
    geometryType: string;
    /** Custom geometry definition when geometryType is 'custGeom' */
    customGeometry?: ParsedCustomGeometry;
    /** Adjustment values for parameterized shapes */
    adjustValues?: Map<string, number>;
    /** Fill definition */
    fill?: Fill;
    /** Stroke definition */
    stroke?: Stroke;
    /** Text body */
    textBody?: TextBody;
    /** Default text color from the shape's style fontRef (p:style/a:fontRef) */
    styleFontColor?: Rgba;
    /** Effects (a:effectLst or a:effectRef) */
    effects?: ParsedEffectList;
    /** Whether to render as hidden */
    hidden: boolean;
}
/**
 * Renders shapes to canvas.
 */
export declare class ShapeRenderer {
    private readonly logger;
    private readonly theme;
    private readonly scaleX;
    private readonly scaleY;
    private readonly parser?;
    private readonly sourcePath?;
    private readonly layoutNode?;
    private readonly masterNode?;
    /** p:defaultTextStyle extracted from the presentation node, if provided */
    private readonly defaultTextStyle?;
    private readonly colorMap;
    private readonly colorResolver;
    private readonly transformCalculator;
    private readonly geometryCalculator;
    private readonly placeholderResolver;
    /** Fill renderer for external use */
    readonly fillRenderer: FillRenderer;
    /** Stroke renderer for external use */
    readonly strokeRenderer: StrokeRenderer;
    /** Text renderer for external use */
    readonly textRenderer: TextRenderer;
    /** Text parser for external use */
    readonly textParser: TextParser;
    /** Effect parser for external use */
    readonly effectParser: EffectParser;
    /** Silhouette-based effect renderer (perspective shadows, glow, ...) */
    readonly effectRenderer: EffectRenderer;
    private imageRenderer;
    private readonly warnings;
    constructor(config: ShapeRendererConfig);
    /**
     * Renders a shape element to the canvas.
     * @param ctx Canvas 2D context
     * @param spNode Shape XML node (p:sp)
     */
    renderShape(ctx: CanvasRenderingContext2D, spNode: PptxXmlNode): Promise<void>;
    /**
     * Renders a shape's effects and its fill/stroke passes in z-order:
     * reflection, outer shadow, glow beneath the shape; then the shape body
     * (through the soft-edge mask when present); then the inner shadow on
     * top of the fill. Plain offset shadows keep the fast ctx.shadow* path;
     * perspective (sx/sy/kx/ky) and picture-fill shadows go through the
     * silhouette-based EffectRenderer.
     *
     * Completes synchronously unless an async picture fill runs
     * (allowAsyncFills with an initialized ImageRenderer), so the
     * synchronous group-rendering path keeps its draw ordering.
     *
     * @param ctx Canvas context, already in the shape's local space
     * @param paths Geometry paths in local pixel space
     * @param bounds Local-space bounds of the shape
     * @param fill Effective fill
     * @param stroke Effective stroke
     * @param effects Effective effects
     * @param allowAsyncFills Whether picture fills may render asynchronously
     * @returns A promise only when an async picture fill is rendering
     */
    private renderEffectsAndBody;
    /**
     * Renders the fill/stroke passes for each path. A plain outer shadow is
     * applied to exactly one pass — the first fill-enabled path, or the
     * first stroked path when nothing fills (stroke-only shapes still cast
     * shadows). The first path of a geometry carries the silhouette;
     * shadowing every subpath would double-darken overlaps.
     * @returns A promise only when an async picture fill is rendering
     */
    private renderBodyPasses;
    /**
     * Async variant of the body passes for picture fills. Picture-fill outer
     * shadows always take the silhouette path (ctx.shadow* output would be
     * clipped away with the fill), so no shadow state is set here; the
     * shadow parameter guards the stroke-only case for completeness.
     */
    private renderBodyPassesAsync;
    /**
     * Builds the silhouette source EffectRenderer works from: the shape's
     * paths, bounds, and stroke width (used for stroke-only silhouettes and
     * layer padding).
     */
    private createSilhouetteSource;
    /**
     * Builds the content callback effect layers redraw the shape through
     * (soft edge and reflection): the plain fill/stroke passes without any
     * shadow state. Asynchronous only for picture fills when allowed.
     */
    private makeFillStrokeCallback;
    /**
     * Creates the renderable paths for a shape's geometry.
     * Unknown preset geometries and unparsable custom geometries fall back to
     * a rectangle (with a logged warning) so fill, stroke, and text still
     * render instead of the shape vanishing.
     *
     * @returns Paths in pixel space; empty only for degenerate bounds
     */
    private createShapePaths;
    /**
     * Computes the text layout bounds for a shape, honoring the geometry's
     * text rectangle (custom geometry a:rect included).
     */
    private getShapeTextBounds;
    /**
     * Parses a shape XML node into renderable data.
     * @param spNode Shape XML node (p:sp)
     */
    parseShape(spNode: PptxXmlNode): ParsedShape | undefined;
    /**
     * Gets a child element, preserving the presence of empty elements. The XML
     * parser yields self-closed childless elements (e.g. the `<p:spPr/>` and
     * `<p:ph/>` placeholder shapes routinely carry) as empty strings, which
     * must not be confused with a missing element.
     */
    private getPresentChild;
    /**
     * Gets a shape's p:spPr node, treating an empty element as an empty node
     * (placeholder shapes get their transform from the layout).
     */
    private getSpPrNode;
    /**
     * Resolves the inherited list-style chain for a placeholder shape, ordered
     * lowest precedence first (presentation defaultTextStyle, master txStyles
     * bucket, master placeholder lstStyle, then the matching layout
     * placeholder's lstStyle). Returns undefined when nothing is inherited.
     */
    private getInheritedListStyles;
    /**
     * Inherited list styles for NON-placeholder shapes (plain text boxes):
     * only the presentation part's p:defaultTextStyle applies (ECMA-376
     * §19.2.1.12) — master txStyles buckets and layout lstStyles are
     * placeholder-inheritance layers and never reach a text box.
     */
    private getTextBoxListStyles;
    /**
     * Parses and resolves a shape's p:style element (a:fillRef, a:lnRef,
     * a:fontRef) against the theme's style matrix. Per ECMA-376 the style
     * supplies defaults that explicit spPr formatting overrides.
     * @param shapeNode The shape node (p:sp / p:cxnSp / p:pic)
     * @returns The resolved style, or undefined when the shape has no p:style
     */
    parseShapeStyle(shapeNode: PptxXmlNode): ParsedShapeStyle | undefined;
    /**
     * Resolves an a:fillRef into a fill using the theme style matrix.
     * Per ECMA-376 §20.1.4.2.10: idx 0 or 1000 means no fill, 1-999 indexes
     * fillStyleLst (1-based), and 1001+ indexes bgFillStyleLst (1001-based).
     * The ref's child color substitutes phClr placeholders.
     */
    private resolveStyleFill;
    /**
     * Resolves an a:lnRef into a stroke using the theme lnStyleLst
     * (ECMA-376 §20.1.4.2.19; 1-based index, 0 means no line).
     * The ref's child color substitutes phClr placeholders.
     */
    private resolveStyleStroke;
    /**
     * Resolves an a:effectRef into an outer shadow using the theme
     * effectStyleLst (ECMA-376 §20.1.4.2.7; 1-based index, 0 means no
     * effect). The ref's child color substitutes phClr placeholders in the
     * shadow color. Only outer shadows are resolved; other themed effects
     * are not rendered.
     */
    private resolveStyleEffect;
    /**
     * Resolves the effective effects for a shape. Per ECMA-376 an explicit
     * spPr a:effectLst (even an empty one) overrides the style's a:effectRef.
     */
    private resolveShapeEffects;
    /**
     * Enables canvas shadow state for a plain outer shadow (no scale/skew;
     * perspective shadows go through EffectRenderer instead). Offsets and
     * blur are converted EMU -> pixels at the renderer's scale. Canvas
     * shadow offsets live in device space (unaffected by the CTM), which
     * matches rotWithShape="0"; rotated shapes with rotWithShape="1" keep a
     * device-space shadow — an accepted approximation.
     */
    private setOuterShadow;
    /**
     * Resets canvas shadow state so later passes (strokes, text) do not
     * inherit the shadow.
     */
    private clearShadow;
    /**
     * Parses geometry information from shape properties.
     */
    private parseGeometry;
    /**
     * Reads the shape's EMU extents from a:xfrm to use as the custom-geometry
     * guide coordinate space (custGeom coordinates are absolute EMU values).
     */
    private getGuideSpace;
    /**
     * Parses adjustment values from preset geometry.
     */
    private parseAdjustValues;
    /**
     * Looks up a placeholder's transform from the slide layout or master.
     * @param placeholderType The placeholder type (title, body, etc.)
     * @param placeholderIdx The placeholder index
     * @returns The parsed transform or undefined if not found
     */
    private getPlaceholderTransform;
    /**
     * Creates the paths for a connector's geometry (straightConnector1,
     * bentConnector2..5, curvedConnector2..5, ...), falling back to a
     * straight line for unknown geometry.
     */
    private createConnectorPaths;
    /**
     * Renders a connection shape (connector line).
     * @param ctx Canvas 2D context
     * @param cxnSpNode Connection shape XML node (p:cxnSp)
     */
    renderConnectionShape(ctx: CanvasRenderingContext2D, cxnSpNode: PptxXmlNode): void;
    /**
     * Renders a connector's stroke passes with its effects: silhouette
     * outer shadow for perspective parameters (else the fast ctx.shadow*
     * path on the first stroked path only), and glow beneath the strokes.
     * Inner shadow, soft edge, and reflection are not applied to connectors.
     */
    private renderConnectorPasses;
    /**
     * Gets the image renderer for external use.
     * Note: This is a getter because imageRenderer can be null and may be recreated.
     */
    getImageRenderer(): ImageRenderer | null;
    /**
     * Sets a new source path for image rendering.
     * Creates a new ImageRenderer for the given path.
     */
    setSourcePath(sourcePath: string): void;
    /**
     * Renders a picture element (p:pic) to the canvas.
     * @param ctx Canvas 2D context
     * @param picNode Picture XML node (p:pic)
     */
    renderPicture(ctx: CanvasRenderingContext2D, picNode: PptxXmlNode): Promise<void>;
    /**
     * Resolves the outer shadow to apply around a picture draw. Tiled blip
     * fills draw many tiles inside a clip — a per-tile shadow would darken
     * seams — so tiled pictures render shadowless until ImageRenderer
     * supports silhouette shadows.
     */
    private getPictureOuterShadow;
    /**
     * Parses a picture element to get its picture data.
     * Used for extracting picture information without rendering.
     */
    parsePicture(picNode: PptxXmlNode): {
        id: string;
        name?: string;
        transform?: PixelTransform;
        blipRelId?: string;
        effects?: ParsedEffectList;
    } | undefined;
    /**
     * Renders a shape element with a pre-computed pixel transform.
     * Used for rendering shapes within groups where the transform has already been calculated.
     * @param ctx Canvas 2D context
     * @param spNode Shape XML node (p:sp)
     * @param pixelTransform Pre-computed pixel transform
     */
    renderShapeWithTransform(ctx: CanvasRenderingContext2D, spNode: PptxXmlNode, pixelTransform: PixelTransform): void;
    /**
     * Renders a connection shape with a pre-computed pixel transform.
     * Used for rendering connection shapes within groups.
     * @param ctx Canvas 2D context
     * @param cxnSpNode Connection shape XML node (p:cxnSp)
     * @param pixelTransform Pre-computed pixel transform
     */
    renderConnectionShapeWithTransform(ctx: CanvasRenderingContext2D, cxnSpNode: PptxXmlNode, pixelTransform: PixelTransform): void;
    /**
     * Renders a picture element with a pre-computed pixel transform.
     * Used for rendering pictures within groups.
     * @param ctx Canvas 2D context
     * @param picNode Picture XML node (p:pic)
     * @param pixelTransform Pre-computed pixel transform
     */
    renderPictureWithTransform(ctx: CanvasRenderingContext2D, picNode: PptxXmlNode, pixelTransform: PixelTransform): Promise<void>;
}
/**
 * Default shape renderer factory.
 */
export declare function createShapeRenderer(theme: ResolvedTheme, scaleX: number, scaleY: number, logger?: ILogger, parser?: PptxParser, sourcePath?: string): ShapeRenderer;
//# sourceMappingURL=ShapeRenderer.d.ts.map