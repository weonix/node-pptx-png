/**
 * Silhouette-based effect rendering for DrawingML shape effects.
 *
 * Builds the shape's silhouette (its geometry paths painted opaque) on an
 * offscreen canvas, then transforms, tints, blurs, and composites it to
 * realize effects the canvas shadow state cannot express: perspective
 * outer shadows (sx/sy/kx/ky, ECMA-376 §20.1.8.45), shadows for picture
 * fills (whose clip discards ctx.shadow* output), inner shadows
 * (§20.1.8.40), glow (§20.1.8.32), soft edges (§20.1.8.53), and
 * reflections (§20.1.8.50). Plain offset shadows keep the faster
 * ctx.shadow* path in ShapeRenderer.
 *
 * All rendering happens in the shape's local pixel space (inside the
 * shape's canvas transform), so effects rotate with the shape — an
 * accepted approximation of the rotWithShape flag.
 */
import { type CanvasRenderingContext2D } from 'skia-canvas';
import type { Path, Point, Rect, Transform2D } from '../types/geometry.js';
import type { Fill } from '../types/elements.js';
import { type GlowEffect, type InnerShadowEffect, type OuterShadowEffect, type ReflectionEffect, type SoftEdgeEffect } from '../parsers/EffectParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * The silhouette source of a shape: its geometry paths in local pixel
 * space plus the local bounds they were built for.
 */
export interface SilhouetteSource {
    /** Geometry paths in local pixel space */
    paths: Path[];
    /** Local-space bounds the paths were built for (x/y usually 0) */
    bounds: Rect;
    /**
     * Stroke width in pixels. Used to build the silhouette from the stroked
     * outline when no path is fillable (stroke-only shapes still cast
     * shadows and glows) and to pad effect layers for stroke overhang.
     */
    strokeWidthPx?: number;
    /**
     * Whether the shape's effective fill paints anything. When false (e.g.
     * a:noFill with a stroke), the silhouette comes from the stroked outline
     * only — a hollow shape must not cast a solid shadow.
     */
    hasFill?: boolean;
}
/**
 * Callback that paints the shape's own content (fill and stroke passes)
 * into a context. May be asynchronous (picture fills); EffectRenderer
 * completes synchronously whenever the callback does, so synchronous
 * callers keep their draw ordering.
 */
export type ShapeContentCallback = (ctx: CanvasRenderingContext2D) => void | Promise<void>;
/**
 * Parameters of the silhouette transform: scale and skew about an
 * alignment anchor, then an offset.
 */
export interface SilhouetteTransformParams {
    /** Horizontal scale factor (1 = 100%) */
    scaleX: number;
    /** Vertical scale factor (1 = 100%; negative flips) */
    scaleY: number;
    /** Horizontal skew angle in degrees (x' += tan(kx)·y) */
    skewX: number;
    /** Vertical skew angle in degrees (y' += tan(ky)·x) */
    skewY: number;
    /** Horizontal offset in pixels, applied after scale/skew */
    offsetX: number;
    /** Vertical offset in pixels, applied after scale/skew */
    offsetY: number;
    /** Alignment anchor (tl/t/tr/l/ctr/r/bl/b/br); default b per ECMA-376 */
    alignment?: string;
}
/**
 * Resolves an effect alignment value (algn) to its anchor point on the
 * bounds. The anchor is the fixed point of the scale/skew transform.
 * Unknown or missing values fall back to 'b' (bottom center), the
 * ECMA-376 default for outer shadows and reflections.
 * @param alignment The algn attribute value
 * @param bounds The shape's local bounds
 * @returns The anchor point in local pixel space
 */
export declare function resolveAlignmentAnchor(alignment: string | undefined, bounds: Rect): Point;
/**
 * Computes the affine transform of a shadow/reflection silhouette:
 * scale (sx/sy) and skew (kx/ky) about the alignment anchor, then the
 * dist/dir offset. The skew convention follows the MS-ODRAW perspective
 * matrix DrawingML maps onto: x' = sx·x + tan(kx)·y, y' = tan(ky)·x + sy·y.
 * The anchor point is invariant (up to the offset).
 * @param params Scale/skew/offset/alignment parameters
 * @param bounds The shape's local bounds (anchor reference)
 * @returns The composed transform matrix
 */
export declare function computeSilhouetteTransform(params: SilhouetteTransformParams, bounds: Rect): Transform2D;
/**
 * Computes the axis-aligned bounding box of a rectangle mapped through an
 * affine transform.
 * @param matrix The transform to apply
 * @param rect The rectangle to transform
 * @returns The bounding box of the transformed corners
 */
export declare function transformRectBounds(matrix: Transform2D, rect: Rect): Rect;
/**
 * Configuration for EffectRenderer.
 */
export interface EffectRendererConfig {
    /** Horizontal scale factor (EMU to pixels) */
    scaleX: number;
    /** Vertical scale factor (EMU to pixels) */
    scaleY: number;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Renders silhouette-based shape effects (perspective/picture-fill outer
 * shadows, inner shadow, glow, soft edge, reflection) onto a canvas.
 */
export declare class EffectRenderer {
    private readonly logger;
    private readonly scaleX;
    private readonly scaleY;
    constructor(config: EffectRendererConfig);
    /**
     * Decides whether an outer shadow needs the silhouette path instead of
     * the faster ctx.shadow* state: any scale/skew component (perspective
     * presets) cannot be expressed by canvas shadows, and picture fills clip
     * their draw so ctx.shadow* output is discarded.
     * @param shadow The outer shadow effect
     * @param fill The shape's effective fill
     * @returns True when the silhouette path must be used
     */
    needsSilhouetteShadow(shadow: OuterShadowEffect, fill?: Fill): boolean;
    /**
     * Renders an outer shadow by transforming the silhouette (scale/skew
     * about the algn anchor, then the dist/dir offset), tinting it to the
     * shadow color, blurring, and drawing it beneath the shape.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param shadow The outer shadow effect
     */
    renderOuterShadow(ctx: CanvasRenderingContext2D, source: SilhouetteSource, shadow: OuterShadowEffect): void;
    /**
     * Renders a glow: the silhouette tinted to the glow color, blurred by
     * the glow radius, drawn beneath the shape with no offset. The blurred
     * layer is composited twice to strengthen the halo (a single gaussian
     * pass reads too faint next to PowerPoint's glow falloff).
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param glow The glow effect
     */
    renderGlow(ctx: CanvasRenderingContext2D, source: SilhouetteSource, glow: GlowEffect): void;
    /**
     * Renders an inner shadow: the inverse of the silhouette offset by
     * dist/dir, tinted, blurred, then clipped back to the silhouette so the
     * shadow darkens only the inside edge. Drawn after the shape's fill.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param shadow The inner shadow effect
     */
    renderInnerShadow(ctx: CanvasRenderingContext2D, source: SilhouetteSource, shadow: InnerShadowEffect): void;
    /**
     * Renders the shape's content with a soft edge: the content is painted
     * onto an offscreen layer, then masked (destination-in) by the
     * silhouette blurred by the soft-edge radius, fading the edges out.
     * Completes synchronously when the content callback is synchronous.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param softEdge The soft edge effect
     * @param drawContent Paints the shape's fill/stroke passes
     * @returns A promise only when drawContent returned one
     */
    renderWithSoftEdge(ctx: CanvasRenderingContext2D, source: SilhouetteSource, softEdge: SoftEdgeEffect, drawContent: ShapeContentCallback): void | Promise<void>;
    /**
     * Renders a reflection: the shape's content redrawn through the
     * reflection transform (typically a vertical flip about the bottom
     * anchor via sy=-100%), blurred, and faded with a linear alpha gradient
     * (stA/endA over stPos/endPos along fadeDir). Drawn beneath the shape.
     * Completes synchronously when the content callback is synchronous.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param reflection The reflection effect
     * @param drawContent Paints the shape's fill/stroke passes
     * @returns A promise only when drawContent returned one
     */
    renderReflection(ctx: CanvasRenderingContext2D, source: SilhouetteSource, reflection: ReflectionEffect, drawContent: ShapeContentCallback): void | Promise<void>;
    /**
     * Creates the alpha-fade gradient of a reflection: white (only alpha
     * matters under destination-in) from startAlpha at startPosition to
     * endAlpha at endPosition along the fade direction, measured across the
     * reflected content's bounding box in canvas coordinates.
     */
    private createFadeGradient;
    /**
     * Paints the silhouette into a context: all fillable paths as one
     * compound fill (a single fill avoids double-darkening overlapping
     * subpaths when the color has alpha), or the stroked outline when
     * nothing is fillable (stroke-only shapes).
     */
    private paintSilhouette;
    /**
     * The local-space bounds the silhouette can reach: the shape bounds,
     * expanded by the stroke width when the silhouette is stroke-based
     * (joins and caps overhang the geometry).
     */
    private silhouetteBounds;
    /**
     * Creates an offscreen layer covering contentBounds padded on all sides,
     * with its context translated so local coordinates draw in place.
     * Returns undefined (and logs) when the layer is degenerate or exceeds
     * the perf guard relative to the target canvas.
     */
    private createLayer;
    /**
     * Copies a canvas through a gaussian blur (identity transform, so the
     * blur is uniform regardless of any silhouette scale/skew). A radius of
     * 0 still copies, giving callers a fresh context for compositing.
     */
    private blurredCopy;
    /**
     * Converts an EMU effect length (blur/glow/soft-edge radius) to pixels,
     * matching the ctx.shadow* fast path's conversion.
     */
    private effectLengthToPx;
    /**
     * Padding needed around a blurred layer so the blur tail is not clipped.
     */
    private blurPadding;
}
/**
 * Default effect renderer factory.
 */
export declare function createEffectRenderer(scaleX: number, scaleY: number, logger?: ILogger): EffectRenderer;
//# sourceMappingURL=EffectRenderer.d.ts.map