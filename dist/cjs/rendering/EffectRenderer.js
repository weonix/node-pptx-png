"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.EffectRenderer = void 0;
exports.resolveAlignmentAnchor = resolveAlignmentAnchor;
exports.computeSilhouetteTransform = computeSilhouetteTransform;
exports.transformRectBounds = transformRectBounds;
exports.createEffectRenderer = createEffectRenderer;
const skia_canvas_1 = require("skia-canvas");
const EffectParser_js_1 = require("../parsers/EffectParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const PathBuilder_js_1 = require("../geometry/PathBuilder.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Effect layers are skipped when their padded pixel area exceeds this
 * multiple of the target canvas area (perf guard: offscreen allocations
 * stay proportional to the slide, not to runaway geometry).
 */
const MAX_EFFECT_AREA_RATIO = 4;
/**
 * skia-canvas maps ctx.filter blur(N) to a gaussian with sigma ≈ N/2
 * (verified empirically; the same semantics as ctx.shadowBlur), so effect
 * blur radii pass through unchanged and silhouette blurs stay consistent
 * with the ctx.shadow* fast path.
 */
const CSS_BLUR_SIGMA_RATIO = 1;
/** Padding factor around blurred layers (≈3 sigma at the ratio above). */
const BLUR_PADDING_FACTOR = 1.5;
/**
 * Resolves an effect alignment value (algn) to its anchor point on the
 * bounds. The anchor is the fixed point of the scale/skew transform.
 * Unknown or missing values fall back to 'b' (bottom center), the
 * ECMA-376 default for outer shadows and reflections.
 * @param alignment The algn attribute value
 * @param bounds The shape's local bounds
 * @returns The anchor point in local pixel space
 */
function resolveAlignmentAnchor(alignment, bounds) {
    const left = bounds.x;
    const top = bounds.y;
    const cx = bounds.x + bounds.width / 2;
    const cy = bounds.y + bounds.height / 2;
    const right = bounds.x + bounds.width;
    const bottom = bounds.y + bounds.height;
    switch (alignment) {
        case 'tl':
            return { x: left, y: top };
        case 't':
            return { x: cx, y: top };
        case 'tr':
            return { x: right, y: top };
        case 'l':
            return { x: left, y: cy };
        case 'ctr':
            return { x: cx, y: cy };
        case 'r':
            return { x: right, y: cy };
        case 'bl':
            return { x: left, y: bottom };
        case 'br':
            return { x: right, y: bottom };
        case 'b':
        default:
            return { x: cx, y: bottom };
    }
}
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
function computeSilhouetteTransform(params, bounds) {
    const a = params.scaleX;
    const b = Math.tan((params.skewY * Math.PI) / 180);
    const c = Math.tan((params.skewX * Math.PI) / 180);
    const d = params.scaleY;
    const anchor = resolveAlignmentAnchor(params.alignment, bounds);
    return {
        a,
        b,
        c,
        d,
        e: params.offsetX + anchor.x - (a * anchor.x + c * anchor.y),
        f: params.offsetY + anchor.y - (b * anchor.x + d * anchor.y),
    };
}
/**
 * Computes the axis-aligned bounding box of a rectangle mapped through an
 * affine transform.
 * @param matrix The transform to apply
 * @param rect The rectangle to transform
 * @returns The bounding box of the transformed corners
 */
function transformRectBounds(matrix, rect) {
    const corners = [
        { x: rect.x, y: rect.y },
        { x: rect.x + rect.width, y: rect.y },
        { x: rect.x + rect.width, y: rect.y + rect.height },
        { x: rect.x, y: rect.y + rect.height },
    ];
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    for (const p of corners) {
        const x = matrix.a * p.x + matrix.c * p.y + matrix.e;
        const y = matrix.b * p.x + matrix.d * p.y + matrix.f;
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
    }
    return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}
/**
 * Renders silhouette-based shape effects (perspective/picture-fill outer
 * shadows, inner shadow, glow, soft edge, reflection) onto a canvas.
 */
class EffectRenderer {
    logger;
    scaleX;
    scaleY;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'EffectRenderer');
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
    }
    /**
     * Decides whether an outer shadow needs the silhouette path instead of
     * the faster ctx.shadow* state: any scale/skew component (perspective
     * presets) cannot be expressed by canvas shadows, and picture fills clip
     * their draw so ctx.shadow* output is discarded.
     * @param shadow The outer shadow effect
     * @param fill The shape's effective fill
     * @returns True when the silhouette path must be used
     */
    needsSilhouetteShadow(shadow, fill) {
        return (shadow.scaleX !== 1 ||
            shadow.scaleY !== 1 ||
            shadow.skewX !== 0 ||
            shadow.skewY !== 0 ||
            fill?.type === 'picture');
    }
    /**
     * Renders an outer shadow by transforming the silhouette (scale/skew
     * about the algn anchor, then the dist/dir offset), tinting it to the
     * shadow color, blurring, and drawing it beneath the shape.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param shadow The outer shadow effect
     */
    renderOuterShadow(ctx, source, shadow) {
        const blurPx = this.effectLengthToPx(shadow.blurRadius);
        const { dx, dy } = (0, EffectParser_js_1.computeShadowOffset)(shadow.distance, shadow.direction);
        const matrix = computeSilhouetteTransform({
            scaleX: shadow.scaleX,
            scaleY: shadow.scaleY,
            skewX: shadow.skewX,
            skewY: shadow.skewY,
            offsetX: (0, UnitConverter_js_1.emuToPixels)(dx) * this.scaleX,
            offsetY: (0, UnitConverter_js_1.emuToPixels)(dy) * this.scaleY,
            alignment: shadow.alignment,
        }, source.bounds);
        const contentBounds = transformRectBounds(matrix, this.silhouetteBounds(source));
        const layer = this.createLayer(contentBounds, this.blurPadding(blurPx), ctx, 'outerShadow');
        if (!layer)
            return;
        layer.ctx.transform(matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f);
        this.paintSilhouette(layer.ctx, source, rgbaToCss(shadow.color));
        const result = blurPx > 0 ? this.blurredCopy(layer.canvas, blurPx) : layer.canvas;
        ctx.drawImage(result, layer.x, layer.y);
    }
    /**
     * Renders a glow: the silhouette tinted to the glow color, blurred by
     * the glow radius, drawn beneath the shape with no offset. The blurred
     * layer is composited twice to strengthen the halo (a single gaussian
     * pass reads too faint next to PowerPoint's glow falloff).
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param glow The glow effect
     */
    renderGlow(ctx, source, glow) {
        const radiusPx = this.effectLengthToPx(glow.radius);
        if (radiusPx <= 0)
            return;
        const layer = this.createLayer(this.silhouetteBounds(source), this.blurPadding(radiusPx), ctx, 'glow');
        if (!layer)
            return;
        this.paintSilhouette(layer.ctx, source, rgbaToCss(glow.color));
        const blurred = this.blurredCopy(layer.canvas, radiusPx);
        ctx.drawImage(blurred, layer.x, layer.y);
        ctx.drawImage(blurred, layer.x, layer.y);
    }
    /**
     * Renders an inner shadow: the inverse of the silhouette offset by
     * dist/dir, tinted, blurred, then clipped back to the silhouette so the
     * shadow darkens only the inside edge. Drawn after the shape's fill.
     * @param ctx Target context, already in the shape's local space
     * @param source The shape's silhouette source
     * @param shadow The inner shadow effect
     */
    renderInnerShadow(ctx, source, shadow) {
        const blurPx = this.effectLengthToPx(shadow.blurRadius);
        const { dx, dy } = (0, EffectParser_js_1.computeShadowOffset)(shadow.distance, shadow.direction);
        const offsetX = (0, UnitConverter_js_1.emuToPixels)(dx) * this.scaleX;
        const offsetY = (0, UnitConverter_js_1.emuToPixels)(dy) * this.scaleY;
        const pad = this.blurPadding(blurPx) + Math.ceil(Math.max(Math.abs(offsetX), Math.abs(offsetY)));
        const silhouette = this.createLayer(source.bounds, pad, ctx, 'innerShadow');
        if (!silhouette)
            return;
        this.paintSilhouette(silhouette.ctx, source, '#000000');
        // Colored inverse of the offset silhouette: flood the layer with the
        // shadow color, then punch out the silhouette shifted by the offset
        const work = this.createLayer(source.bounds, pad, ctx, 'innerShadow');
        if (!work)
            return;
        work.ctx.fillStyle = rgbaToCss(shadow.color);
        work.ctx.fillRect(work.x, work.y, work.canvas.width, work.canvas.height);
        work.ctx.globalCompositeOperation = 'destination-out';
        work.ctx.translate(offsetX, offsetY);
        this.paintSilhouette(work.ctx, source, '#000000');
        // Blur, then keep only the part inside the shape
        const result = this.blurredCopy(work.canvas, blurPx);
        const resultCtx = result.getContext('2d');
        resultCtx.globalCompositeOperation = 'destination-in';
        resultCtx.drawImage(silhouette.canvas, 0, 0);
        ctx.drawImage(result, work.x, work.y);
    }
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
    renderWithSoftEdge(ctx, source, softEdge, drawContent) {
        // Fallback paths draw fill+stroke directly on the target: any caller
        // shadow state would then be applied per pass (double-darkening), so
        // it is cleared first — the rare fallbacks trade the shadow away.
        const drawContentNoShadow = (target) => {
            target.save();
            target.shadowColor = 'transparent';
            target.shadowBlur = 0;
            target.shadowOffsetX = 0;
            target.shadowOffsetY = 0;
            const result = drawContent(target);
            if (result instanceof Promise) {
                return result.finally(() => target.restore());
            }
            target.restore();
            return result;
        };
        const radiusPx = this.effectLengthToPx(softEdge.radius);
        if (radiusPx <= 0) {
            return drawContentNoShadow(ctx);
        }
        const pad = this.blurPadding(radiusPx) + Math.ceil(source.strokeWidthPx ?? 0);
        const layer = this.createLayer(source.bounds, pad, ctx, 'softEdge');
        if (!layer) {
            // Perf guard tripped: draw the content unfaded rather than dropping it
            return drawContentNoShadow(ctx);
        }
        const finish = () => {
            const mask = new skia_canvas_1.Canvas(layer.canvas.width, layer.canvas.height);
            const maskCtx = mask.getContext('2d');
            maskCtx.translate(-layer.x, -layer.y);
            // Filter-on-fill blurs measure sigma = N (unlike drawImage where
            // sigma = N/2, matching shadowBlur) — halve the mask blur so the
            // soft edge width matches the other effect paths.
            maskCtx.filter = `blur(${(radiusPx * CSS_BLUR_SIGMA_RATIO) / 2}px)`;
            this.paintSilhouette(maskCtx, source, '#000000');
            layer.ctx.setTransform(1, 0, 0, 1, 0, 0);
            layer.ctx.globalCompositeOperation = 'destination-in';
            layer.ctx.drawImage(mask, 0, 0);
            ctx.drawImage(layer.canvas, layer.x, layer.y);
        };
        const drawn = drawContent(layer.ctx);
        if (isThenable(drawn)) {
            return drawn.then(finish);
        }
        finish();
    }
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
    renderReflection(ctx, source, reflection, drawContent) {
        const blurPx = this.effectLengthToPx(reflection.blurRadius);
        const { dx, dy } = (0, EffectParser_js_1.computeShadowOffset)(reflection.distance, reflection.direction);
        const matrix = computeSilhouetteTransform({
            scaleX: reflection.scaleX,
            scaleY: reflection.scaleY,
            skewX: reflection.skewX,
            skewY: reflection.skewY,
            offsetX: (0, UnitConverter_js_1.emuToPixels)(dx) * this.scaleX,
            offsetY: (0, UnitConverter_js_1.emuToPixels)(dy) * this.scaleY,
            alignment: reflection.alignment,
        }, source.bounds);
        const contentBounds = transformRectBounds(matrix, this.silhouetteBounds(source));
        const layer = this.createLayer(contentBounds, this.blurPadding(blurPx), ctx, 'reflection');
        if (!layer)
            return;
        layer.ctx.transform(matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f);
        const finish = () => {
            // Blur (blurredCopy also normalizes to a fresh context for the fade)
            const faded = this.blurredCopy(layer.canvas, blurPx);
            const fadedCtx = faded.getContext('2d');
            // Alpha fade along the fade direction across the reflected content
            const gradient = this.createFadeGradient(fadedCtx, reflection, {
                x: contentBounds.x - layer.x,
                y: contentBounds.y - layer.y,
                width: contentBounds.width,
                height: contentBounds.height,
            });
            fadedCtx.globalCompositeOperation = 'destination-in';
            fadedCtx.fillStyle = gradient;
            fadedCtx.fillRect(0, 0, faded.width, faded.height);
            ctx.drawImage(faded, layer.x, layer.y);
        };
        const drawn = drawContent(layer.ctx);
        if (isThenable(drawn)) {
            return drawn.then(finish);
        }
        finish();
    }
    /**
     * Creates the alpha-fade gradient of a reflection: white (only alpha
     * matters under destination-in) from startAlpha at startPosition to
     * endAlpha at endPosition along the fade direction, measured across the
     * reflected content's bounding box in canvas coordinates.
     */
    createFadeGradient(ctx, reflection, contentRect) {
        const rad = (reflection.fadeDirection * Math.PI) / 180;
        const ux = Math.cos(rad);
        const uy = Math.sin(rad);
        // Project the content corners onto the fade axis through the center
        const cx = contentRect.x + contentRect.width / 2;
        const cy = contentRect.y + contentRect.height / 2;
        const halfSpan = (Math.abs(ux) * contentRect.width + Math.abs(uy) * contentRect.height) / 2 || 1;
        const gradient = ctx.createLinearGradient(cx - ux * halfSpan, cy - uy * halfSpan, cx + ux * halfSpan, cy + uy * halfSpan);
        const clamp01 = (v) => Math.min(1, Math.max(0, v));
        const start = clamp01(reflection.startPosition);
        const end = Math.max(clamp01(reflection.endPosition), start);
        const startAlpha = clamp01(reflection.startAlpha);
        const endAlpha = clamp01(reflection.endAlpha);
        if (start > 0)
            gradient.addColorStop(0, `rgba(255, 255, 255, ${startAlpha})`);
        gradient.addColorStop(start, `rgba(255, 255, 255, ${startAlpha})`);
        gradient.addColorStop(end, `rgba(255, 255, 255, ${endAlpha})`);
        if (end < 1)
            gradient.addColorStop(1, `rgba(255, 255, 255, ${endAlpha})`);
        return gradient;
    }
    /**
     * Paints the silhouette into a context: all fillable paths as one
     * compound fill (a single fill avoids double-darkening overlapping
     * subpaths when the color has alpha), or the stroked outline when
     * nothing is fillable (stroke-only shapes).
     */
    paintSilhouette(ctx, source, cssColor) {
        const fillable = source.hasFill === false ? [] : source.paths.filter((p) => p.fill !== false);
        if (fillable.length > 0) {
            ctx.beginPath();
            for (const path of fillable) {
                (0, PathBuilder_js_1.applyPathToContext)(ctx, path, false);
            }
            ctx.fillStyle = cssColor;
            ctx.fill();
            return;
        }
        if (source.strokeWidthPx && source.strokeWidthPx > 0) {
            const strokeable = source.paths.filter((p) => p.stroke !== false);
            if (strokeable.length === 0)
                return;
            ctx.beginPath();
            for (const path of strokeable) {
                (0, PathBuilder_js_1.applyPathToContext)(ctx, path, false);
            }
            ctx.strokeStyle = cssColor;
            ctx.lineWidth = source.strokeWidthPx;
            ctx.stroke();
        }
    }
    /**
     * The local-space bounds the silhouette can reach: the shape bounds,
     * expanded by the stroke width when the silhouette is stroke-based
     * (joins and caps overhang the geometry).
     */
    silhouetteBounds(source) {
        const hasFillable = source.paths.some((p) => p.fill !== false);
        const overhang = hasFillable ? 0 : (source.strokeWidthPx ?? 0);
        return {
            x: source.bounds.x - overhang,
            y: source.bounds.y - overhang,
            width: source.bounds.width + 2 * overhang,
            height: source.bounds.height + 2 * overhang,
        };
    }
    /**
     * Creates an offscreen layer covering contentBounds padded on all sides,
     * with its context translated so local coordinates draw in place.
     * Returns undefined (and logs) when the layer is degenerate or exceeds
     * the perf guard relative to the target canvas.
     */
    createLayer(contentBounds, padPx, target, effectKind) {
        const x = Math.floor(contentBounds.x - padPx);
        const y = Math.floor(contentBounds.y - padPx);
        const width = Math.ceil(contentBounds.width + 2 * padPx) + 1;
        const height = Math.ceil(contentBounds.height + 2 * padPx) + 1;
        if (!(width > 0) || !(height > 0)) {
            return undefined;
        }
        const targetArea = target.canvas.width * target.canvas.height;
        if (width * height > MAX_EFFECT_AREA_RATIO * targetArea) {
            this.logger.debug('Skipping effect: layer exceeds size guard', {
                effectKind,
                width,
                height,
                targetArea,
            });
            return undefined;
        }
        const canvas = new skia_canvas_1.Canvas(width, height);
        const ctx = canvas.getContext('2d');
        ctx.translate(-x, -y);
        return { canvas, ctx, x, y };
    }
    /**
     * Copies a canvas through a gaussian blur (identity transform, so the
     * blur is uniform regardless of any silhouette scale/skew). A radius of
     * 0 still copies, giving callers a fresh context for compositing.
     */
    blurredCopy(source, blurRadiusPx) {
        const out = new skia_canvas_1.Canvas(source.width, source.height);
        const ctx = out.getContext('2d');
        if (blurRadiusPx > 0) {
            ctx.filter = `blur(${blurRadiusPx * CSS_BLUR_SIGMA_RATIO}px)`;
        }
        ctx.drawImage(source, 0, 0);
        return out;
    }
    /**
     * Converts an EMU effect length (blur/glow/soft-edge radius) to pixels,
     * matching the ctx.shadow* fast path's conversion.
     */
    effectLengthToPx(emu) {
        return (0, UnitConverter_js_1.emuToPixels)(emu) * Math.min(this.scaleX, this.scaleY);
    }
    /**
     * Padding needed around a blurred layer so the blur tail is not clipped.
     */
    blurPadding(blurRadiusPx) {
        return Math.ceil(blurRadiusPx * BLUR_PADDING_FACTOR) + 1;
    }
}
exports.EffectRenderer = EffectRenderer;
/**
 * Formats an Rgba as a CSS color (alpha as a 0-1 decimal).
 */
function rgbaToCss(color) {
    if (color.a === 255) {
        return `rgb(${color.r}, ${color.g}, ${color.b})`;
    }
    return `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a / 255).toFixed(3)})`;
}
/**
 * Narrow check for a thenable so synchronous content callbacks complete
 * without deferring compositing to a microtask.
 */
function isThenable(value) {
    return !!value && typeof value.then === 'function';
}
/**
 * Default effect renderer factory.
 */
function createEffectRenderer(scaleX, scaleY, logger) {
    return new EffectRenderer({ scaleX, scaleY, logger });
}
