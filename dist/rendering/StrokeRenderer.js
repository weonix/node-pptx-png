/**
 * Renders shape outlines (strokes) to canvas context.
 * Handles line properties including width, color, dash patterns, caps, and joins.
 */
import { getXmlChild, getXmlAttr, hasXmlChild } from '../core/PptxParser.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import { UnitConverter } from '../core/UnitConverter.js';
import { applyPathToContext } from '../geometry/PathBuilder.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Default stroke width in EMU (1 point = 12700 EMU).
 */
export const DEFAULT_STROKE_WIDTH_EMU = 12700;
/**
 * Default stroke color (dark gray) used when no color is specified anywhere.
 */
const DEFAULT_STROKE_COLOR = { r: 64, g: 64, b: 64, a: 255 };
/**
 * Preset dash patterns mapped to canvas dash arrays.
 * Values are relative to stroke width.
 */
const PRESET_DASH_PATTERNS = {
    solid: [],
    dot: [1, 2],
    dash: [4, 3],
    lgDash: [8, 3],
    dashDot: [4, 3, 1, 3],
    lgDashDot: [8, 3, 1, 3],
    lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDash: [3, 1],
    sysDot: [1, 1],
    sysDashDot: [3, 1, 1, 1],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
/**
 * Line end (arrowhead) size multipliers relative to the stroke width.
 * ECMA-376 (ST_LineEndWidth / ST_LineEndLength, 20.1.10.33-34) only defines
 * the sm/med/lg enumeration; the numeric mapping mirrors how Office scales
 * arrowheads with the line width (roughly 2x / 3x / 5x).
 */
const LINE_END_SIZE_MULTIPLIERS = {
    sm: 2,
    med: 3,
    lg: 5,
};
/**
 * Distance of the concave back notch of a stealth arrowhead from the tip,
 * as a fraction of the arrowhead length. Approximates the Office dart shape.
 */
const STEALTH_NOTCH_FACTOR = 0.6;
/**
 * Valid ST_LineEndType decoration values (excluding 'none').
 */
const LINE_END_TYPES = new Set([
    'triangle',
    'stealth',
    'arrow',
    'diamond',
    'oval',
]);
/**
 * Normalizes a vector to unit length; undefined for (near-)zero vectors.
 */
function normalizeVector(dx, dy) {
    const length = Math.hypot(dx, dy);
    if (length < 1e-9)
        return undefined;
    return { x: dx / length, y: dy / length };
}
/**
 * Computes the endpoints and outward tangent directions of an open path from
 * its first and last drawable segments. For bezier segments the tangent runs
 * through the control point adjacent to the endpoint (falling back to the
 * segment chord when the control point is degenerate); arcs use their chord
 * as an approximation. Returns undefined for closed or empty paths, which
 * have no free line ends to decorate.
 * @param path The path to analyze
 * @returns Start/end info, or undefined when the path has no open ends
 */
export function getPathEndPoints(path) {
    let current;
    let start;
    let last;
    for (const segment of path.segments) {
        // Closed paths have no free ends.
        if (segment.type === 'close')
            return undefined;
        const points = segment.points;
        if (segment.type === 'moveTo') {
            current = points?.[0] ?? current;
            continue;
        }
        if (!points || points.length === 0 || !current)
            continue;
        const endPoint = points[points.length - 1];
        let startControl;
        let endControl;
        switch (segment.type) {
            case 'cubicBezierTo':
                // points = [cp1, cp2, end]
                startControl = points[0];
                endControl = points.length >= 3 ? points[1] : points[0];
                break;
            case 'quadBezierTo':
                // points = [cp, end]
                startControl = points[0];
                endControl = points[0];
                break;
            default:
                // lineTo, and chord approximation for arcTo
                startControl = endPoint;
                endControl = current;
                break;
        }
        // Skip zero-length segments (duplicated vertices from freeform/custGeom
        // emitters) so terminal directions come from a real segment.
        const isZeroLength = current.x === endPoint.x &&
            current.y === endPoint.y &&
            startControl.x === endPoint.x &&
            startControl.y === endPoint.y;
        if (!isZeroLength) {
            start ??= { anchor: current, control: startControl, fallback: endPoint };
            last = { from: current, control: endControl, to: endPoint };
        }
        current = endPoint;
    }
    if (!start || !last)
        return undefined;
    // Head direction points backward, away from the first drawable segment.
    const startDirection = normalizeVector(start.anchor.x - start.control.x, start.anchor.y - start.control.y) ??
        normalizeVector(start.anchor.x - start.fallback.x, start.anchor.y - start.fallback.y);
    // Tail direction points forward, out of the last drawable segment.
    const endDirection = normalizeVector(last.to.x - last.control.x, last.to.y - last.control.y) ??
        normalizeVector(last.to.x - last.from.x, last.to.y - last.from.y);
    if (!startDirection || !endDirection)
        return undefined;
    return {
        start: { point: start.anchor, direction: startDirection },
        end: { point: last.to, direction: endDirection },
    };
}
/**
 * Computes the drawable geometry for a line end decoration.
 * Sizes follow Office conventions: width/length are the stroke width times
 * the sm/med/lg multiplier (2/3/5). Triangle and stealth heads sit with the
 * tip on the endpoint and extend back into the line; diamond and oval are
 * centered on the endpoint; the open arrow is two stroked barbs meeting at
 * the endpoint.
 * @param point Path endpoint the decoration is anchored to
 * @param direction Unit vector pointing outward from the path
 * @param end Line end decoration properties
 * @param lineWidthPx Stroke width in pixels
 * @returns The geometry, or undefined for type 'none'
 */
export function computeLineEndGeometry(point, direction, end, lineWidthPx) {
    const length = LINE_END_SIZE_MULTIPLIERS[end.length] * lineWidthPx;
    const halfWidth = (LINE_END_SIZE_MULTIPLIERS[end.width] * lineWidthPx) / 2;
    const { x: dx, y: dy } = direction;
    // Perpendicular unit vector
    const nx = -dy;
    const ny = dx;
    /** A point at (along, across) in the decoration's local frame. */
    const at = (along, across) => ({
        x: point.x + along * dx + across * nx,
        y: point.y + along * dy + across * ny,
    });
    switch (end.type) {
        case 'triangle':
            // Filled triangle: tip on the endpoint, base inside the line.
            return {
                kind: 'polygon',
                points: [at(0, 0), at(-length, halfWidth), at(-length, -halfWidth)],
                inset: length,
            };
        case 'stealth':
            // Dart: triangle with a concave back notch on the line axis.
            return {
                kind: 'polygon',
                points: [
                    at(0, 0),
                    at(-length, halfWidth),
                    at(-length * STEALTH_NOTCH_FACTOR, 0),
                    at(-length, -halfWidth),
                ],
                inset: length * STEALTH_NOTCH_FACTOR,
            };
        case 'diamond':
            // Rhombus centered on the endpoint.
            return {
                kind: 'polygon',
                points: [at(length / 2, 0), at(0, halfWidth), at(-length / 2, 0), at(0, -halfWidth)],
                inset: 0,
            };
        case 'oval':
            // Ellipse centered on the endpoint, long axis along the line.
            return {
                kind: 'ellipse',
                center: { ...point },
                radiusAlong: length / 2,
                radiusAcross: halfWidth,
                rotation: Math.atan2(dy, dx),
                inset: 0,
            };
        case 'arrow':
            // Open arrow: two stroked barbs meeting at the endpoint.
            return {
                kind: 'openArrow',
                tip: at(0, 0),
                barbs: [at(-length, halfWidth), at(-length, -halfWidth)],
                inset: 0,
            };
        case 'none':
        default:
            return undefined;
    }
}
/**
 * Returns a copy of the path with its straight terminal segments pulled back
 * by the given insets so the line does not poke past filled arrowhead tips.
 * Only lineTo terminal segments are shortened; curved/arc ends keep their
 * full extent (acceptable simplification: the arrowhead fill covers the
 * overlap in most cases). Insets are clamped to the terminal segment length,
 * and scaled down proportionally when both ends shorten the same segment.
 * @param path The path to shorten (not mutated)
 * @param headInset Pixels to pull back the start of the path
 * @param tailInset Pixels to pull back the end of the path
 * @returns A new path with adjusted terminal points
 */
export function shortenPathEnds(path, headInset, tailInset) {
    if (headInset <= 0 && tailInset <= 0)
        return path;
    const segments = path.segments.map((segment) => ({
        ...segment,
        ...(segment.points ? { points: segment.points.map((p) => ({ ...p })) } : {}),
    }));
    // Locate the first and last drawable segments and their preceding points.
    let current;
    let head;
    let tail;
    for (const segment of segments) {
        const points = segment.points;
        if (segment.type === 'moveTo') {
            current = points?.[0] ?? current;
            continue;
        }
        if (segment.type === 'close' || !points || points.length === 0 || !current)
            continue;
        head ??= { anchor: current, segment };
        tail = { from: current, segment };
        current = points[points.length - 1];
    }
    // Only straight terminal segments are shortened.
    const headTarget = head?.segment.type === 'lineTo' ? head.segment.points?.[0] : undefined;
    const tailTarget = tail?.segment.type === 'lineTo' ? tail.segment.points?.[0] : undefined;
    const distance = (a, b) => Math.hypot(b.x - a.x, b.y - a.y);
    let effectiveHead = head && headTarget ? Math.min(headInset, distance(head.anchor, headTarget)) : 0;
    let effectiveTail = tail && tailTarget ? Math.min(tailInset, distance(tail.from, tailTarget)) : 0;
    // When both ends shorten the same single segment, keep the total within
    // the segment length so the ends cannot cross.
    if (head && head.segment === tail?.segment && headTarget && tailTarget) {
        const segmentLength = distance(head.anchor, headTarget);
        const total = effectiveHead + effectiveTail;
        if (total > segmentLength && total > 0) {
            const scale = segmentLength / total;
            effectiveHead *= scale;
            effectiveTail *= scale;
        }
    }
    const moveToward = (from, toward, amount) => {
        const dir = normalizeVector(toward.x - from.x, toward.y - from.y);
        if (!dir)
            return;
        from.x += dir.x * amount;
        from.y += dir.y * amount;
    };
    if (head && headTarget && effectiveHead > 0) {
        moveToward(head.anchor, headTarget, effectiveHead);
    }
    if (tail && tailTarget && effectiveTail > 0) {
        moveToward(tailTarget, tail.from, effectiveTail);
    }
    return { ...path, segments };
}
/**
 * Renders strokes for shapes.
 */
export class StrokeRenderer {
    logger;
    colorResolver;
    unitConverter;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'StrokeRenderer');
        this.colorResolver = new ColorResolver(config.theme.colors, config.colorMap);
        this.unitConverter = new UnitConverter();
    }
    /**
     * Renders a stroke to the canvas for the given path.
     * @param ctx Canvas 2D context
     * @param path Path to stroke
     * @param stroke Stroke definition
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     */
    renderStroke(ctx, path, stroke, scaleX, scaleY) {
        // Calculate stroke width in pixels
        const widthPixels = this.unitConverter.emuToPixels(stroke.width) * Math.min(scaleX, scaleY);
        // Skip very thin strokes
        if (widthPixels < 0.1) {
            return;
        }
        const lineWidth = Math.max(widthPixels, 0.5); // Minimum width for visibility
        // Resolve line end decorations (a:headEnd / a:tailEnd) for open paths
        let headGeometry;
        let tailGeometry;
        if (stroke.headEnd || stroke.tailEnd) {
            const ends = getPathEndPoints(path);
            if (ends) {
                if (stroke.headEnd) {
                    headGeometry = computeLineEndGeometry(ends.start.point, ends.start.direction, stroke.headEnd, lineWidth);
                }
                if (stroke.tailEnd) {
                    tailGeometry = computeLineEndGeometry(ends.end.point, ends.end.direction, stroke.tailEnd, lineWidth);
                }
            }
        }
        // Pull straight line ends back so they do not poke past arrowhead tips
        let drawPath = path;
        const headInset = headGeometry?.inset ?? 0;
        const tailInset = tailGeometry?.inset ?? 0;
        if (headInset > 0 || tailInset > 0) {
            drawPath = shortenPathEnds(path, headInset, tailInset);
        }
        ctx.save();
        // Set stroke style
        ctx.strokeStyle = this.colorResolver.rgbaToCss(stroke.color);
        ctx.lineWidth = lineWidth;
        // Set line cap
        ctx.lineCap = this.mapLineCap(stroke.cap ?? 'flat');
        // Set line join
        ctx.lineJoin = this.mapLineJoin(stroke.join ?? 'miter');
        // Set dash pattern if specified
        if (stroke.dashPattern && stroke.dashPattern.length > 0) {
            // Scale dash pattern by stroke width
            const scaledPattern = stroke.dashPattern.map((v) => v * widthPixels);
            ctx.setLineDash(scaledPattern);
        }
        else {
            ctx.setLineDash([]);
        }
        // Apply path and stroke
        applyPathToContext(ctx, drawPath, true);
        ctx.stroke();
        // Draw line end decorations (always solid, never dashed)
        if (headGeometry || tailGeometry) {
            ctx.setLineDash([]);
            if (headGeometry)
                this.drawLineEnd(ctx, headGeometry, stroke.color, lineWidth);
            if (tailGeometry)
                this.drawLineEnd(ctx, tailGeometry, stroke.color, lineWidth);
        }
        ctx.restore();
        this.logger.debug('Rendered stroke', {
            color: this.colorResolver.rgbaToHex(stroke.color),
            width: widthPixels.toFixed(2),
            cap: stroke.cap,
            join: stroke.join,
        });
    }
    /**
     * Parses stroke (outline) properties from a shape properties node.
     * Explicit a:ln properties win; properties the outline does not specify
     * fall back to the referenced style stroke (from a:lnRef), then defaults,
     * per the ECMA-376 style matrix precedence.
     * @param spPr Shape properties node containing a:ln
     * @param styleStroke Stroke resolved from the shape's a:lnRef, if any
     * @returns Parsed stroke or undefined if no stroke should render
     */
    parseStroke(spPr, styleStroke) {
        if (!spPr)
            return styleStroke;
        const ln = getXmlChild(spPr, 'a:ln');
        if (!ln)
            return styleStroke;
        // Explicit no fill on the outline beats any style reference.
        // hasXmlChild: a self-closed <a:noFill/> parses to '' which is falsy.
        if (hasXmlChild(ln, 'a:noFill')) {
            return undefined;
        }
        // Get stroke width
        const w = getXmlAttr(ln, 'w');
        const width = w !== undefined ? parseInt(w, 10) : (styleStroke?.width ?? DEFAULT_STROKE_WIDTH_EMU);
        // Get stroke color
        const color = this.parseStrokeColor(ln) ?? styleStroke?.color ?? { ...DEFAULT_STROKE_COLOR };
        // Get dash pattern
        const dashPattern = this.parseDashPattern(ln) ?? styleStroke?.dashPattern;
        // Get line cap
        const cap = this.parseLineCap(ln) ?? styleStroke?.cap ?? 'flat';
        // Get line join
        const join = this.parseLineJoin(ln) ?? styleStroke?.join ?? 'miter';
        // Get line end decorations (arrowheads); an explicit type="none"
        // suppresses any decoration inherited from the style stroke
        const headEnd = this.resolveLineEnd(ln, 'a:headEnd', styleStroke?.headEnd);
        const tailEnd = this.resolveLineEnd(ln, 'a:tailEnd', styleStroke?.tailEnd);
        return {
            width,
            color,
            dashPattern,
            cap,
            join,
            headEnd,
            tailEnd,
        };
    }
    /**
     * Resolves a theme line style (from lnStyleLst) into a renderable Stroke,
     * substituting phClr placeholders with the a:lnRef child color.
     * @param style The theme line style
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved stroke, or undefined when the style has no line fill
     */
    resolveThemeLineStyle(style, phClr) {
        if (style.fill?.type === 'none') {
            return undefined;
        }
        let color;
        if (style.fill?.type === 'solid') {
            color = this.colorResolver.resolveThemeStyleColor(style.fill.color, phClr);
        }
        else if (style.fill?.type === 'gradient') {
            const firstStop = style.fill.stops[0];
            color = firstStop
                ? this.colorResolver.resolveThemeStyleColor(firstStop.color, phClr)
                : undefined;
        }
        else if (style.fill?.type === 'pattern') {
            color = this.colorResolver.resolveThemeStyleColor(style.fill.foregroundColor, phClr);
        }
        const dashPattern = style.dash !== undefined ? PRESET_DASH_PATTERNS[style.dash] : undefined;
        return {
            width: style.width ?? DEFAULT_STROKE_WIDTH_EMU,
            color: color ?? phClr ?? { ...DEFAULT_STROKE_COLOR },
            dashPattern,
            cap: this.mapOoxmlCap(style.cap),
            join: style.join ?? 'miter',
        };
    }
    /**
     * Maps a raw OOXML cap attribute value to a LineCap.
     */
    mapOoxmlCap(cap) {
        switch (cap) {
            case 'rnd':
                return 'round';
            case 'sq':
                return 'square';
            default:
                return 'flat';
        }
    }
    /**
     * Parses stroke color from outline element.
     * @returns The explicit outline color, or undefined when none is declared
     */
    parseStrokeColor(ln) {
        // Check for solid fill
        const solidFill = getXmlChild(ln, 'a:solidFill');
        if (solidFill) {
            const color = this.colorResolver.resolveColorElement(solidFill);
            if (color)
                return color;
        }
        // Check for gradient fill (use first stop color)
        const gradFill = getXmlChild(ln, 'a:gradFill');
        if (gradFill) {
            const gsLst = getXmlChild(gradFill, 'a:gsLst');
            if (gsLst) {
                const gs = getXmlChild(gsLst, 'a:gs');
                if (gs) {
                    const color = this.colorResolver.resolveColorElement(gs);
                    if (color)
                        return color;
                }
            }
        }
        return undefined;
    }
    /**
     * Parses dash pattern from outline element.
     */
    parseDashPattern(ln) {
        // Check for preset dash
        const prstDash = getXmlChild(ln, 'a:prstDash');
        if (prstDash) {
            const val = getXmlAttr(prstDash, 'val');
            if (val && PRESET_DASH_PATTERNS[val]) {
                return PRESET_DASH_PATTERNS[val];
            }
        }
        // Check for custom dash
        const custDash = getXmlChild(ln, 'a:custDash');
        if (custDash) {
            // Custom dash patterns would be parsed here
            // For now, return undefined to use solid line
            this.logger.debug('Custom dash pattern not fully supported');
        }
        return undefined;
    }
    /**
     * Resolves a line end decoration for one end of the outline, applying the
     * ECMA-376 precedence: an explicit a:headEnd/a:tailEnd wins (including an
     * explicit or defaulted type="none", which clears the decoration), an
     * absent element inherits from the style stroke.
     * @param ln Outline element
     * @param tagName 'a:headEnd' or 'a:tailEnd'
     * @param inherited Decoration inherited from the style stroke, if any
     * @returns The effective decoration, or undefined for none
     */
    resolveLineEnd(ln, tagName, inherited) {
        const parsed = this.parseLineEnd(ln, tagName);
        if (parsed === 'none')
            return undefined;
        return parsed ?? inherited;
    }
    /**
     * Parses an a:headEnd / a:tailEnd child of an outline element.
     * Per ECMA-376 CT_LineEndProperties, type defaults to 'none' and w/len
     * default to 'med'.
     * @returns The decoration, the sentinel 'none' when the element is present
     * but declares no decoration, or undefined when the element is absent
     */
    parseLineEnd(ln, tagName) {
        // hasXmlChild: a self-closed element without attributes parses to a
        // falsy value, but its presence still means "type defaults to none"
        if (!hasXmlChild(ln, tagName))
            return undefined;
        const lineEnd = getXmlChild(ln, tagName);
        const type = lineEnd ? getXmlAttr(lineEnd, 'type') : undefined;
        if (!type || !LINE_END_TYPES.has(type))
            return 'none';
        return {
            type: type,
            width: this.parseLineEndSize(getXmlAttr(lineEnd, 'w')),
            length: this.parseLineEndSize(getXmlAttr(lineEnd, 'len')),
        };
    }
    /**
     * Parses an ST_LineEndWidth / ST_LineEndLength attribute value, defaulting
     * to 'med' per ECMA-376.
     */
    parseLineEndSize(value) {
        return value === 'sm' || value === 'lg' ? value : 'med';
    }
    /**
     * Draws a single line end decoration.
     * Filled shapes (triangle, stealth, diamond, oval) fill with the stroke
     * color; the open arrow strokes its two barbs at the line width with round
     * caps and joins.
     */
    drawLineEnd(ctx, geometry, color, lineWidth) {
        const css = this.colorResolver.rgbaToCss(color);
        switch (geometry.kind) {
            case 'polygon': {
                ctx.beginPath();
                geometry.points.forEach((point, index) => {
                    if (index === 0) {
                        ctx.moveTo(point.x, point.y);
                    }
                    else {
                        ctx.lineTo(point.x, point.y);
                    }
                });
                ctx.closePath();
                ctx.fillStyle = css;
                ctx.fill();
                break;
            }
            case 'ellipse': {
                ctx.beginPath();
                ctx.ellipse(geometry.center.x, geometry.center.y, geometry.radiusAlong, geometry.radiusAcross, geometry.rotation, 0, Math.PI * 2);
                ctx.fillStyle = css;
                ctx.fill();
                break;
            }
            case 'openArrow': {
                const [barb1, barb2] = geometry.barbs;
                ctx.beginPath();
                ctx.moveTo(barb1.x, barb1.y);
                ctx.lineTo(geometry.tip.x, geometry.tip.y);
                ctx.lineTo(barb2.x, barb2.y);
                ctx.strokeStyle = css;
                ctx.lineWidth = lineWidth;
                ctx.lineCap = 'round';
                ctx.lineJoin = 'round';
                ctx.stroke();
                break;
            }
        }
    }
    /**
     * Parses line cap from outline element.
     * @returns The explicit cap, or undefined when the attribute is absent
     */
    parseLineCap(ln) {
        const cap = getXmlAttr(ln, 'cap');
        switch (cap) {
            case 'rnd':
                return 'round';
            case 'sq':
                return 'square';
            case 'flat':
                return 'flat';
            default:
                return undefined;
        }
    }
    /**
     * Parses line join from outline element.
     * @returns The explicit join, or undefined when no join child is present
     */
    parseLineJoin(ln) {
        // Check for explicit join type
        if (getXmlChild(ln, 'a:round')) {
            return 'round';
        }
        if (getXmlChild(ln, 'a:bevel')) {
            return 'bevel';
        }
        if (getXmlChild(ln, 'a:miter')) {
            return 'miter';
        }
        return undefined;
    }
    /**
     * Maps LineCap type to canvas lineCap value.
     */
    mapLineCap(cap) {
        switch (cap) {
            case 'round':
                return 'round';
            case 'square':
                return 'square';
            case 'flat':
            default:
                return 'butt';
        }
    }
    /**
     * Maps LineJoin type to canvas lineJoin value.
     */
    mapLineJoin(join) {
        switch (join) {
            case 'round':
                return 'round';
            case 'bevel':
                return 'bevel';
            case 'miter':
            default:
                return 'miter';
        }
    }
    /**
     * Gets stroke style string for a color (utility method).
     */
    getStrokeStyle(color) {
        return this.colorResolver.rgbaToCss(color);
    }
    /**
     * Converts EMU stroke width to pixels.
     */
    strokeWidthToPixels(widthEmu, scale) {
        return this.unitConverter.emuToPixels(widthEmu) * scale;
    }
}
/**
 * Default stroke renderer factory.
 */
export function createStrokeRenderer(theme, logger) {
    return new StrokeRenderer({ theme, logger });
}
//# sourceMappingURL=StrokeRenderer.js.map