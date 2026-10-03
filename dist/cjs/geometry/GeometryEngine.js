"use strict";
/**
 * Shared geometry engine that evaluates OOXML shape geometry definitions
 * (preset geometries from presetShapeDefinitions.xml and a:custGeom custom
 * geometries) into renderable pixel-space paths.
 *
 * The drawing language is identical for both: an optional list of adjust
 * value defaults (avLst), an ordered list of guide formulas (gdLst), and a
 * list of paths (pathLst) whose commands may reference literal coordinates
 * or evaluated guides. See ECMA-376 §20.1.9.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ooxmlArcToBeziers = ooxmlArcToBeziers;
exports.buildGeometryPaths = buildGeometryPaths;
exports.computeTextRect = computeTextRect;
const PathBuilder_js_1 = require("./PathBuilder.js");
const GuideEvaluator_js_1 = require("./GuideEvaluator.js");
const TWO_PI = Math.PI * 2;
/** Maximum bezier segments per arc (supports sweeps beyond 360°). */
const MAX_ARC_SEGMENTS = 16;
/**
 * Converts an OOXML polar ellipse angle to the parametric angle of the
 * standard ellipse form (x = rx·cos t, y = ry·sin t).
 *
 * Per ECMA-376 §20.1.9.4, arcTo angles denote the direction of the ray from
 * the ellipse center; the point used is where that ray intersects the
 * ellipse, which differs from the parametric angle for non-circular radii.
 */
function polarToParametric(theta, rx, ry) {
    const t = Math.atan2(rx * Math.sin(theta), ry * Math.cos(theta));
    // atan2 collapses to (-π, π]; restore the original revolution count.
    // The conversion never moves an angle across quadrants, so rounding to
    // the nearest full turn recovers the correct branch.
    return t + Math.round((theta - t) / TWO_PI) * TWO_PI;
}
/**
 * Converts an OOXML arcTo (radii + start/swing angle) starting at the given
 * point into cubic bezier segments.
 *
 * @param startX Current point X (lies on the ellipse at stAng)
 * @param startY Current point Y
 * @param rx Horizontal radius (a:arcTo/@wR, already scaled)
 * @param ry Vertical radius (a:arcTo/@hR, already scaled)
 * @param stAng Start angle in 60000ths of a degree
 * @param swAng Swing angle in 60000ths of a degree (positive = clockwise)
 * @returns Bezier segments and the arc end point
 */
function ooxmlArcToBeziers(startX, startY, rx, ry, stAng, swAng) {
    const start = { x: startX, y: startY };
    if (swAng === 0 || rx <= 0 || ry <= 0) {
        return { segments: [], end: start };
    }
    const theta1 = (0, GuideEvaluator_js_1.ooxmlAngleToRadians)(stAng);
    const theta2 = theta1 + (0, GuideEvaluator_js_1.ooxmlAngleToRadians)(swAng);
    const t1 = polarToParametric(theta1, rx, ry);
    const t2 = polarToParametric(theta2, rx, ry);
    const sweep = t2 - t1;
    // The current point lies at parametric angle t1 on the ellipse.
    const cx = startX - rx * Math.cos(t1);
    const cy = startY - ry * Math.sin(t1);
    const segmentCount = Math.min(Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2))), MAX_ARC_SEGMENTS);
    const delta = sweep / segmentCount;
    // Control point distance for a cubic approximation of an elliptical arc span
    const kappa = (4 / 3) * Math.tan(delta / 4);
    const segments = [];
    let t = t1;
    let current = start;
    for (let i = 0; i < segmentCount; i++) {
        const tNext = t + delta;
        const end = {
            x: cx + rx * Math.cos(tNext),
            y: cy + ry * Math.sin(tNext),
        };
        segments.push({
            cp1: {
                x: current.x - kappa * rx * Math.sin(t),
                y: current.y + kappa * ry * Math.cos(t),
            },
            cp2: {
                x: end.x + kappa * rx * Math.sin(tNext),
                y: end.y - kappa * ry * Math.cos(tNext),
            },
            end,
        });
        t = tNext;
        current = end;
    }
    return { segments, end: current };
}
/**
 * Creates the guide context for a definition: guide space dimensions,
 * adjust value defaults, external adjust overrides, then evaluated guides.
 */
function createDefinitionContext(def, space, adjustValues) {
    const ctx = (0, GuideEvaluator_js_1.createGuideContext)(space.w, space.h, def.av);
    if (adjustValues) {
        for (const [name, value] of adjustValues) {
            ctx.guides.set(name, value);
        }
    }
    if (def.gd) {
        (0, GuideEvaluator_js_1.evaluateGuideList)(def.gd, ctx);
    }
    return ctx;
}
/**
 * Evaluates a geometry definition into renderable pixel-space paths.
 *
 * Guides are evaluated in the guide coordinate space; each path's
 * coordinates are then scaled from its declared coordinate space
 * (a:path/@w/@h, defaulting to the guide space) onto the pixel bounds.
 * OOXML arcs are converted to cubic beziers.
 *
 * @param def Geometry definition (preset or custom)
 * @param bounds Target bounds in pixels
 * @param adjustValues Adjust value overrides (from the shape's a:avLst)
 * @param guideSpace Guide coordinate space; defaults to the pixel bounds
 *   (correct for preset shapes). Pass the shape's EMU extents for custGeom.
 * @returns Paths in pixel space with per-path fill/stroke flags
 * @throws Error on malformed guide formulas or unknown guide references
 */
function buildGeometryPaths(def, bounds, adjustValues, guideSpace) {
    const space = guideSpace ?? { w: bounds.width, h: bounds.height };
    const ctx = createDefinitionContext(def, space, adjustValues);
    const paths = [];
    for (const pathDef of def.paths) {
        const pathSpaceW = pathDef.w ?? space.w;
        const pathSpaceH = pathDef.h ?? space.h;
        const sx = pathSpaceW > 0 ? bounds.width / pathSpaceW : 0;
        const sy = pathSpaceH > 0 ? bounds.height / pathSpaceH : 0;
        const toX = (v) => bounds.x + (0, GuideEvaluator_js_1.resolveGuideValue)(v, ctx) * sx;
        const toY = (v) => bounds.y + (0, GuideEvaluator_js_1.resolveGuideValue)(v, ctx) * sy;
        const builder = new PathBuilder_js_1.PathBuilder();
        let hasSegments = false;
        for (const cmd of pathDef.cmds) {
            switch (cmd[0]) {
                case 'M':
                    builder.moveTo(toX(cmd[1]), toY(cmd[2]));
                    hasSegments = true;
                    break;
                case 'L':
                    builder.lineTo(toX(cmd[1]), toY(cmd[2]));
                    hasSegments = true;
                    break;
                case 'Q':
                    builder.quadBezierTo(toX(cmd[1]), toY(cmd[2]), toX(cmd[3]), toY(cmd[4]));
                    hasSegments = true;
                    break;
                case 'C':
                    builder.cubicBezierTo(toX(cmd[1]), toY(cmd[2]), toX(cmd[3]), toY(cmd[4]), toX(cmd[5]), toY(cmd[6]));
                    hasSegments = true;
                    break;
                case 'A': {
                    const rx = (0, GuideEvaluator_js_1.resolveGuideValue)(cmd[1], ctx) * sx;
                    const ry = (0, GuideEvaluator_js_1.resolveGuideValue)(cmd[2], ctx) * sy;
                    const stAng = (0, GuideEvaluator_js_1.resolveGuideValue)(cmd[3], ctx);
                    const swAng = (0, GuideEvaluator_js_1.resolveGuideValue)(cmd[4], ctx);
                    const current = builder.getCurrentPoint();
                    const { segments } = ooxmlArcToBeziers(current.x, current.y, rx, ry, stAng, swAng);
                    for (const seg of segments) {
                        builder.cubicBezierTo(seg.cp1.x, seg.cp1.y, seg.cp2.x, seg.cp2.y, seg.end.x, seg.end.y);
                        hasSegments = true;
                    }
                    break;
                }
                case 'Z':
                    builder.closePath();
                    break;
            }
        }
        if (!hasSegments)
            continue;
        paths.push(builder.build({
            fill: pathDef.fill !== 'none',
            stroke: pathDef.stroke !== false,
        }));
    }
    return paths;
}
/**
 * Evaluates the text rectangle (a:rect) of a geometry definition into pixel
 * space. Returns undefined when the definition has no rect or the evaluated
 * rect is degenerate.
 *
 * @param def Geometry definition
 * @param bounds Target bounds in pixels
 * @param adjustValues Adjust value overrides
 * @param guideSpace Guide coordinate space (see {@link buildGeometryPaths})
 */
function computeTextRect(def, bounds, adjustValues, guideSpace) {
    if (!def.rect)
        return undefined;
    const space = guideSpace ?? { w: bounds.width, h: bounds.height };
    try {
        const ctx = createDefinitionContext(def, space, adjustValues);
        const sx = space.w > 0 ? bounds.width / space.w : 0;
        const sy = space.h > 0 ? bounds.height / space.h : 0;
        const left = bounds.x + (0, GuideEvaluator_js_1.resolveGuideValue)(def.rect[0], ctx) * sx;
        const top = bounds.y + (0, GuideEvaluator_js_1.resolveGuideValue)(def.rect[1], ctx) * sy;
        const right = bounds.x + (0, GuideEvaluator_js_1.resolveGuideValue)(def.rect[2], ctx) * sx;
        const bottom = bounds.y + (0, GuideEvaluator_js_1.resolveGuideValue)(def.rect[3], ctx) * sy;
        const width = right - left;
        const height = bottom - top;
        if (!isFinite(width) || !isFinite(height) || width <= 0 || height <= 0) {
            return undefined;
        }
        return { x: left, y: top, width, height };
    }
    catch {
        return undefined;
    }
}
