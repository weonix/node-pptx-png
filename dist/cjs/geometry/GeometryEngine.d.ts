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
import type { Path, Point, Rect } from '../types/geometry.js';
/**
 * A coordinate or angle in a geometry definition: either a numeric literal
 * or the name of a guide (adjust value, gdLst guide, or built-in).
 */
export type GeomValue = string | number;
/**
 * A path command in a geometry definition.
 * - `['M', x, y]` — moveTo
 * - `['L', x, y]` — lineTo
 * - `['Q', cx, cy, x, y]` — quadratic bezier
 * - `['C', c1x, c1y, c2x, c2y, x, y]` — cubic bezier
 * - `['A', wR, hR, stAng, swAng]` — OOXML arc (radii + start/swing angles
 *   in 60000ths of a degree)
 * - `['Z']` — close path
 */
export type GeometryPathCommand = ['M' | 'L', GeomValue, GeomValue] | ['Q', GeomValue, GeomValue, GeomValue, GeomValue] | ['C', GeomValue, GeomValue, GeomValue, GeomValue, GeomValue, GeomValue] | ['A', GeomValue, GeomValue, GeomValue, GeomValue] | ['Z'];
/**
 * Fill mode of a geometry path (a:path/@fill). Only 'none' suppresses the
 * fill; the lighten/darken variants currently render as a normal fill.
 */
export type GeometryPathFillMode = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
/**
 * A single path within a geometry definition.
 */
export interface GeometryPathDef {
    /** Width of the path coordinate space (a:path/@w). Defaults to the guide space width. */
    w?: number;
    /** Height of the path coordinate space (a:path/@h). Defaults to the guide space height. */
    h?: number;
    /** Fill mode (a:path/@fill). Defaults to 'norm'. */
    fill?: GeometryPathFillMode;
    /** Whether the path is stroked (a:path/@stroke). Defaults to true. */
    stroke?: boolean;
    /** Path commands in document order. */
    cmds: GeometryPathCommand[];
}
/**
 * A complete shape geometry definition: preset shape or custom geometry.
 */
export interface GeometryDefinition {
    /** Adjust value defaults from avLst: ordered [name, value] pairs. */
    av?: [string, number][];
    /** Guide formulas from gdLst in document order: [name, fmla] pairs. */
    gd?: [string, string][];
    /** Text rectangle [left, top, right, bottom] from a:rect. */
    rect?: [GeomValue, GeomValue, GeomValue, GeomValue];
    /** Paths to render, in document order. */
    paths: GeometryPathDef[];
}
/**
 * The guide coordinate space a geometry definition is evaluated in.
 * For preset shapes this is simply the pixel bounds (formulas are
 * scale-relative). For custom geometry it is the shape's EMU extents,
 * because custGeom paths and guides use absolute EMU coordinates.
 */
export interface GeometrySpace {
    w: number;
    h: number;
}
/**
 * A cubic bezier approximation segment of an arc.
 */
interface ArcBezierSegment {
    cp1: Point;
    cp2: Point;
    end: Point;
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
export declare function ooxmlArcToBeziers(startX: number, startY: number, rx: number, ry: number, stAng: number, swAng: number): {
    segments: ArcBezierSegment[];
    end: Point;
};
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
export declare function buildGeometryPaths(def: GeometryDefinition, bounds: Rect, adjustValues?: ReadonlyMap<string, number>, guideSpace?: GeometrySpace): Path[];
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
export declare function computeTextRect(def: GeometryDefinition, bounds: Rect, adjustValues?: ReadonlyMap<string, number>, guideSpace?: GeometrySpace): Rect | undefined;
export {};
