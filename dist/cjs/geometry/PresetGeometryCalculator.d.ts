/**
 * Generates paths for preset shapes defined in ECMA-376.
 *
 * A small set of common shapes is hand-coded (kept as overrides so existing
 * visual baselines stay stable); every other preset geometry is evaluated
 * from the generated ECMA-376 definition table (presetShapeDefinitions.ts)
 * through the guide-formula interpreter in GeometryEngine.
 */
import type { Path, Rect } from '../types/geometry.js';
/**
 * Calculator for preset geometry shapes.
 */
export declare class PresetGeometryCalculator {
    /**
     * List of hand-coded preset geometry names. These take precedence over
     * the generated ECMA-376 definitions to keep visual baselines stable.
     */
    static readonly SUPPORTED_SHAPES: readonly ["rect", "roundRect", "ellipse", "triangle", "rtTriangle", "diamond", "parallelogram", "trapezoid", "pentagon", "hexagon", "octagon", "line", "rightArrow", "leftArrow", "upArrow", "downArrow", "star5", "plus", "heart", "wedgeRectCallout", "flowChartProcess", "flowChartDecision", "flowChartTerminator"];
    /**
     * Checks if a preset geometry is supported (hand-coded or via the
     * generated ECMA-376 definition table).
     */
    isSupported(presetName: string): boolean;
    /**
     * Names handled by the hand-coded switch in createHandCodedPath.
     * Must stay in sync with its cases ('oval', 'chevron' and 'homePlate'
     * are handled there without appearing in SUPPORTED_SHAPES).
     */
    private static readonly HAND_CODED_SHAPES;
    /**
     * Checks if a preset geometry has a hand-coded implementation.
     */
    private isHandCoded;
    /**
     * Creates all paths for the specified preset geometry, preserving
     * per-path fill/stroke flags for multi-path presets (e.g. stroke-only
     * outline paths). Hand-coded shapes yield a single path.
     *
     * @param presetName OpenXML preset geometry name
     * @param bounds Bounding rectangle in pixels
     * @param adjustValues Optional adjustment values for parameterized shapes
     * @returns Paths for rendering, or undefined if the geometry is unknown
     *   or the bounds are degenerate
     */
    createPaths(presetName: string, bounds: Rect, adjustValues?: Map<string, number>): Path[] | undefined;
    /**
     * Evaluates a preset from the generated ECMA-376 definition table.
     */
    private createDefinitionPaths;
    /**
     * Creates a path for the specified preset geometry.
     * Multi-path presets are merged into a single path (fill/stroke flags
     * OR-ed); use {@link createPaths} to preserve per-path flags.
     *
     * @param presetName OpenXML preset geometry name
     * @param bounds Bounding rectangle in pixels
     * @param adjustValues Optional adjustment values for parameterized shapes
     * @returns Path for rendering, or undefined if geometry is not supported
     */
    createPath(presetName: string, bounds: Rect, adjustValues?: Map<string, number>): Path | undefined;
    /**
     * Creates a path for a hand-coded preset geometry.
     */
    private createHandCodedPath;
    /**
     * Resolves adjustment values, using defaults for missing values.
     */
    private resolveAdjustments;
    /**
     * Gets an adjustment value with fallback.
     */
    private getAdj;
    private createRectangle;
    private createRoundedRectangle;
    private createEllipse;
    private createTriangle;
    private createRightTriangle;
    private createDiamond;
    private createParallelogram;
    private createTrapezoid;
    private createPentagon;
    private createHexagon;
    private createOctagon;
    private createLine;
    private createArrow;
    /**
     * Computes arrow head length and shaft thickness from adjustment values.
     * Per the ECMA-376 preset definitions (rightArrow/leftArrow/upArrow/downArrow):
     * adj1 is the shaft thickness as a fraction of the cross-axis dimension
     * (dy1 = h * a1 / 200000 for horizontal arrows, shaft = 2 * dy1), and adj2 is
     * the head length relative to the shortest side (dx2 = ss * a2 / 100000),
     * pinned so the head never exceeds the arrow's long axis.
     */
    private getArrowMetrics;
    private createHorizontalArrow;
    private createVerticalArrow;
    /**
     * Creates a chevron shape (arrow pointing right with V-notch on left).
     * The 'adj' value controls how far the point extends (default 50000 = 50%).
     * Per OOXML spec, the indent is based on height, not width.
     */
    private createChevron;
    /**
     * Creates a homePlate shape (pentagon arrow pointing right).
     * Similar to chevron but with a flat back.
     */
    private createHomePlate;
    private createStar;
    private createPlus;
    private createHeart;
    private createWedgeRectCallout;
    private createFlowChartTerminator;
    private createRegularPolygon;
    /**
     * Gets the text bounds for a shape, accounting for non-rectangular shapes.
     * For shapes like chevrons, the text area is smaller than the full bounds.
     *
     * @param presetName The preset geometry name
     * @param bounds The full shape bounds
     * @param adjustValues Optional adjustment values
     * @returns The bounds to use for text layout
     */
    getTextBounds(presetName: string, bounds: Rect, adjustValues?: Map<string, number>): Rect;
}
/**
 * Default preset geometry calculator instance.
 */
export declare const presetGeometryCalculator: PresetGeometryCalculator;
