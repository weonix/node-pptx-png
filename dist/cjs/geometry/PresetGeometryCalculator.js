"use strict";
/**
 * Generates paths for preset shapes defined in ECMA-376.
 *
 * A small set of common shapes is hand-coded (kept as overrides so existing
 * visual baselines stay stable); every other preset geometry is evaluated
 * from the generated ECMA-376 definition table (presetShapeDefinitions.ts)
 * through the guide-formula interpreter in GeometryEngine.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.presetGeometryCalculator = exports.PresetGeometryCalculator = void 0;
const PathBuilder_js_1 = require("./PathBuilder.js");
const GeometryEngine_js_1 = require("./GeometryEngine.js");
const presetShapeDefinitions_js_1 = require("./presetShapeDefinitions.js");
/**
 * Adjustment value scale factor.
 * OpenXML adjustment values are percentages as integers (0-100000 = 0-100%).
 */
const ADJUSTMENT_SCALE = 100000;
/**
 * Default adjustment values for various shapes.
 */
const DEFAULT_ADJUSTMENTS = {
    roundRect: { adj: 16667 }, // ~16.67% corner radius
    parallelogram: { adj: 25000 }, // 25% offset
    trapezoid: { adj: 25000 }, // 25% offset
    hexagon: { adj: 25000 }, // 25% offset
    octagon: { adj: 29289 }, // ~29.3% (1 - sqrt(2)/2)
    rightArrow: { adj1: 50000, adj2: 50000 },
    leftArrow: { adj1: 50000, adj2: 50000 },
    upArrow: { adj1: 50000, adj2: 50000 },
    downArrow: { adj1: 50000, adj2: 50000 },
    chevron: { adj: 50000 },
    homePlate: { adj: 50000 },
    plus: { adj: 25000 },
    wedgeRectCallout: { adj1: -20000, adj2: 62500 },
};
/**
 * Calculator for preset geometry shapes.
 */
class PresetGeometryCalculator {
    /**
     * List of hand-coded preset geometry names. These take precedence over
     * the generated ECMA-376 definitions to keep visual baselines stable.
     */
    static SUPPORTED_SHAPES = [
        // Basic shapes
        'rect',
        'roundRect',
        'ellipse',
        'triangle',
        'rtTriangle',
        'diamond',
        'parallelogram',
        'trapezoid',
        'pentagon',
        'hexagon',
        'octagon',
        'line',
        // Arrows
        'rightArrow',
        'leftArrow',
        'upArrow',
        'downArrow',
        // Stars
        'star5',
        // Plus and special
        'plus',
        'heart',
        // Callouts
        'wedgeRectCallout',
        // Flowchart
        'flowChartProcess',
        'flowChartDecision',
        'flowChartTerminator',
    ];
    /**
     * Checks if a preset geometry is supported (hand-coded or via the
     * generated ECMA-376 definition table).
     */
    isSupported(presetName) {
        return (this.isHandCoded(presetName) ||
            Object.prototype.hasOwnProperty.call(presetShapeDefinitions_js_1.PRESET_SHAPE_DEFINITIONS, presetName));
    }
    /**
     * Names handled by the hand-coded switch in createHandCodedPath.
     * Must stay in sync with its cases ('oval', 'chevron' and 'homePlate'
     * are handled there without appearing in SUPPORTED_SHAPES).
     */
    static HAND_CODED_SHAPES = new Set([
        ...PresetGeometryCalculator.SUPPORTED_SHAPES,
        'oval',
        'chevron',
        'homePlate',
    ]);
    /**
     * Checks if a preset geometry has a hand-coded implementation.
     */
    isHandCoded(presetName) {
        return PresetGeometryCalculator.HAND_CODED_SHAPES.has(presetName);
    }
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
    createPaths(presetName, bounds, adjustValues) {
        if (this.isHandCoded(presetName)) {
            const path = this.createHandCodedPath(presetName, bounds, adjustValues);
            return path ? [path] : undefined;
        }
        return this.createDefinitionPaths(presetName, bounds, adjustValues);
    }
    /**
     * Evaluates a preset from the generated ECMA-376 definition table.
     */
    createDefinitionPaths(presetName, bounds, adjustValues) {
        const definition = presetShapeDefinitions_js_1.PRESET_SHAPE_DEFINITIONS[presetName];
        if (!definition)
            return undefined;
        // Allow one dimension to be zero (connectors, flat shapes) but not both
        if (bounds.width <= 0 && bounds.height <= 0)
            return undefined;
        try {
            const paths = (0, GeometryEngine_js_1.buildGeometryPaths)(definition, bounds, adjustValues);
            return paths.length > 0 ? paths : undefined;
        }
        catch {
            // Malformed adjust values or formulas: let the caller fall back
            return undefined;
        }
    }
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
    createPath(presetName, bounds, adjustValues) {
        if (!this.isHandCoded(presetName)) {
            const paths = this.createDefinitionPaths(presetName, bounds, adjustValues);
            if (!paths || paths.length === 0)
                return undefined;
            if (paths.length === 1)
                return paths[0];
            return {
                segments: paths.flatMap((p) => p.segments),
                fill: paths.some((p) => p.fill !== false),
                stroke: paths.some((p) => p.stroke !== false),
            };
        }
        return this.createHandCodedPath(presetName, bounds, adjustValues);
    }
    /**
     * Creates a path for a hand-coded preset geometry.
     */
    createHandCodedPath(presetName, bounds, adjustValues) {
        // Validate bounds
        if (presetName === 'line') {
            // Lines can have zero height or width
            if (bounds.width <= 0 && bounds.height <= 0)
                return undefined;
        }
        else {
            if (bounds.width <= 0 || bounds.height <= 0)
                return undefined;
        }
        const adj = this.resolveAdjustments(presetName, adjustValues);
        switch (presetName) {
            // Basic shapes
            case 'rect':
                return this.createRectangle(bounds);
            case 'roundRect':
                return this.createRoundedRectangle(bounds, adj);
            case 'ellipse':
            case 'oval':
                return this.createEllipse(bounds);
            case 'triangle':
                return this.createTriangle(bounds);
            case 'rtTriangle':
                return this.createRightTriangle(bounds);
            case 'diamond':
                return this.createDiamond(bounds);
            case 'parallelogram':
                return this.createParallelogram(bounds, adj);
            case 'trapezoid':
                return this.createTrapezoid(bounds, adj);
            case 'pentagon':
                return this.createPentagon(bounds);
            case 'hexagon':
                return this.createHexagon(bounds, adj);
            case 'octagon':
                return this.createOctagon(bounds, adj);
            case 'line':
                return this.createLine(bounds);
            // Arrows
            case 'rightArrow':
                return this.createArrow(bounds, 'right', adj);
            case 'leftArrow':
                return this.createArrow(bounds, 'left', adj);
            case 'upArrow':
                return this.createArrow(bounds, 'up', adj);
            case 'downArrow':
                return this.createArrow(bounds, 'down', adj);
            case 'chevron':
                return this.createChevron(bounds, adj);
            case 'homePlate':
                return this.createHomePlate(bounds, adj);
            // Stars
            case 'star5':
                return this.createStar(bounds, 5);
            // Plus and special
            case 'plus':
                return this.createPlus(bounds, adj);
            case 'heart':
                return this.createHeart(bounds);
            // Callouts
            case 'wedgeRectCallout':
                return this.createWedgeRectCallout(bounds, adj);
            // Flowchart
            case 'flowChartProcess':
                return this.createRectangle(bounds);
            case 'flowChartDecision':
                return this.createDiamond(bounds);
            case 'flowChartTerminator':
                return this.createFlowChartTerminator(bounds);
            default:
                return undefined;
        }
    }
    /**
     * Resolves adjustment values, using defaults for missing values.
     */
    resolveAdjustments(presetName, adjustValues) {
        const defaults = DEFAULT_ADJUSTMENTS[presetName] ?? {};
        const result = { ...defaults };
        if (adjustValues) {
            for (const [key, value] of adjustValues) {
                result[key] = value;
            }
        }
        return result;
    }
    /**
     * Gets an adjustment value with fallback.
     */
    getAdj(adj, name, defaultValue) {
        return adj[name] ?? defaultValue;
    }
    // ============================================================
    // Basic Shapes
    // ============================================================
    createRectangle(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.addRectangle(bounds.x, bounds.y, bounds.width, bounds.height);
        return builder.build();
    }
    createRoundedRectangle(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 16667);
        const minDim = Math.min(bounds.width, bounds.height);
        let cornerRadius = minDim * (adjValue / ADJUSTMENT_SCALE);
        // Clamp to half of minimum dimension
        cornerRadius = Math.min(cornerRadius, minDim / 2);
        if (cornerRadius <= 0) {
            return this.createRectangle(bounds);
        }
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.addRoundedRectangle(bounds.x, bounds.y, bounds.width, bounds.height, cornerRadius);
        return builder.build();
    }
    createEllipse(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.addEllipse(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2, bounds.width / 2, bounds.height / 2);
        return builder.build();
    }
    createTriangle(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        // Isoceles triangle: top-center, bottom-left, bottom-right
        builder.moveTo(bounds.x + bounds.width / 2, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height);
        builder.lineTo(bounds.x, bounds.y + bounds.height);
        builder.closePath();
        return builder.build();
    }
    createRightTriangle(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        // Right triangle: top-left, bottom-left, bottom-right
        builder.moveTo(bounds.x, bounds.y);
        builder.lineTo(bounds.x, bounds.y + bounds.height);
        builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height);
        builder.closePath();
        return builder.build();
    }
    createDiamond(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        const centerX = bounds.x + bounds.width / 2;
        const centerY = bounds.y + bounds.height / 2;
        builder.moveTo(centerX, bounds.y); // Top
        builder.lineTo(bounds.x + bounds.width, centerY); // Right
        builder.lineTo(centerX, bounds.y + bounds.height); // Bottom
        builder.lineTo(bounds.x, centerY); // Left
        builder.closePath();
        return builder.build();
    }
    createParallelogram(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 25000);
        const offset = bounds.width * (adjValue / ADJUSTMENT_SCALE);
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.moveTo(bounds.x + offset, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bounds.y);
        builder.lineTo(bounds.x + bounds.width - offset, bounds.y + bounds.height);
        builder.lineTo(bounds.x, bounds.y + bounds.height);
        builder.closePath();
        return builder.build();
    }
    createTrapezoid(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 25000);
        const offset = bounds.width * (adjValue / ADJUSTMENT_SCALE);
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.moveTo(bounds.x + offset, bounds.y);
        builder.lineTo(bounds.x + bounds.width - offset, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height);
        builder.lineTo(bounds.x, bounds.y + bounds.height);
        builder.closePath();
        return builder.build();
    }
    createPentagon(bounds) {
        return this.createRegularPolygon(bounds, 5, -90);
    }
    createHexagon(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 25000);
        const offset = bounds.width * (adjValue / ADJUSTMENT_SCALE);
        const centerY = bounds.y + bounds.height / 2;
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.moveTo(bounds.x + offset, bounds.y);
        builder.lineTo(bounds.x + bounds.width - offset, bounds.y);
        builder.lineTo(bounds.x + bounds.width, centerY);
        builder.lineTo(bounds.x + bounds.width - offset, bounds.y + bounds.height);
        builder.lineTo(bounds.x + offset, bounds.y + bounds.height);
        builder.lineTo(bounds.x, centerY);
        builder.closePath();
        return builder.build();
    }
    createOctagon(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 29289);
        const offsetX = bounds.width * (adjValue / ADJUSTMENT_SCALE);
        const offsetY = bounds.height * (adjValue / ADJUSTMENT_SCALE);
        const builder = new PathBuilder_js_1.PathBuilder();
        builder.moveTo(bounds.x + offsetX, bounds.y);
        builder.lineTo(bounds.x + bounds.width - offsetX, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bounds.y + offsetY);
        builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height - offsetY);
        builder.lineTo(bounds.x + bounds.width - offsetX, bounds.y + bounds.height);
        builder.lineTo(bounds.x + offsetX, bounds.y + bounds.height);
        builder.lineTo(bounds.x, bounds.y + bounds.height - offsetY);
        builder.lineTo(bounds.x, bounds.y + offsetY);
        builder.closePath();
        return builder.build();
    }
    createLine(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        if (bounds.height < 1) {
            // Horizontal line
            builder.moveTo(bounds.x, bounds.y);
            builder.lineTo(bounds.x + bounds.width, bounds.y);
        }
        else if (bounds.width < 1) {
            // Vertical line
            builder.moveTo(bounds.x, bounds.y);
            builder.lineTo(bounds.x, bounds.y + bounds.height);
        }
        else {
            // Diagonal line
            builder.moveTo(bounds.x, bounds.y);
            builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height);
        }
        return builder.build({ fill: false, stroke: true });
    }
    // ============================================================
    // Arrows
    // ============================================================
    createArrow(bounds, direction, adj) {
        const adj1 = this.getAdj(adj, 'adj1', 50000);
        const adj2 = this.getAdj(adj, 'adj2', 50000);
        const builder = new PathBuilder_js_1.PathBuilder();
        switch (direction) {
            case 'right':
                this.createHorizontalArrow(builder, bounds, adj1, adj2, false);
                break;
            case 'left':
                this.createHorizontalArrow(builder, bounds, adj1, adj2, true);
                break;
            case 'up':
                this.createVerticalArrow(builder, bounds, adj1, adj2, true);
                break;
            case 'down':
                this.createVerticalArrow(builder, bounds, adj1, adj2, false);
                break;
        }
        return builder.build();
    }
    /**
     * Computes arrow head length and shaft thickness from adjustment values.
     * Per the ECMA-376 preset definitions (rightArrow/leftArrow/upArrow/downArrow):
     * adj1 is the shaft thickness as a fraction of the cross-axis dimension
     * (dy1 = h * a1 / 200000 for horizontal arrows, shaft = 2 * dy1), and adj2 is
     * the head length relative to the shortest side (dx2 = ss * a2 / 100000),
     * pinned so the head never exceeds the arrow's long axis.
     */
    getArrowMetrics(bounds, adj1, adj2, orientation) {
        const shortestSide = Math.min(bounds.width, bounds.height);
        const axisLength = orientation === 'horizontal' ? bounds.width : bounds.height;
        const crossLength = orientation === 'horizontal' ? bounds.height : bounds.width;
        // a1 = pin(0, adj1, 100000); a2 = pin(0, adj2, 100000 * axis / ss)
        const shaftRatio = Math.min(Math.max(adj1, 0), ADJUSTMENT_SCALE) / ADJUSTMENT_SCALE;
        const headRatio = Math.max(adj2, 0) / ADJUSTMENT_SCALE;
        return {
            headLength: Math.min(shortestSide * headRatio, axisLength),
            shaftThickness: crossLength * shaftRatio,
        };
    }
    createHorizontalArrow(builder, bounds, adj1, adj2, leftPointing) {
        const { headLength, shaftThickness } = this.getArrowMetrics(bounds, adj1, adj2, 'horizontal');
        const headWidth = headLength;
        const shaftHeight = shaftThickness;
        const shaftTop = bounds.y + (bounds.height - shaftHeight) / 2;
        const shaftBottom = shaftTop + shaftHeight;
        const centerY = bounds.y + bounds.height / 2;
        if (leftPointing) {
            builder.moveTo(bounds.x, centerY);
            builder.lineTo(bounds.x + headWidth, bounds.y);
            builder.lineTo(bounds.x + headWidth, shaftTop);
            builder.lineTo(bounds.x + bounds.width, shaftTop);
            builder.lineTo(bounds.x + bounds.width, shaftBottom);
            builder.lineTo(bounds.x + headWidth, shaftBottom);
            builder.lineTo(bounds.x + headWidth, bounds.y + bounds.height);
        }
        else {
            builder.moveTo(bounds.x + bounds.width, centerY);
            builder.lineTo(bounds.x + bounds.width - headWidth, bounds.y);
            builder.lineTo(bounds.x + bounds.width - headWidth, shaftTop);
            builder.lineTo(bounds.x, shaftTop);
            builder.lineTo(bounds.x, shaftBottom);
            builder.lineTo(bounds.x + bounds.width - headWidth, shaftBottom);
            builder.lineTo(bounds.x + bounds.width - headWidth, bounds.y + bounds.height);
        }
        builder.closePath();
    }
    createVerticalArrow(builder, bounds, adj1, adj2, upPointing) {
        const { headLength, shaftThickness } = this.getArrowMetrics(bounds, adj1, adj2, 'vertical');
        const headHeight = headLength;
        const shaftWidth = shaftThickness;
        const shaftLeft = bounds.x + (bounds.width - shaftWidth) / 2;
        const shaftRight = shaftLeft + shaftWidth;
        const centerX = bounds.x + bounds.width / 2;
        if (upPointing) {
            builder.moveTo(centerX, bounds.y);
            builder.lineTo(bounds.x + bounds.width, bounds.y + headHeight);
            builder.lineTo(shaftRight, bounds.y + headHeight);
            builder.lineTo(shaftRight, bounds.y + bounds.height);
            builder.lineTo(shaftLeft, bounds.y + bounds.height);
            builder.lineTo(shaftLeft, bounds.y + headHeight);
            builder.lineTo(bounds.x, bounds.y + headHeight);
        }
        else {
            builder.moveTo(centerX, bounds.y + bounds.height);
            builder.lineTo(bounds.x + bounds.width, bounds.y + bounds.height - headHeight);
            builder.lineTo(shaftRight, bounds.y + bounds.height - headHeight);
            builder.lineTo(shaftRight, bounds.y);
            builder.lineTo(shaftLeft, bounds.y);
            builder.lineTo(shaftLeft, bounds.y + bounds.height - headHeight);
            builder.lineTo(bounds.x, bounds.y + bounds.height - headHeight);
        }
        builder.closePath();
    }
    /**
     * Creates a chevron shape (arrow pointing right with V-notch on left).
     * The 'adj' value controls how far the point extends (default 50000 = 50%).
     * Per OOXML spec, the indent is based on height, not width.
     */
    createChevron(bounds, adj) {
        const builder = new PathBuilder_js_1.PathBuilder();
        // adj controls the indentation as percentage of height (not width)
        // This maintains proper chevron proportions regardless of aspect ratio
        const adjValue = this.getAdj(adj, 'adj', 50000);
        const indentRatio = adjValue / ADJUSTMENT_SCALE;
        // The indent is based on height to maintain consistent chevron appearance
        const indent = bounds.height * indentRatio;
        const centerY = bounds.y + bounds.height / 2;
        // Draw chevron shape (arrow pointing right with V-notch on left)
        // The shape is like a pentagon with a triangular notch cut from the left side
        //
        //    1 -------- 2
        //     \          \
        //      6          3
        //     /          /
        //    5 -------- 4
        //
        // Start at top-left outer corner
        builder.moveTo(bounds.x, bounds.y); // 1
        // Go to top-right corner (before the point)
        builder.lineTo(bounds.x + bounds.width - indent, bounds.y); // 2
        // Go to right point (tip of arrow)
        builder.lineTo(bounds.x + bounds.width, centerY); // 3
        // Go to bottom-right corner (after the point)
        builder.lineTo(bounds.x + bounds.width - indent, bounds.y + bounds.height); // 4
        // Go to bottom-left outer corner
        builder.lineTo(bounds.x, bounds.y + bounds.height); // 5
        // Go to left notch point (V pointing into the shape)
        builder.lineTo(bounds.x + indent, centerY); // 6
        builder.closePath();
        return builder.build();
    }
    /**
     * Creates a homePlate shape (pentagon arrow pointing right).
     * Similar to chevron but with a flat back.
     */
    createHomePlate(bounds, adj) {
        const builder = new PathBuilder_js_1.PathBuilder();
        const adjValue = this.getAdj(adj, 'adj', 50000);
        const pointRatio = adjValue / ADJUSTMENT_SCALE;
        const pointWidth = bounds.width * pointRatio;
        const centerY = bounds.y + bounds.height / 2;
        // Draw homePlate shape (pentagon pointing right)
        builder.moveTo(bounds.x, bounds.y);
        builder.lineTo(bounds.x + bounds.width - pointWidth, bounds.y);
        builder.lineTo(bounds.x + bounds.width, centerY);
        builder.lineTo(bounds.x + bounds.width - pointWidth, bounds.y + bounds.height);
        builder.lineTo(bounds.x, bounds.y + bounds.height);
        builder.closePath();
        return builder.build();
    }
    // ============================================================
    // Stars
    // ============================================================
    createStar(bounds, points) {
        const builder = new PathBuilder_js_1.PathBuilder();
        const centerX = bounds.x + bounds.width / 2;
        const centerY = bounds.y + bounds.height / 2;
        const outerRadiusX = bounds.width / 2;
        const outerRadiusY = bounds.height / 2;
        const innerRadiusX = outerRadiusX * 0.4;
        const innerRadiusY = outerRadiusY * 0.4;
        const startAngle = -Math.PI / 2; // Start at top
        const angleStep = Math.PI / points;
        for (let i = 0; i < points * 2; i++) {
            const angle = startAngle + i * angleStep;
            const isOuter = i % 2 === 0;
            const rx = isOuter ? outerRadiusX : innerRadiusX;
            const ry = isOuter ? outerRadiusY : innerRadiusY;
            const x = centerX + rx * Math.cos(angle);
            const y = centerY + ry * Math.sin(angle);
            if (i === 0) {
                builder.moveTo(x, y);
            }
            else {
                builder.lineTo(x, y);
            }
        }
        builder.closePath();
        return builder.build();
    }
    // ============================================================
    // Plus and Special Shapes
    // ============================================================
    createPlus(bounds, adj) {
        const adjValue = this.getAdj(adj, 'adj', 25000);
        const armThickness = Math.min(bounds.width, bounds.height) * (adjValue / ADJUSTMENT_SCALE);
        const builder = new PathBuilder_js_1.PathBuilder();
        const centerX = bounds.x + bounds.width / 2;
        const centerY = bounds.y + bounds.height / 2;
        const halfArm = armThickness / 2;
        // Draw plus sign clockwise from top-left of top arm
        builder.moveTo(centerX - halfArm, bounds.y);
        builder.lineTo(centerX + halfArm, bounds.y);
        builder.lineTo(centerX + halfArm, centerY - halfArm);
        builder.lineTo(bounds.x + bounds.width, centerY - halfArm);
        builder.lineTo(bounds.x + bounds.width, centerY + halfArm);
        builder.lineTo(centerX + halfArm, centerY + halfArm);
        builder.lineTo(centerX + halfArm, bounds.y + bounds.height);
        builder.lineTo(centerX - halfArm, bounds.y + bounds.height);
        builder.lineTo(centerX - halfArm, centerY + halfArm);
        builder.lineTo(bounds.x, centerY + halfArm);
        builder.lineTo(bounds.x, centerY - halfArm);
        builder.lineTo(centerX - halfArm, centerY - halfArm);
        builder.closePath();
        return builder.build();
    }
    createHeart(bounds) {
        const builder = new PathBuilder_js_1.PathBuilder();
        const centerX = bounds.x + bounds.width / 2;
        const topY = bounds.y + bounds.height * 0.25;
        const bottomY = bounds.y + bounds.height;
        // Start at bottom point
        builder.moveTo(centerX, bottomY);
        // Left side curve
        builder.cubicBezierTo(bounds.x - bounds.width * 0.1, bounds.y + bounds.height * 0.6, bounds.x, bounds.y, centerX, topY);
        // Right side curve
        builder.cubicBezierTo(bounds.x + bounds.width, bounds.y, bounds.x + bounds.width + bounds.width * 0.1, bounds.y + bounds.height * 0.6, centerX, bottomY);
        builder.closePath();
        return builder.build();
    }
    // ============================================================
    // Callouts
    // ============================================================
    createWedgeRectCallout(bounds, adj) {
        const adj1 = this.getAdj(adj, 'adj1', -20000);
        const adj2 = this.getAdj(adj, 'adj2', 62500);
        const builder = new PathBuilder_js_1.PathBuilder();
        const bodyBottom = bounds.y + bounds.height - bounds.height * 0.2;
        // Pointer tip position
        const pointerTipX = bounds.x + bounds.width / 2 + bounds.width * (adj1 / ADJUSTMENT_SCALE);
        let pointerTipY = bounds.y + bounds.height * (adj2 / ADJUSTMENT_SCALE);
        // Clamp pointer tip to be outside the body
        pointerTipY = Math.max(pointerTipY, bodyBottom);
        // Pointer base positions
        const pointerBaseLeft = bounds.x + bounds.width * 0.35;
        const pointerBaseRight = bounds.x + bounds.width * 0.65;
        builder.moveTo(bounds.x, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bounds.y);
        builder.lineTo(bounds.x + bounds.width, bodyBottom);
        builder.lineTo(pointerBaseRight, bodyBottom);
        builder.lineTo(pointerTipX, pointerTipY);
        builder.lineTo(pointerBaseLeft, bodyBottom);
        builder.lineTo(bounds.x, bodyBottom);
        builder.closePath();
        return builder.build();
    }
    // ============================================================
    // Flowchart Shapes
    // ============================================================
    createFlowChartTerminator(bounds) {
        // Stadium/pill shape (rounded rectangle with semicircular ends)
        const builder = new PathBuilder_js_1.PathBuilder();
        const radius = bounds.height / 2;
        if (bounds.width < bounds.height) {
            // Too narrow, just use ellipse
            return this.createEllipse(bounds);
        }
        // Start at top edge, after the left semicircle
        builder.moveTo(bounds.x + radius, bounds.y);
        // Top edge (left to right)
        builder.lineTo(bounds.x + bounds.width - radius, bounds.y);
        // Right semicircle: from top-right going clockwise to bottom-right
        // Arc from (x + width - radius, y) to (x + width - radius, y + height)
        // The arc goes through (x + width, y + radius) at the rightmost point
        builder.arcTo(radius, radius, 0, false, true, bounds.x + bounds.width - radius, bounds.y + bounds.height);
        // Bottom edge (right to left)
        builder.lineTo(bounds.x + radius, bounds.y + bounds.height);
        // Left semicircle: from bottom-left going clockwise back to top-left
        // Arc from (x + radius, y + height) to (x + radius, y)
        // The arc goes through (x, y + radius) at the leftmost point
        builder.arcTo(radius, radius, 0, false, true, bounds.x + radius, bounds.y);
        builder.closePath();
        return builder.build();
    }
    // ============================================================
    // Helper Methods
    // ============================================================
    createRegularPolygon(bounds, sides, startAngleDegrees) {
        const builder = new PathBuilder_js_1.PathBuilder();
        const centerX = bounds.x + bounds.width / 2;
        const centerY = bounds.y + bounds.height / 2;
        const radiusX = bounds.width / 2;
        const radiusY = bounds.height / 2;
        const startAngle = (startAngleDegrees * Math.PI) / 180;
        const angleStep = (2 * Math.PI) / sides;
        for (let i = 0; i < sides; i++) {
            const angle = startAngle + i * angleStep;
            const x = centerX + radiusX * Math.cos(angle);
            const y = centerY + radiusY * Math.sin(angle);
            if (i === 0) {
                builder.moveTo(x, y);
            }
            else {
                builder.lineTo(x, y);
            }
        }
        builder.closePath();
        return builder.build();
    }
    /**
     * Gets the text bounds for a shape, accounting for non-rectangular shapes.
     * For shapes like chevrons, the text area is smaller than the full bounds.
     *
     * @param presetName The preset geometry name
     * @param bounds The full shape bounds
     * @param adjustValues Optional adjustment values
     * @returns The bounds to use for text layout
     */
    getTextBounds(presetName, bounds, adjustValues) {
        const adj = this.resolveAdjustments(presetName, adjustValues);
        switch (presetName) {
            case 'chevron': {
                // Chevron has indents on both left (notch) and right (point) based on height
                const adjValue = this.getAdj(adj, 'adj', 50000);
                const indentRatio = adjValue / ADJUSTMENT_SCALE;
                const indent = bounds.height * indentRatio; // Height-based, matching the shape geometry
                // Text area starts after left indent and ends before right indent
                return {
                    x: bounds.x + indent,
                    y: bounds.y,
                    width: bounds.width - indent * 2,
                    height: bounds.height,
                };
            }
            case 'homePlate': {
                // HomePlate only has indent on the right (arrow point)
                const adjValue = this.getAdj(adj, 'adj', 50000);
                const pointRatio = adjValue / ADJUSTMENT_SCALE;
                const pointWidth = bounds.width * pointRatio;
                return {
                    x: bounds.x,
                    y: bounds.y,
                    width: bounds.width - pointWidth,
                    height: bounds.height,
                };
            }
            case 'rightArrow': {
                // Arrow has shaft and head, text goes in shaft area.
                // Metrics match createHorizontalArrow: adj1 = shaft thickness,
                // adj2 = head length relative to the shortest side.
                const adj1 = this.getAdj(adj, 'adj1', 50000);
                const adj2 = this.getAdj(adj, 'adj2', 50000);
                const { headLength, shaftThickness } = this.getArrowMetrics(bounds, adj1, adj2, 'horizontal');
                const shaftY = bounds.y + (bounds.height - shaftThickness) / 2;
                return {
                    x: bounds.x,
                    y: shaftY,
                    width: bounds.width - headLength,
                    height: shaftThickness,
                };
            }
            case 'leftArrow': {
                const adj1 = this.getAdj(adj, 'adj1', 50000);
                const adj2 = this.getAdj(adj, 'adj2', 50000);
                const { headLength, shaftThickness } = this.getArrowMetrics(bounds, adj1, adj2, 'horizontal');
                const shaftY = bounds.y + (bounds.height - shaftThickness) / 2;
                return {
                    x: bounds.x + headLength,
                    y: shaftY,
                    width: bounds.width - headLength,
                    height: shaftThickness,
                };
            }
            default: {
                // Definition-based shapes: use the ECMA-376 text rectangle.
                // Hand-coded shapes keep their existing behavior (full bounds).
                if (!this.isHandCoded(presetName)) {
                    const definition = presetShapeDefinitions_js_1.PRESET_SHAPE_DEFINITIONS[presetName];
                    if (definition) {
                        const textRect = (0, GeometryEngine_js_1.computeTextRect)(definition, bounds, adjustValues);
                        if (textRect)
                            return textRect;
                    }
                }
                return bounds;
            }
        }
    }
}
exports.PresetGeometryCalculator = PresetGeometryCalculator;
/**
 * Default preset geometry calculator instance.
 */
exports.presetGeometryCalculator = new PresetGeometryCalculator();
