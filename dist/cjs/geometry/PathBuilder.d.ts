/**
 * Path construction utilities for building canvas paths.
 * Provides a fluent API for constructing paths that can be rendered to canvas.
 */
import { Path2D, type CanvasRenderingContext2D } from 'skia-canvas';
import type { Point, Path, PathBounds } from '../types/geometry.js';
/**
 * Builder for constructing paths from segments.
 * Supports moveTo, lineTo, curveTo, arcTo, and closePath operations.
 */
export declare class PathBuilder {
    private segments;
    private currentPoint;
    private startPoint;
    private hasMoved;
    /**
     * Moves the current point without drawing.
     */
    moveTo(x: number, y: number): this;
    /**
     * Draws a line from the current point to the specified point.
     */
    lineTo(x: number, y: number): this;
    /**
     * Draws a cubic bezier curve.
     * @param cp1x First control point X
     * @param cp1y First control point Y
     * @param cp2x Second control point X
     * @param cp2y Second control point Y
     * @param x End point X
     * @param y End point Y
     */
    cubicBezierTo(cp1x: number, cp1y: number, cp2x: number, cp2y: number, x: number, y: number): this;
    /**
     * Draws a quadratic bezier curve.
     * @param cpx Control point X
     * @param cpy Control point Y
     * @param x End point X
     * @param y End point Y
     */
    quadBezierTo(cpx: number, cpy: number, x: number, y: number): this;
    /**
     * Draws an elliptical arc.
     * Uses SVG arc notation for compatibility.
     * @param rx Radius X
     * @param ry Radius Y
     * @param xAxisRotation Rotation of the ellipse in degrees
     * @param largeArcFlag Whether to use the larger arc
     * @param sweepFlag Direction of the arc (true = clockwise)
     * @param x End point X
     * @param y End point Y
     */
    arcTo(rx: number, ry: number, xAxisRotation: number, largeArcFlag: boolean, sweepFlag: boolean, x: number, y: number): this;
    /**
     * Closes the current path by drawing a line to the start point.
     */
    closePath(): this;
    /**
     * Adds a rectangle to the path.
     */
    addRectangle(x: number, y: number, width: number, height: number): this;
    /**
     * Adds an ellipse to the path using bezier curves.
     * @param cx Center X
     * @param cy Center Y
     * @param rx Radius X
     * @param ry Radius Y
     */
    addEllipse(cx: number, cy: number, rx: number, ry: number): this;
    /**
     * Adds a rounded rectangle to the path.
     * @param x Top-left X
     * @param y Top-left Y
     * @param width Width
     * @param height Height
     * @param radius Corner radius
     */
    addRoundedRectangle(x: number, y: number, width: number, height: number, radius: number): this;
    /**
     * Resets the builder to start a new path.
     */
    reset(): this;
    /**
     * Builds and returns the completed path.
     */
    build(options?: {
        fill?: boolean;
        stroke?: boolean;
    }): Path;
    /**
     * Gets the current point.
     */
    getCurrentPoint(): Point;
}
/**
 * Calculates the bounding box of a path.
 */
export declare function calculatePathBounds(path: Path): PathBounds;
/**
 * Applies a path to a canvas 2D context.
 * Creates a Path2D object or draws directly to the context.
 */
export declare function applyPathToContext(ctx: CanvasRenderingContext2D, path: Path, startNewPath?: boolean): void;
/**
 * Creates a Path2D object from a path definition.
 */
export declare function pathToPath2D(path: Path): Path2D;
