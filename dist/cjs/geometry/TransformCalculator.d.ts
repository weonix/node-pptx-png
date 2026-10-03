/**
 * Handles shape transforms (position, size, rotation, flip).
 * Parses xfrm element from OpenXML and applies transforms to canvas context.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { ShapeTransform, Rect, Point } from '../types/geometry.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import { UnitConverter } from '../core/UnitConverter.js';
/**
 * Transform data parsed from xfrm element.
 */
export interface ParsedTransform {
    /** X offset in EMU */
    x: number;
    /** Y offset in EMU */
    y: number;
    /** Width in EMU */
    width: number;
    /** Height in EMU */
    height: number;
    /** Rotation in degrees (0-360) */
    rotation: number;
    /** Horizontal flip */
    flipH: boolean;
    /** Vertical flip */
    flipV: boolean;
}
/**
 * Pixel-space transform after unit conversion.
 */
export interface PixelTransform {
    /** X offset in pixels */
    x: number;
    /** Y offset in pixels */
    y: number;
    /** Width in pixels */
    width: number;
    /** Height in pixels */
    height: number;
    /** Rotation in radians */
    rotation: number;
    /** Horizontal flip */
    flipH: boolean;
    /** Vertical flip */
    flipV: boolean;
}
/**
 * Calculator for shape transforms.
 */
export declare class TransformCalculator {
    private readonly unitConverter;
    constructor(unitConverter?: UnitConverter);
    /**
     * Parses an xfrm element from shape properties.
     * @param spPr Shape properties node containing a:xfrm
     * @returns Parsed transform or undefined if no transform found
     */
    parseTransform(spPr: PptxXmlNode | undefined): ParsedTransform | undefined;
    /**
     * Parses an xfrm element directly.
     * @param xfrm The a:xfrm element node
     * @returns Parsed transform
     */
    parseXfrmElement(xfrm: PptxXmlNode): ParsedTransform;
    /**
     * Converts EMU transform to pixel transform.
     * @param transform EMU-based transform
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     * @returns Pixel-based transform
     */
    toPixelTransform(transform: ParsedTransform, scaleX: number, scaleY: number): PixelTransform;
    /**
     * Applies a transform to a canvas context.
     * The transform is applied as: translate -> rotate -> flip -> translate to origin.
     *
     * **Important:** This method does NOT call ctx.save(). The caller is responsible
     * for managing the canvas context state with save/restore pairs.
     *
     * @param ctx Canvas 2D context
     * @param transform Pixel-based transform to apply
     */
    applyTransform(ctx: CanvasRenderingContext2D, transform: PixelTransform): void;
    /**
     * Resets the canvas context to before the transform was applied.
     *
     * **Note:** This is a convenience wrapper around ctx.restore(). Since applyTransform
     * does not call ctx.save(), the caller must ensure proper save/restore pairing.
     *
     * @param ctx Canvas 2D context
     * @deprecated Use ctx.restore() directly with your own save/restore management
     */
    resetTransform(ctx: CanvasRenderingContext2D): void;
    /**
     * Gets the bounding rectangle in pixel space for a transform.
     * Takes rotation into account for the axis-aligned bounding box.
     * @param transform Pixel-based transform
     * @returns Axis-aligned bounding rectangle
     */
    getBoundingRect(transform: PixelTransform): Rect;
    /**
     * Gets the four corners of a transformed rectangle in pixel space.
     * @param transform Pixel-based transform
     * @returns Array of four corner points
     */
    getTransformedCorners(transform: PixelTransform): Point[];
    /**
     * Converts a ShapeTransform to ParsedTransform.
     * @param shapeTransform Shape transform from types
     * @returns Parsed transform
     */
    fromShapeTransform(shapeTransform: ShapeTransform): ParsedTransform;
    /**
     * Converts a ParsedTransform to ShapeTransform.
     * @param parsed Parsed transform
     * @returns Shape transform for types
     */
    toShapeTransform(parsed: ParsedTransform): ShapeTransform;
}
/**
 * Default transform calculator instance.
 */
export declare const defaultTransformCalculator: TransformCalculator;
