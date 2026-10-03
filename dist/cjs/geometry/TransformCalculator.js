"use strict";
/**
 * Handles shape transforms (position, size, rotation, flip).
 * Parses xfrm element from OpenXML and applies transforms to canvas context.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.defaultTransformCalculator = exports.TransformCalculator = void 0;
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
/**
 * Calculator for shape transforms.
 */
class TransformCalculator {
    unitConverter;
    constructor(unitConverter) {
        this.unitConverter = unitConverter ?? new UnitConverter_js_1.UnitConverter();
    }
    /**
     * Parses an xfrm element from shape properties.
     * @param spPr Shape properties node containing a:xfrm
     * @returns Parsed transform or undefined if no transform found
     */
    parseTransform(spPr) {
        if (!spPr)
            return undefined;
        const xfrm = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:xfrm');
        if (!xfrm)
            return undefined;
        return this.parseXfrmElement(xfrm);
    }
    /**
     * Parses an xfrm element directly.
     * @param xfrm The a:xfrm element node
     * @returns Parsed transform
     */
    parseXfrmElement(xfrm) {
        // Get offset (position)
        const off = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:off');
        const x = off ? parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'x') ?? '0', 10) : 0;
        const y = off ? parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'y') ?? '0', 10) : 0;
        // Get extent (size)
        const ext = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:ext');
        const width = ext ? parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cx') ?? '0', 10) : 0;
        const height = ext ? parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cy') ?? '0', 10) : 0;
        // Get rotation (in 60000ths of a degree)
        const rotAttr = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'rot');
        const rotationUnits = rotAttr ? parseInt(rotAttr, 10) : 0;
        const rotation = rotationUnits / UnitConverter_js_1.ANGLE_UNIT_PER_DEGREE;
        // Get flip flags
        const flipH = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipH') === '1';
        const flipV = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipV') === '1';
        return {
            x,
            y,
            width,
            height,
            rotation,
            flipH,
            flipV,
        };
    }
    /**
     * Converts EMU transform to pixel transform.
     * @param transform EMU-based transform
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     * @returns Pixel-based transform
     */
    toPixelTransform(transform, scaleX, scaleY) {
        return {
            x: this.unitConverter.emuToPixels(transform.x) * scaleX,
            y: this.unitConverter.emuToPixels(transform.y) * scaleY,
            width: this.unitConverter.emuToPixels(transform.width) * scaleX,
            height: this.unitConverter.emuToPixels(transform.height) * scaleY,
            rotation: (transform.rotation * Math.PI) / 180,
            flipH: transform.flipH,
            flipV: transform.flipV,
        };
    }
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
    applyTransform(ctx, transform) {
        const centerX = transform.x + transform.width / 2;
        const centerY = transform.y + transform.height / 2;
        // Move to center
        ctx.translate(centerX, centerY);
        // Apply rotation
        if (transform.rotation !== 0) {
            ctx.rotate(transform.rotation);
        }
        // Apply flips
        if (transform.flipH || transform.flipV) {
            ctx.scale(transform.flipH ? -1 : 1, transform.flipV ? -1 : 1);
        }
        // Move back to top-left corner
        ctx.translate(-transform.width / 2, -transform.height / 2);
    }
    /**
     * Resets the canvas context to before the transform was applied.
     *
     * **Note:** This is a convenience wrapper around ctx.restore(). Since applyTransform
     * does not call ctx.save(), the caller must ensure proper save/restore pairing.
     *
     * @param ctx Canvas 2D context
     * @deprecated Use ctx.restore() directly with your own save/restore management
     */
    resetTransform(ctx) {
        ctx.restore();
    }
    /**
     * Gets the bounding rectangle in pixel space for a transform.
     * Takes rotation into account for the axis-aligned bounding box.
     * @param transform Pixel-based transform
     * @returns Axis-aligned bounding rectangle
     */
    getBoundingRect(transform) {
        if (transform.rotation === 0) {
            return {
                x: transform.x,
                y: transform.y,
                width: transform.width,
                height: transform.height,
            };
        }
        // Calculate corners
        const corners = this.getTransformedCorners(transform);
        // Find bounding box
        const xs = corners.map((c) => c.x);
        const ys = corners.map((c) => c.y);
        const minX = Math.min(...xs);
        const minY = Math.min(...ys);
        const maxX = Math.max(...xs);
        const maxY = Math.max(...ys);
        return {
            x: minX,
            y: minY,
            width: maxX - minX,
            height: maxY - minY,
        };
    }
    /**
     * Gets the four corners of a transformed rectangle in pixel space.
     * @param transform Pixel-based transform
     * @returns Array of four corner points
     */
    getTransformedCorners(transform) {
        const centerX = transform.x + transform.width / 2;
        const centerY = transform.y + transform.height / 2;
        const halfWidth = transform.width / 2;
        const halfHeight = transform.height / 2;
        const cos = Math.cos(transform.rotation);
        const sin = Math.sin(transform.rotation);
        // Calculate corners relative to center, then rotate and translate
        const offsets = [
            { x: -halfWidth, y: -halfHeight }, // top-left
            { x: halfWidth, y: -halfHeight }, // top-right
            { x: halfWidth, y: halfHeight }, // bottom-right
            { x: -halfWidth, y: halfHeight }, // bottom-left
        ];
        return offsets.map((offset) => {
            // Apply flip
            const x = transform.flipH ? -offset.x : offset.x;
            const y = transform.flipV ? -offset.y : offset.y;
            // Rotate
            const rotatedX = x * cos - y * sin;
            const rotatedY = x * sin + y * cos;
            // Translate to final position
            return {
                x: centerX + rotatedX,
                y: centerY + rotatedY,
            };
        });
    }
    /**
     * Converts a ShapeTransform to ParsedTransform.
     * @param shapeTransform Shape transform from types
     * @returns Parsed transform
     */
    fromShapeTransform(shapeTransform) {
        return {
            x: shapeTransform.offX,
            y: shapeTransform.offY,
            width: shapeTransform.extCx,
            height: shapeTransform.extCy,
            rotation: shapeTransform.rotation ?? 0,
            flipH: shapeTransform.flipH ?? false,
            flipV: shapeTransform.flipV ?? false,
        };
    }
    /**
     * Converts a ParsedTransform to ShapeTransform.
     * @param parsed Parsed transform
     * @returns Shape transform for types
     */
    toShapeTransform(parsed) {
        return {
            offX: parsed.x,
            offY: parsed.y,
            extCx: parsed.width,
            extCy: parsed.height,
            rotation: parsed.rotation,
            flipH: parsed.flipH,
            flipV: parsed.flipV,
        };
    }
}
exports.TransformCalculator = TransformCalculator;
/**
 * Default transform calculator instance.
 */
exports.defaultTransformCalculator = new TransformCalculator();
