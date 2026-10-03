/**
 * Renders group shapes (p:grpSp elements) to canvas.
 * Handles nested transforms, recursive child rendering, and group-level clipping.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { PptxXmlNode } from '../core/PptxParser.js';
import { type ParsedTransform, type PixelTransform } from '../geometry/TransformCalculator.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Group transform data parsed from p:grpSpPr.
 */
export interface GroupTransform {
    /** Group position and size in EMU */
    groupBounds: ParsedTransform;
    /** Child coordinate space offset and size in EMU */
    childBounds: ParsedTransform;
    /** Scale factor for child X coordinates */
    scaleX: number;
    /** Scale factor for child Y coordinates */
    scaleY: number;
}
/**
 * Configuration for GroupShapeRenderer.
 */
export interface GroupShapeRendererConfig {
    /** Horizontal scale factor (EMU to pixels) */
    scaleX: number;
    /** Vertical scale factor (EMU to pixels) */
    scaleY: number;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Callback function type for rendering child elements.
 * This allows the SlideRenderer to pass its rendering functions.
 */
export type ChildRenderCallback = (ctx: CanvasRenderingContext2D, tagName: string, node: PptxXmlNode, parentGroupTransform?: GroupTransform) => Promise<void>;
/**
 * Renders group shapes (p:grpSp) to canvas.
 */
export declare class GroupShapeRenderer {
    private readonly logger;
    private readonly scaleX;
    private readonly scaleY;
    private readonly transformCalculator;
    private readonly unitConverter;
    constructor(config: GroupShapeRendererConfig);
    /**
     * Parses the group transform from a p:grpSp element.
     * @param grpSpNode The p:grpSp XML node
     * @returns Group transform data or undefined if invalid
     */
    parseGroupTransform(grpSpNode: PptxXmlNode): GroupTransform | undefined;
    /**
     * Transforms a child shape's EMU coordinates to account for group transform.
     * Maps the child rectangle from the group's child coordinate space (chOff/chExt)
     * onto the group rectangle, then applies the group's flip (mirroring the child's
     * position within the group) and rotation (rotating the child's center around
     * the group center), matching PowerPoint's transform composition.
     * @param childTransform The child's transform in child coordinate space
     * @param groupTransform The group's transform data
     * @returns Transformed coordinates in slide coordinate space (EMU)
     */
    transformChildToSlide(childTransform: ParsedTransform, groupTransform: GroupTransform): ParsedTransform;
    /**
     * Gets the ordered child elements from a group shape.
     * Uses proper document order parsing to preserve z-order.
     * @param grpSpNode The p:grpSp XML node
     * @returns Array of ordered child elements in document order
     */
    getOrderedChildren(grpSpNode: PptxXmlNode): Array<{
        tagName: string;
        node: PptxXmlNode;
    }>;
    /**
     * Renders a group shape and all its children.
     * @param ctx Canvas 2D context
     * @param grpSpNode The p:grpSp XML node
     * @param renderChild Callback function to render individual child elements
     * @param parentGroupTransform Optional parent group transform for nested groups
     */
    renderGroupShape(ctx: CanvasRenderingContext2D, grpSpNode: PptxXmlNode, renderChild: ChildRenderCallback, parentGroupTransform?: GroupTransform): Promise<void>;
    /**
     * Parses a shape's transform from its spPr element.
     * @param spPr The shape properties node
     * @returns Parsed transform or undefined
     */
    parseShapeTransform(spPr: PptxXmlNode | undefined): ParsedTransform | undefined;
    /**
     * Converts an EMU transform to pixel transform, applying group scaling.
     * @param transform EMU-based transform
     * @param groupTransform Optional group transform for nested shapes
     * @returns Pixel-based transform
     */
    toPixelTransform(transform: ParsedTransform, groupTransform?: GroupTransform): PixelTransform;
}
/**
 * Creates a GroupShapeRenderer instance.
 */
export declare function createGroupShapeRenderer(scaleX: number, scaleY: number, logger?: ILogger): GroupShapeRenderer;
