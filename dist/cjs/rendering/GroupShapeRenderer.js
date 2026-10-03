"use strict";
/**
 * Renders group shapes (p:grpSp elements) to canvas.
 * Handles nested transforms, recursive child rendering, and group-level clipping.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.GroupShapeRenderer = void 0;
exports.createGroupShapeRenderer = createGroupShapeRenderer;
const PptxParser_js_1 = require("../core/PptxParser.js");
const TransformCalculator_js_1 = require("../geometry/TransformCalculator.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const constants_js_1 = require("../core/constants.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Renders group shapes (p:grpSp) to canvas.
 */
class GroupShapeRenderer {
    logger;
    scaleX;
    scaleY;
    transformCalculator;
    unitConverter;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'GroupShapeRenderer');
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
        this.transformCalculator = new TransformCalculator_js_1.TransformCalculator();
        this.unitConverter = new UnitConverter_js_1.UnitConverter();
    }
    /**
     * Parses the group transform from a p:grpSp element.
     * @param grpSpNode The p:grpSp XML node
     * @returns Group transform data or undefined if invalid
     */
    parseGroupTransform(grpSpNode) {
        // Get group shape properties
        const grpSpPr = (0, PptxParser_js_1.getXmlChild)(grpSpNode, 'p:grpSpPr');
        if (!grpSpPr) {
            this.logger.debug('Group shape has no grpSpPr');
            return undefined;
        }
        // Get the group transform (a:xfrm)
        const xfrm = (0, PptxParser_js_1.getXmlChild)(grpSpPr, 'a:xfrm');
        if (!xfrm) {
            this.logger.debug('Group shape has no xfrm');
            return undefined;
        }
        // Parse group offset and extent
        const off = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:off');
        const ext = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:ext');
        if (!off || !ext) {
            this.logger.debug('Group shape missing offset or extent');
            return undefined;
        }
        const groupX = parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'x') ?? '0', 10);
        const groupY = parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'y') ?? '0', 10);
        const groupCx = parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cx') ?? '0', 10);
        const groupCy = parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cy') ?? '0', 10);
        // Parse rotation and flip
        const rotation = parseInt((0, PptxParser_js_1.getXmlAttr)(xfrm, 'rot') ?? '0', 10) / 60000; // Convert from 60000ths
        const flipH = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipH') === '1';
        const flipV = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipV') === '1';
        const groupBounds = {
            x: groupX,
            y: groupY,
            width: groupCx,
            height: groupCy,
            rotation,
            flipH,
            flipV,
        };
        // Parse child offset and extent (coordinate space for children)
        const chOff = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:chOff');
        const chExt = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:chExt');
        // Default child bounds to group bounds if not specified
        const childX = chOff ? parseInt((0, PptxParser_js_1.getXmlAttr)(chOff, 'x') ?? '0', 10) : groupX;
        const childY = chOff ? parseInt((0, PptxParser_js_1.getXmlAttr)(chOff, 'y') ?? '0', 10) : groupY;
        const childCx = chExt ? parseInt((0, PptxParser_js_1.getXmlAttr)(chExt, 'cx') ?? '0', 10) : groupCx;
        const childCy = chExt ? parseInt((0, PptxParser_js_1.getXmlAttr)(chExt, 'cy') ?? '0', 10) : groupCy;
        const childBounds = {
            x: childX,
            y: childY,
            width: childCx,
            height: childCy,
            rotation: 0,
            flipH: false,
            flipV: false,
        };
        // Calculate scale factors for transforming child coordinates to group coordinates
        const scaleX = childCx > 0 ? groupCx / childCx : 1;
        const scaleY = childCy > 0 ? groupCy / childCy : 1;
        this.logger.debug('Parsed group transform', {
            group: { x: groupX, y: groupY, cx: groupCx, cy: groupCy },
            child: { x: childX, y: childY, cx: childCx, cy: childCy },
            scale: { x: scaleX, y: scaleY },
        });
        return {
            groupBounds,
            childBounds,
            scaleX,
            scaleY,
        };
    }
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
    transformChildToSlide(childTransform, groupTransform) {
        const { groupBounds, childBounds, scaleX, scaleY } = groupTransform;
        // Transform child position from child coordinate space to group coordinate space
        // 1. Subtract child offset to get position relative to child origin
        // 2. Scale by group/child ratio
        // 3. Add group offset to get position in the group's unrotated frame
        const relX = childTransform.x - childBounds.x;
        const relY = childTransform.y - childBounds.y;
        const transformedWidth = childTransform.width * scaleX;
        const transformedHeight = childTransform.height * scaleY;
        // Child center within the group's unrotated, unflipped frame
        let centerX = groupBounds.x + relX * scaleX + transformedWidth / 2;
        let centerY = groupBounds.y + relY * scaleY + transformedHeight / 2;
        const groupCenterX = groupBounds.x + groupBounds.width / 2;
        const groupCenterY = groupBounds.y + groupBounds.height / 2;
        // Group flips mirror the child's position within the group
        if (groupBounds.flipH) {
            centerX = 2 * groupCenterX - centerX;
        }
        if (groupBounds.flipV) {
            centerY = 2 * groupCenterY - centerY;
        }
        // Group rotation rotates the child's center around the group center
        // (positive angles are clockwise in the y-down slide coordinate space)
        if (groupBounds.rotation !== 0) {
            const angleRad = (groupBounds.rotation * Math.PI) / 180;
            const cos = Math.cos(angleRad);
            const sin = Math.sin(angleRad);
            const dx = centerX - groupCenterX;
            const dy = centerY - groupCenterY;
            centerX = groupCenterX + dx * cos - dy * sin;
            centerY = groupCenterY + dx * sin + dy * cos;
        }
        // Combine rotations and flips. Mirroring on exactly one axis reverses the
        // direction of the child's own rotation.
        const groupMirrored = groupBounds.flipH !== groupBounds.flipV;
        const childRotation = groupMirrored ? -childTransform.rotation : childTransform.rotation;
        const combinedRotation = groupBounds.rotation + childRotation;
        const combinedFlipH = groupBounds.flipH !== childTransform.flipH;
        const combinedFlipV = groupBounds.flipV !== childTransform.flipV;
        return {
            x: centerX - transformedWidth / 2,
            y: centerY - transformedHeight / 2,
            width: transformedWidth,
            height: transformedHeight,
            rotation: combinedRotation,
            flipH: combinedFlipH,
            flipV: combinedFlipV,
        };
    }
    /**
     * Gets the ordered child elements from a group shape.
     * Uses proper document order parsing to preserve z-order.
     * @param grpSpNode The p:grpSp XML node
     * @returns Array of ordered child elements in document order
     */
    getOrderedChildren(grpSpNode) {
        // Use getChildrenInDocumentOrder to preserve z-order across mixed element types
        const orderedElements = (0, PptxParser_js_1.getChildrenInDocumentOrder)(grpSpNode, 'p:grpSp', constants_js_1.SHAPE_ELEMENT_TYPES);
        // Check if we have multiple element types and log a warning if z-order may be affected
        const elementTypesPresent = new Set();
        for (const { tagName } of orderedElements) {
            elementTypesPresent.add(tagName);
        }
        if (elementTypesPresent.size > 1) {
            this.logger.debug('Group contains multiple element types - z-order preserved via ordered parsing', {
                elementTypes: Array.from(elementTypesPresent),
                count: orderedElements.length,
            });
        }
        // Convert to the expected return format
        return orderedElements.map(({ tagName, node }) => ({ tagName, node }));
    }
    /**
     * Renders a group shape and all its children.
     * @param ctx Canvas 2D context
     * @param grpSpNode The p:grpSp XML node
     * @param renderChild Callback function to render individual child elements
     * @param parentGroupTransform Optional parent group transform for nested groups
     */
    async renderGroupShape(ctx, grpSpNode, renderChild, parentGroupTransform) {
        // Check if hidden
        const nvGrpSpPr = (0, PptxParser_js_1.getXmlChild)(grpSpNode, 'p:nvGrpSpPr');
        const cNvPr = nvGrpSpPr ? (0, PptxParser_js_1.getXmlChild)(nvGrpSpPr, 'p:cNvPr') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden) {
            this.logger.debug('Group shape is hidden, skipping');
            return;
        }
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        this.logger.debug('Rendering group shape', { id, name });
        // Parse group transform
        let groupTransform = this.parseGroupTransform(grpSpNode);
        if (!groupTransform) {
            this.logger.warn('Could not parse group transform, skipping', { id });
            return;
        }
        // If this group is nested inside another group, apply the parent transform.
        // The effective scale factors must be recomputed from the transformed bounds
        // so the parent group's scaling is inherited by this group's children.
        if (parentGroupTransform) {
            const transformedBounds = this.transformChildToSlide(groupTransform.groupBounds, parentGroupTransform);
            const { childBounds } = groupTransform;
            groupTransform = {
                ...groupTransform,
                groupBounds: transformedBounds,
                scaleX: childBounds.width > 0
                    ? transformedBounds.width / childBounds.width
                    : parentGroupTransform.scaleX,
                scaleY: childBounds.height > 0
                    ? transformedBounds.height / childBounds.height
                    : parentGroupTransform.scaleY,
            };
        }
        // Get all children in order
        const children = this.getOrderedChildren(grpSpNode);
        this.logger.debug('Group has children', { count: children.length });
        // Render each child
        for (const { tagName, node } of children) {
            try {
                await renderChild(ctx, tagName, node, groupTransform);
            }
            catch (error) {
                this.logger.warn('Failed to render group child', {
                    tagName,
                    error: error instanceof Error ? error.message : String(error),
                });
            }
        }
    }
    /**
     * Parses a shape's transform from its spPr element.
     * @param spPr The shape properties node
     * @returns Parsed transform or undefined
     */
    parseShapeTransform(spPr) {
        return this.transformCalculator.parseTransform(spPr);
    }
    /**
     * Converts an EMU transform to pixel transform, applying group scaling.
     * @param transform EMU-based transform
     * @param groupTransform Optional group transform for nested shapes
     * @returns Pixel-based transform
     */
    toPixelTransform(transform, groupTransform) {
        // If there's a group transform, apply it first
        const finalTransform = groupTransform
            ? this.transformChildToSlide(transform, groupTransform)
            : transform;
        return this.transformCalculator.toPixelTransform(finalTransform, this.scaleX, this.scaleY);
    }
}
exports.GroupShapeRenderer = GroupShapeRenderer;
/**
 * Creates a GroupShapeRenderer instance.
 */
function createGroupShapeRenderer(scaleX, scaleY, logger) {
    return new GroupShapeRenderer({
        scaleX,
        scaleY,
        logger,
    });
}
