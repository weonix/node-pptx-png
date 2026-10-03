"use strict";
/**
 * SmartArt diagram rendering via the pre-rendered drawing part.
 *
 * PowerPoint persists a pre-baked rendering of every SmartArt graphic as a
 * "diagram drawing" part (MS-ODRAWXML, dsp: namespace): the slide's
 * graphicFrame carries <dgm:relIds r:dm=.../> pointing at the diagram data
 * part (diagrams/dataN.xml), and the drawing part (diagrams/drawingN.xml) is
 * reachable either through the data part's own relationships or through the
 * <dsp:dataModelExt relId=.../> extension inside the data model, whose relId
 * resolves against the referencing (slide) part. The drawing part's
 * <dsp:spTree> holds plain DrawingML shapes (dsp:sp with spPr/style/txBody,
 * structurally identical to p:sp), so no diagram layout engine is needed.
 * dsp:grpSp groups nest exactly like p:grpSp (grpSpPr xfrm with
 * off/ext/chOff/chExt) and are flattened here; a dsp:txXfrm on a shape lays
 * its text out in a rectangle independent of the shape geometry.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.DiagramRenderer = exports.DIAGRAM_GRAPHIC_DATA_URI = void 0;
exports.createDiagramRenderer = createDiagramRenderer;
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const TransformCalculator_js_1 = require("../geometry/TransformCalculator.js");
const RelationshipParser_js_1 = require("../parsers/RelationshipParser.js");
const ShapeRenderer_js_1 = require("./ShapeRenderer.js");
const GroupShapeRenderer_js_1 = require("./GroupShapeRenderer.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * The a:graphicData uri identifying a SmartArt diagram graphic frame.
 */
exports.DIAGRAM_GRAPHIC_DATA_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
/**
 * Substring identifying diagram-drawing relationships. Matched as a suffix so
 * both the Microsoft type (schemas.microsoft.com/office/2007/relationships/
 * diagramDrawing) and any transitional variants resolve.
 */
const DIAGRAM_DRAWING_TYPE_SUFFIX = '/diagramDrawing';
/**
 * Maps dsp:sp child element names onto their p:sp equivalents so the shape
 * can be rendered by the standard ShapeRenderer. The inner content is plain
 * DrawingML (a: namespace) and needs no translation.
 */
const DSP_SHAPE_CHILD_TAGS = {
    'dsp:nvSpPr': 'p:nvSpPr',
    'dsp:spPr': 'p:spPr',
    'dsp:style': 'p:style',
    'dsp:txBody': 'p:txBody',
};
/**
 * Maps dsp:nvSpPr child element names onto their p:nvSpPr equivalents.
 */
const DSP_NV_CHILD_TAGS = {
    'dsp:cNvPr': 'p:cNvPr',
    'dsp:cNvSpPr': 'p:cNvSpPr',
};
/**
 * Maps dsp:grpSp child element names onto their p:grpSp equivalents so the
 * group transform (grpSpPr > a:xfrm with off/ext/chOff/chExt) can be parsed
 * by GroupShapeRenderer.
 */
const DSP_GROUP_CHILD_TAGS = {
    'dsp:grpSpPr': 'p:grpSpPr',
};
/**
 * Renders SmartArt diagrams by drawing the shapes of the pre-rendered
 * drawing part (dsp:spTree) through the standard ShapeRenderer.
 */
class DiagramRenderer {
    logger;
    theme;
    parser;
    slidePath;
    slideNode;
    layoutNode;
    masterNode;
    relationshipParser;
    unitConverter;
    transformCalculator;
    groupShapeRenderer;
    warnings;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'DiagramRenderer');
        this.theme = config.theme;
        this.parser = config.parser;
        this.slidePath = config.slidePath;
        this.warnings = config.warnings;
        this.slideNode = config.slideNode;
        this.layoutNode = config.layoutNode;
        this.masterNode = config.masterNode;
        this.relationshipParser = new RelationshipParser_js_1.RelationshipParser({
            parser: config.parser,
            logger: this.logger.child('Rels'),
        });
        this.unitConverter = new UnitConverter_js_1.UnitConverter();
        this.transformCalculator = new TransformCalculator_js_1.TransformCalculator();
        // Only the EMU-space helpers (parseGroupTransform/transformChildToSlide)
        // are used, so the pixel scale factors are never consulted.
        this.groupShapeRenderer = new GroupShapeRenderer_js_1.GroupShapeRenderer({
            scaleX: 1,
            scaleY: 1,
            logger: this.logger.child('Group'),
        });
    }
    /**
     * Renders a SmartArt diagram graphic frame.
     *
     * Diagram shape coordinates live in the drawing part's own EMU space,
     * anchored at the graphicFrame origin and 1:1 with the frame's EMU extents
     * (PowerPoint regenerates the drawing part whenever the frame is resized).
     * Rendering therefore translates to the frame's pixel origin and scales
     * shapes by pixels-per-EMU derived from the frame bounds, which also
     * absorbs any group scaling already applied to the bounds.
     *
     * Missing drawing parts (files saved by producers that never pre-render)
     * degrade gracefully: a warning is logged and nothing is drawn.
     *
     * @param ctx Canvas 2D context
     * @param graphicData The a:graphicData node containing dgm:relIds
     * @param bounds The frame bounds in canvas pixels
     * @param frameExtEmu The frame's own EMU extents (p:xfrm/a:ext cx/cy)
     */
    async renderDiagram(ctx, graphicData, bounds, frameExtEmu) {
        const relIds = (0, PptxParser_js_1.getXmlChild)(graphicData, 'dgm:relIds');
        const dataRelId = relIds ? (0, PptxParser_js_1.getXmlAttr)(relIds, 'r:dm') : undefined;
        if (!dataRelId) {
            this.logger.warn('Diagram frame has no dgm:relIds r:dm reference');
            this.warnings?.push({
                code: 'diagram-missing-part',
                message: 'SmartArt frame has no diagram data reference; diagram skipped',
            });
            return;
        }
        const dataPath = await this.relationshipParser.resolveRelationshipId(this.slidePath, dataRelId);
        if (!dataPath) {
            this.logger.warn('Could not resolve diagram data relationship', { id: dataRelId });
            this.warnings?.push({
                code: 'diagram-missing-part',
                message: `Could not resolve diagram data relationship "${dataRelId}"; diagram skipped`,
                detail: { relationshipId: dataRelId },
            });
            return;
        }
        const drawingPath = await this.resolveDrawingPart(dataPath);
        if (!drawingPath) {
            this.logger.warn('SmartArt has no pre-rendered drawing part, skipping diagram', {
                dataPath,
            });
            this.warnings?.push({
                code: 'diagram-missing-part',
                message: 'SmartArt has no pre-rendered drawing part; diagram skipped',
                detail: { dataPath },
            });
            return;
        }
        const nativeWidth = this.unitConverter.emuToPixels(frameExtEmu.cx);
        const nativeHeight = this.unitConverter.emuToPixels(frameExtEmu.cy);
        if (nativeWidth <= 0 || nativeHeight <= 0) {
            this.logger.debug('Diagram frame has degenerate EMU extents', frameExtEmu);
            return;
        }
        const shapes = await this.readDrawingShapes(drawingPath);
        if (shapes.length === 0) {
            this.logger.debug('Diagram drawing part has no shapes', { drawingPath });
            return;
        }
        // Pixels-per-EMU inside the frame; equals the slide scale unless the
        // frame was additionally scaled (e.g. as a group child).
        const scaleX = bounds.width / nativeWidth;
        const scaleY = bounds.height / nativeHeight;
        const shapeRenderer = new ShapeRenderer_js_1.ShapeRenderer({
            theme: this.theme,
            scaleX,
            scaleY,
            parser: this.parser,
            // Blip fills inside the drawing resolve against the drawing part rels
            sourcePath: drawingPath,
            slideNode: this.slideNode,
            layoutNode: this.layoutNode,
            masterNode: this.masterNode,
            logger: this.logger.child('Shape'),
            warnings: this.warnings,
        });
        this.logger.debug('Rendering diagram drawing part', {
            drawingPath,
            shapes: shapes.length,
        });
        ctx.save();
        ctx.translate(bounds.x, bounds.y);
        try {
            // Document order is z-order: later shapes render on top
            for (const entry of shapes) {
                try {
                    await this.renderDiagramShape(ctx, shapeRenderer, entry, scaleX, scaleY);
                }
                catch (error) {
                    this.logger.warn('Failed to render diagram shape', {
                        error: error instanceof Error ? error.message : String(error),
                    });
                    this.warnings?.push({
                        code: 'other',
                        message: `Failed to render a SmartArt diagram shape: ${error instanceof Error ? error.message : String(error)}`,
                    });
                }
            }
        }
        finally {
            ctx.restore();
        }
    }
    /**
     * Renders one collected diagram shape: composes the enclosing group's
     * child-space mapping into the shape transform, and honors a dsp:txXfrm
     * text rectangle by rendering the geometry and the text as separate passes
     * when it differs from the shape's own xfrm.
     */
    async renderDiagramShape(ctx, shapeRenderer, entry, scaleX, scaleY) {
        const pSp = this.toPShapeNode(entry.node);
        const spPr = (0, PptxParser_js_1.getXmlChild)(pSp, 'p:spPr');
        const shapeXfrm = this.transformCalculator.parseTransform(spPr);
        // Map the shape from the group's child coordinate space (chOff/chExt)
        // into the drawing part's own space before ShapeRenderer sees it.
        if (entry.groupTransform && shapeXfrm && spPr) {
            const composed = this.groupShapeRenderer.transformChildToSlide(shapeXfrm, entry.groupTransform);
            pSp['p:spPr'] = { ...spPr, 'a:xfrm': buildXfrmNode(composed) };
        }
        const txBody = (0, PptxParser_js_1.getXmlChild)(pSp, 'p:txBody');
        const textXfrm = txBody ? this.parseTextXfrm(entry, shapeXfrm) : undefined;
        if (!txBody || !textXfrm) {
            await shapeRenderer.renderShape(ctx, pSp);
            return;
        }
        // Separate text rectangle: draw the shape without its text, then lay the
        // text out in the txXfrm rectangle (with its own rotation/flip).
        delete pSp['p:txBody'];
        await shapeRenderer.renderShape(ctx, pSp);
        this.renderTextInRect(ctx, shapeRenderer, pSp, txBody, textXfrm, scaleX, scaleY);
    }
    /**
     * Parses a dsp:sp's text transform (dsp:txXfrm, MS-ODRAWXML). Returns the
     * drawing-space EMU rectangle to lay text out in (composed through the
     * enclosing group mapping), or undefined when the shape carries none, when
     * it is degenerate, or when it matches the shape xfrm (where the normal
     * in-shape text path is already correct).
     */
    parseTextXfrm(entry, shapeXfrm) {
        const txXfrmNode = (0, PptxParser_js_1.getXmlChild)(entry.node, 'dsp:txXfrm');
        if (!txXfrmNode || typeof txXfrmNode !== 'object') {
            return undefined;
        }
        const txXfrm = this.transformCalculator.parseXfrmElement(txXfrmNode);
        if (txXfrm.width <= 0 || txXfrm.height <= 0) {
            return undefined;
        }
        if (shapeXfrm && transformsEqual(txXfrm, shapeXfrm)) {
            return undefined;
        }
        return entry.groupTransform
            ? this.groupShapeRenderer.transformChildToSlide(txXfrm, entry.groupTransform)
            : txXfrm;
    }
    /**
     * Renders a diagram shape's text body inside an EMU rectangle independent
     * of the shape geometry (dsp:txXfrm), applying the rectangle's own
     * rotation/flip. The default text color follows ShapeRenderer's
     * precedence: the style fontRef color, then the fill-contrast heuristic.
     */
    renderTextInRect(ctx, shapeRenderer, pSp, txBody, textXfrm, scaleX, scaleY) {
        const textBody = shapeRenderer.textParser.parseTextBody(txBody);
        if (!textBody || textBody.paragraphs.length === 0) {
            return;
        }
        const style = shapeRenderer.parseShapeStyle(pSp);
        const spPr = (0, PptxParser_js_1.getXmlChild)(pSp, 'p:spPr');
        const fill = (spPr ? shapeRenderer.fillRenderer.parseFill(spPr) : undefined) ?? style?.fill;
        const fillColor = fill?.type === 'solid' ? fill.color : undefined;
        const defaultTextColor = style?.fontColor ??
            (fillColor ? shapeRenderer.textRenderer.getContrastingColor(fillColor) : undefined);
        const pixel = this.transformCalculator.toPixelTransform(textXfrm, scaleX, scaleY);
        ctx.save();
        try {
            this.transformCalculator.applyTransform(ctx, pixel);
            shapeRenderer.textRenderer.renderText(ctx, textBody, { x: 0, y: 0, width: pixel.width, height: pixel.height }, defaultTextColor);
        }
        finally {
            ctx.restore();
        }
    }
    /**
     * Resolves the diagram drawing part path for a diagram data part.
     *
     * Tries, in order:
     * 1. A diagramDrawing relationship in the data part's own rels
     *    (spec-conformant producers).
     * 2. The dsp:dataModelExt relId extension inside the data model, resolved
     *    against the data part rels and then the slide rels (PowerPoint stores
     *    the diagramDrawing relationship on the slide).
     *
     * @param dataPath Path of the diagram data part (e.g., ppt/diagrams/data1.xml)
     * @returns The drawing part path, or undefined when no drawing part exists
     */
    async resolveDrawingPart(dataPath) {
        // 1. The data part's own relationships
        const dataRels = await this.relationshipParser.getAllRelationships(dataPath);
        const drawingRel = dataRels.find((rel) => rel.type.endsWith(DIAGRAM_DRAWING_TYPE_SUFFIX));
        if (drawingRel) {
            const path = this.parser.resolvePath(dataPath, drawingRel.target);
            if (this.parser.fileExists(path)) {
                return path;
            }
            this.logger.debug('Diagram drawing target missing from package', { path });
        }
        // 2. The dsp:dataModelExt extension's relId
        const extRelId = await this.readDataModelExtRelId(dataPath);
        if (!extRelId) {
            return undefined;
        }
        for (const sourcePath of [dataPath, this.slidePath]) {
            const rel = await this.relationshipParser.getRelationship(sourcePath, extRelId);
            if (rel && this.isDiagramDrawingRel(rel)) {
                const path = this.parser.resolvePath(sourcePath, rel.target);
                if (this.parser.fileExists(path)) {
                    return path;
                }
                this.logger.debug('Diagram drawing target missing from package', { path });
            }
        }
        return undefined;
    }
    /**
     * Checks that a relationship is a diagram drawing relationship, so a
     * dataModelExt relId is never resolved through an unrelated relationship.
     */
    isDiagramDrawingRel(rel) {
        return rel.type.endsWith(DIAGRAM_DRAWING_TYPE_SUFFIX);
    }
    /**
     * Reads the relId attribute of the dsp:dataModelExt extension inside a
     * diagram data part (dgm:dataModel > dgm:extLst > a:ext > dsp:dataModelExt).
     * Returns undefined when the data model carries no such extension.
     */
    async readDataModelExtRelId(dataPath) {
        let ordered;
        try {
            ordered = await this.parser.readXmlOrdered(dataPath);
        }
        catch (error) {
            this.logger.debug('Could not read diagram data part', {
                dataPath,
                error: error instanceof Error ? error.message : String(error),
            });
            return undefined;
        }
        const dataModel = (0, PptxParser_js_1.getOrderedChildren)(ordered, ['dgm:dataModel'])[0]?.node;
        const extLst = dataModel ? (0, PptxParser_js_1.getXmlChild)(dataModel, 'dgm:extLst') : undefined;
        if (!extLst) {
            return undefined;
        }
        for (const ext of (0, PptxParser_js_1.getXmlChildren)(extLst, 'a:ext')) {
            const dataModelExt = (0, PptxParser_js_1.getXmlChild)(ext, 'dsp:dataModelExt');
            const relId = dataModelExt ? (0, PptxParser_js_1.getXmlAttr)(dataModelExt, 'relId') : undefined;
            if (relId) {
                return relId;
            }
        }
        return undefined;
    }
    /**
     * Reads the drawing part and returns its shapes in document order, with
     * dsp:grpSp subtrees flattened into shape entries carrying the composed
     * group child-space mapping. Flattening preserves document order, so
     * z-order is unchanged.
     */
    async readDrawingShapes(drawingPath) {
        let ordered;
        try {
            ordered = await this.parser.readXmlOrdered(drawingPath);
        }
        catch (error) {
            this.logger.warn('Could not read diagram drawing part', {
                drawingPath,
                error: error instanceof Error ? error.message : String(error),
            });
            this.warnings?.push({
                code: 'diagram-missing-part',
                message: `Could not read diagram drawing part "${drawingPath}"; diagram skipped`,
                detail: { drawingPath },
            });
            return [];
        }
        const spTree = this.navigateToSpTree(ordered);
        if (!spTree) {
            this.logger.debug('Diagram drawing part has no dsp:spTree', { drawingPath });
            return [];
        }
        const entries = [];
        for (const el of (0, PptxParser_js_1.getOrderedChildren)(spTree, ['dsp:sp', 'dsp:grpSp'])) {
            if (el.tagName === 'dsp:grpSp') {
                this.collectGroupShapes(el.node, undefined, entries);
            }
            else {
                entries.push({ node: el.node });
            }
        }
        return entries;
    }
    /**
     * Flattens a dsp:grpSp subtree into shape entries, composing each nesting
     * level's child-space mapping (chOff/chExt onto off/ext, plus the group
     * xfrm's rotation/flip) exactly like GroupShapeRenderer does for p:grpSp
     * trees.
     */
    collectGroupShapes(grpSpNode, parentTransform, out) {
        const nvGrpSpPr = (0, PptxParser_js_1.getXmlChild)(grpSpNode, 'dsp:nvGrpSpPr');
        const cNvPr = nvGrpSpPr ? (0, PptxParser_js_1.getXmlChild)(nvGrpSpPr, 'dsp:cNvPr') : undefined;
        if (cNvPr && (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1') {
            this.logger.debug('Diagram group is hidden, skipping');
            return;
        }
        const ownTransform = this.groupShapeRenderer.parseGroupTransform(remapKeys(grpSpNode, DSP_GROUP_CHILD_TAGS));
        if (!ownTransform) {
            // A group without an xfrm maps children 1:1 onto its parent space
            this.logger.debug('Diagram group has no transform, passing children through');
        }
        const groupTransform = ownTransform
            ? this.composeGroupTransform(ownTransform, parentTransform)
            : parentTransform;
        for (const el of (0, PptxParser_js_1.getChildrenInDocumentOrder)(grpSpNode, 'dsp:grpSp', ['dsp:sp', 'dsp:grpSp'])) {
            if (el.tagName === 'dsp:grpSp') {
                this.collectGroupShapes(el.node, groupTransform, out);
            }
            else {
                out.push({ node: el.node, groupTransform });
            }
        }
    }
    /**
     * Composes a nested group's transform with its parent's, mirroring
     * GroupShapeRenderer.renderGroupShape: the inner group's bounds are mapped
     * through the parent's child-space transform, and the effective child
     * scale factors are recomputed from the transformed bounds so the parent's
     * scaling is inherited by the inner group's children.
     */
    composeGroupTransform(own, parent) {
        if (!parent) {
            return own;
        }
        const transformedBounds = this.groupShapeRenderer.transformChildToSlide(own.groupBounds, parent);
        const { childBounds } = own;
        return {
            ...own,
            groupBounds: transformedBounds,
            scaleX: childBounds.width > 0 ? transformedBounds.width / childBounds.width : parent.scaleX,
            scaleY: childBounds.height > 0 ? transformedBounds.height / childBounds.height : parent.scaleY,
        };
    }
    /**
     * Navigates the ordered drawing XML (dsp:drawing > dsp:spTree) to the
     * spTree children array, preserving document order.
     */
    navigateToSpTree(orderedDrawing) {
        if (!Array.isArray(orderedDrawing)) {
            return undefined;
        }
        for (const element of orderedDrawing) {
            const drawingChildren = element?.['dsp:drawing'];
            if (!Array.isArray(drawingChildren)) {
                continue;
            }
            for (const child of drawingChildren) {
                const spTreeChildren = child['dsp:spTree'];
                if (Array.isArray(spTreeChildren)) {
                    return spTreeChildren;
                }
            }
        }
        return undefined;
    }
    /**
     * Rewraps a dsp:sp node as a p:sp-shaped node so ShapeRenderer can render
     * it unchanged. Child nodes are shared by reference, so ordered-source
     * metadata on spPr/txBody content (path segment order, run interleaving)
     * is preserved.
     */
    toPShapeNode(dspSp) {
        const shape = remapKeys(dspSp, DSP_SHAPE_CHILD_TAGS);
        const nvSpPr = shape['p:nvSpPr'];
        if (nvSpPr && typeof nvSpPr === 'object') {
            shape['p:nvSpPr'] = remapKeys(nvSpPr, DSP_NV_CHILD_TAGS);
        }
        return shape;
    }
}
exports.DiagramRenderer = DiagramRenderer;
/**
 * Returns a shallow copy of a node with the given child tag names renamed.
 * Unmapped keys (attributes, other children) are copied through unchanged.
 */
function remapKeys(node, mapping) {
    const result = {};
    for (const [key, value] of Object.entries(node)) {
        result[mapping[key] ?? key] = value;
    }
    return result;
}
/**
 * Builds a synthetic a:xfrm node from an EMU transform so a group-composed
 * child transform can flow through ShapeRenderer's normal parse path.
 * Values are rounded to whole EMU (sub-EMU error is far below a pixel).
 */
function buildXfrmNode(transform) {
    const xfrm = {
        'a:off': {
            '@_x': String(Math.round(transform.x)),
            '@_y': String(Math.round(transform.y)),
        },
        'a:ext': {
            '@_cx': String(Math.round(transform.width)),
            '@_cy': String(Math.round(transform.height)),
        },
    };
    if (transform.rotation !== 0) {
        xfrm['@_rot'] = String(Math.round(transform.rotation * UnitConverter_js_1.ANGLE_UNIT_PER_DEGREE));
    }
    if (transform.flipH) {
        xfrm['@_flipH'] = '1';
    }
    if (transform.flipV) {
        xfrm['@_flipV'] = '1';
    }
    return xfrm;
}
/**
 * Compares two EMU transforms for exact equality.
 */
function transformsEqual(a, b) {
    return (a.x === b.x &&
        a.y === b.y &&
        a.width === b.width &&
        a.height === b.height &&
        a.rotation === b.rotation &&
        a.flipH === b.flipH &&
        a.flipV === b.flipV);
}
/**
 * Creates a DiagramRenderer instance.
 */
function createDiagramRenderer(config) {
    return new DiagramRenderer(config);
}
