"use strict";
/**
 * Parses shape elements (p:sp, p:cxnSp) from slide XML.
 * Extracts structured shape data for rendering.
 *
 * Note: Fill and stroke parsing is handled by FillRenderer and StrokeRenderer
 * respectively. This parser extracts raw XML nodes for those components.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ShapeParser = void 0;
exports.createShapeParser = createShapeParser;
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const GeometryEngine_js_1 = require("../geometry/GeometryEngine.js");
const CustomGeometryParser_js_1 = require("../geometry/CustomGeometryParser.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Parses shape elements from slide XML.
 * Note: Fill and stroke are returned as raw XML nodes (spPrNode).
 * Use FillRenderer.parseFill() and StrokeRenderer.parseStroke() for actual parsing.
 */
class ShapeParser {
    logger;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'ShapeParser');
    }
    /**
     * Parses a shape element (p:sp) from XML.
     * @param spNode Shape XML node
     * @returns Parsed ShapeElement or undefined if invalid
     */
    parseShape(spNode) {
        // Parse non-visual properties
        const nvSpPr = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:nvSpPr');
        if (!nvSpPr) {
            this.logger.debug('Shape missing nvSpPr');
            return undefined;
        }
        const cNvPr = (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:cNvPr');
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        // Parse placeholder reference (parsed but not yet consumed downstream)
        const nvPr = (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr');
        const _placeholder = this.parsePlaceholder(nvPr);
        // Parse shape properties
        const spPr = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:spPr');
        if (!spPr) {
            this.logger.debug('Shape missing spPr', { id });
            return undefined;
        }
        // Parse transform
        const transform = this.parseTransform(spPr);
        if (!transform) {
            this.logger.debug('Shape missing transform', { id });
            return undefined;
        }
        // Parse geometry
        const { presetGeometry, customGeometry } = this.parseGeometry(spPr, transform);
        // Note: Fill and stroke parsing is delegated to FillRenderer and StrokeRenderer
        // The spPr node is available for external parsing if needed
        // Parse text body
        const txBody = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:txBody');
        const textBody = txBody ? this.parseTextBody(txBody) : undefined;
        return {
            type: 'shape',
            id,
            name,
            transform,
            presetGeometry,
            customGeometry,
            textBody,
            hidden,
        };
    }
    /**
     * Parses a placeholder reference from non-visual properties.
     */
    parsePlaceholder(nvPr) {
        if (!nvPr)
            return undefined;
        const ph = (0, PptxParser_js_1.getXmlChild)(nvPr, 'p:ph');
        if (!ph)
            return undefined;
        const typeAttr = (0, PptxParser_js_1.getXmlAttr)(ph, 'type');
        const idxAttr = (0, PptxParser_js_1.getXmlAttr)(ph, 'idx');
        return {
            type: typeAttr,
            idx: idxAttr !== undefined ? parseInt(idxAttr, 10) : undefined,
        };
    }
    /**
     * Parses transform (xfrm) from shape properties.
     */
    parseTransform(spPr) {
        const xfrm = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:xfrm');
        if (!xfrm)
            return undefined;
        const off = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:off');
        const ext = (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:ext');
        const offX = off ? parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'x') ?? '0', 10) : 0;
        const offY = off ? parseInt((0, PptxParser_js_1.getXmlAttr)(off, 'y') ?? '0', 10) : 0;
        const extCx = ext ? parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cx') ?? '0', 10) : 0;
        const extCy = ext ? parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cy') ?? '0', 10) : 0;
        const rotAttr = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'rot');
        const rotation = rotAttr ? parseInt(rotAttr, 10) / UnitConverter_js_1.ANGLE_UNIT_PER_DEGREE : undefined;
        const flipH = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipH') === '1';
        const flipV = (0, PptxParser_js_1.getXmlAttr)(xfrm, 'flipV') === '1';
        return {
            offX,
            offY,
            extCx,
            extCy,
            rotation,
            flipH,
            flipV,
        };
    }
    /**
     * Parses geometry from shape properties.
     *
     * Custom geometry (a:custGeom) is parsed in document order and evaluated
     * through the shared geometry engine (guide formulas, per-path coordinate
     * spaces, arcs). The resulting paths are in the shape's EMU coordinate
     * space (origin at the shape's top-left).
     */
    parseGeometry(spPr, transform) {
        // Check for preset geometry
        const prstGeom = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:prstGeom');
        if (prstGeom) {
            const prst = (0, PptxParser_js_1.getXmlAttr)(prstGeom, 'prst') ?? 'rect';
            return { presetGeometry: prst };
        }
        // Check for custom geometry
        const custGeom = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:custGeom');
        if (custGeom) {
            const definition = (0, CustomGeometryParser_js_1.parseCustomGeometryDefinition)(custGeom);
            if (definition) {
                try {
                    const space = { w: transform.extCx, h: transform.extCy };
                    const bounds = { x: 0, y: 0, width: transform.extCx, height: transform.extCy };
                    const customGeometry = (0, GeometryEngine_js_1.buildGeometryPaths)(definition, bounds, undefined, space);
                    if (customGeometry.length > 0) {
                        return { customGeometry };
                    }
                }
                catch (error) {
                    const message = error instanceof Error ? error.message : String(error);
                    this.logger.debug('Custom geometry evaluation failed', { error: message });
                }
            }
            this.logger.debug('Custom geometry could not be parsed, using rectangle');
            return { presetGeometry: 'rect' };
        }
        return { presetGeometry: 'rect' };
    }
    // Note: Fill and stroke parsing methods have been removed from ShapeParser.
    // Use FillRenderer.parseFill() and StrokeRenderer.parseStroke() for fill/stroke parsing.
    // This eliminates code duplication between the parser and renderers.
    /**
     * Parses text body from shape.
     */
    parseTextBody(txBody) {
        const paragraphs = [];
        const pNodes = (0, PptxParser_js_1.getXmlChildren)(txBody, 'a:p');
        for (const pNode of pNodes) {
            const paragraph = this.parseParagraph(pNode);
            if (paragraph) {
                paragraphs.push(paragraph);
            }
        }
        if (paragraphs.length === 0)
            return undefined;
        return { paragraphs };
    }
    /**
     * Parses a paragraph.
     */
    parseParagraph(pNode) {
        const runs = [];
        // Parse text runs
        const rNodes = (0, PptxParser_js_1.getXmlChildren)(pNode, 'a:r');
        for (const rNode of rNodes) {
            const run = this.parseTextRun(rNode);
            if (run) {
                runs.push(run);
            }
        }
        return { runs };
    }
    /**
     * Parses a text run.
     */
    parseTextRun(rNode) {
        const tNode = (0, PptxParser_js_1.getXmlChild)(rNode, 'a:t');
        if (!tNode)
            return undefined;
        // Get text content
        const text = typeof tNode === 'string' ? tNode : (tNode['#text'] ?? '');
        return { text };
    }
}
exports.ShapeParser = ShapeParser;
/**
 * Creates a ShapeParser instance.
 */
function createShapeParser(theme, logger) {
    return new ShapeParser({ theme, logger });
}
