/**
 * Parses shape elements (p:sp, p:cxnSp) from slide XML.
 * Extracts structured shape data for rendering.
 *
 * Note: Fill and stroke parsing is handled by FillRenderer and StrokeRenderer
 * respectively. This parser extracts raw XML nodes for those components.
 */
import { getXmlChild, getXmlAttr, getXmlChildren } from '../core/PptxParser.js';
import { ANGLE_UNIT_PER_DEGREE } from '../core/UnitConverter.js';
import { buildGeometryPaths } from '../geometry/GeometryEngine.js';
import { parseCustomGeometryDefinition } from '../geometry/CustomGeometryParser.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Parses shape elements from slide XML.
 * Note: Fill and stroke are returned as raw XML nodes (spPrNode).
 * Use FillRenderer.parseFill() and StrokeRenderer.parseStroke() for actual parsing.
 */
export class ShapeParser {
    logger;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'ShapeParser');
    }
    /**
     * Parses a shape element (p:sp) from XML.
     * @param spNode Shape XML node
     * @returns Parsed ShapeElement or undefined if invalid
     */
    parseShape(spNode) {
        // Parse non-visual properties
        const nvSpPr = getXmlChild(spNode, 'p:nvSpPr');
        if (!nvSpPr) {
            this.logger.debug('Shape missing nvSpPr');
            return undefined;
        }
        const cNvPr = getXmlChild(nvSpPr, 'p:cNvPr');
        const id = cNvPr ? (getXmlAttr(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? getXmlAttr(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? getXmlAttr(cNvPr, 'hidden') === '1' : false;
        // Parse placeholder reference (parsed but not yet consumed downstream)
        const nvPr = getXmlChild(nvSpPr, 'p:nvPr');
        const _placeholder = this.parsePlaceholder(nvPr);
        // Parse shape properties
        const spPr = getXmlChild(spNode, 'p:spPr');
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
        const txBody = getXmlChild(spNode, 'p:txBody');
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
        const ph = getXmlChild(nvPr, 'p:ph');
        if (!ph)
            return undefined;
        const typeAttr = getXmlAttr(ph, 'type');
        const idxAttr = getXmlAttr(ph, 'idx');
        return {
            type: typeAttr,
            idx: idxAttr !== undefined ? parseInt(idxAttr, 10) : undefined,
        };
    }
    /**
     * Parses transform (xfrm) from shape properties.
     */
    parseTransform(spPr) {
        const xfrm = getXmlChild(spPr, 'a:xfrm');
        if (!xfrm)
            return undefined;
        const off = getXmlChild(xfrm, 'a:off');
        const ext = getXmlChild(xfrm, 'a:ext');
        const offX = off ? parseInt(getXmlAttr(off, 'x') ?? '0', 10) : 0;
        const offY = off ? parseInt(getXmlAttr(off, 'y') ?? '0', 10) : 0;
        const extCx = ext ? parseInt(getXmlAttr(ext, 'cx') ?? '0', 10) : 0;
        const extCy = ext ? parseInt(getXmlAttr(ext, 'cy') ?? '0', 10) : 0;
        const rotAttr = getXmlAttr(xfrm, 'rot');
        const rotation = rotAttr ? parseInt(rotAttr, 10) / ANGLE_UNIT_PER_DEGREE : undefined;
        const flipH = getXmlAttr(xfrm, 'flipH') === '1';
        const flipV = getXmlAttr(xfrm, 'flipV') === '1';
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
        const prstGeom = getXmlChild(spPr, 'a:prstGeom');
        if (prstGeom) {
            const prst = getXmlAttr(prstGeom, 'prst') ?? 'rect';
            return { presetGeometry: prst };
        }
        // Check for custom geometry
        const custGeom = getXmlChild(spPr, 'a:custGeom');
        if (custGeom) {
            const definition = parseCustomGeometryDefinition(custGeom);
            if (definition) {
                try {
                    const space = { w: transform.extCx, h: transform.extCy };
                    const bounds = { x: 0, y: 0, width: transform.extCx, height: transform.extCy };
                    const customGeometry = buildGeometryPaths(definition, bounds, undefined, space);
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
        const pNodes = getXmlChildren(txBody, 'a:p');
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
        const rNodes = getXmlChildren(pNode, 'a:r');
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
        const tNode = getXmlChild(rNode, 'a:t');
        if (!tNode)
            return undefined;
        // Get text content
        const text = typeof tNode === 'string' ? tNode : (tNode['#text'] ?? '');
        return { text };
    }
}
/**
 * Creates a ShapeParser instance.
 */
export function createShapeParser(theme, logger) {
    return new ShapeParser({ theme, logger });
}
//# sourceMappingURL=ShapeParser.js.map