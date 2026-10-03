/**
 * Parses custom geometry (a:custGeom) XML into a GeometryDefinition that the
 * shared geometry engine can evaluate — the same representation used for
 * preset shapes, since custom geometry uses the identical drawing language
 * (ECMA-376 §20.1.9.8).
 *
 * Path segments are read in true document order via
 * getChildrenInDocumentOrder, which is order-preserving for nodes produced
 * by the ordered slide parse (interleaved moveTo/lnTo/arcTo sequences and
 * multiple subpaths are handled correctly).
 */
import { getChildrenInDocumentOrder, getXmlAttr, getXmlChild, getXmlChildren, } from '../core/PptxParser.js';
/** Path segment tags in the custom geometry drawing language. */
const PATH_SEGMENT_TAGS = [
    'a:moveTo',
    'a:lnTo',
    'a:cubicBezTo',
    'a:quadBezTo',
    'a:arcTo',
    'a:close',
];
/** Matches a numeric literal coordinate (ST_AdjCoordinate numeric branch). */
const NUMERIC_LITERAL = /^[-+]?\d+(?:\.\d+)?$/;
/**
 * Converts a raw attribute value to a GeomValue: a number for numeric
 * literals, otherwise the guide name string.
 */
function toGeomValue(raw) {
    if (raw === undefined)
        return undefined;
    const trimmed = raw.trim();
    if (trimmed.length === 0)
        return undefined;
    return NUMERIC_LITERAL.test(trimmed) ? parseFloat(trimmed) : trimmed;
}
/**
 * Extracts the x/y coordinate pairs from a segment's a:pt children.
 * Returns undefined if fewer than `expected` points are present.
 */
function extractPoints(node, expected) {
    const ptNodes = getXmlChildren(node, 'a:pt');
    if (ptNodes.length < expected)
        return undefined;
    const values = [];
    for (let i = 0; i < expected; i++) {
        const pt = ptNodes[i];
        const x = toGeomValue(getXmlAttr(pt, 'x'));
        const y = toGeomValue(getXmlAttr(pt, 'y'));
        if (x === undefined || y === undefined)
            return undefined;
        values.push(x, y);
    }
    return values;
}
/**
 * Parses a single path segment element into a path command.
 */
function parseSegment(tagName, node) {
    switch (tagName) {
        case 'a:moveTo': {
            const pts = extractPoints(node, 1);
            return pts ? ['M', pts[0], pts[1]] : undefined;
        }
        case 'a:lnTo': {
            const pts = extractPoints(node, 1);
            return pts ? ['L', pts[0], pts[1]] : undefined;
        }
        case 'a:quadBezTo': {
            const pts = extractPoints(node, 2);
            return pts ? ['Q', pts[0], pts[1], pts[2], pts[3]] : undefined;
        }
        case 'a:cubicBezTo': {
            const pts = extractPoints(node, 3);
            return pts ? ['C', pts[0], pts[1], pts[2], pts[3], pts[4], pts[5]] : undefined;
        }
        case 'a:arcTo': {
            const wR = toGeomValue(getXmlAttr(node, 'wR'));
            const hR = toGeomValue(getXmlAttr(node, 'hR'));
            const stAng = toGeomValue(getXmlAttr(node, 'stAng'));
            const swAng = toGeomValue(getXmlAttr(node, 'swAng'));
            if (wR === undefined || hR === undefined || stAng === undefined || swAng === undefined) {
                return undefined;
            }
            return ['A', wR, hR, stAng, swAng];
        }
        case 'a:close':
            return ['Z'];
        default:
            return undefined;
    }
}
/**
 * Parses a single a:path element into a path definition, iterating its
 * segments in document order.
 */
function parsePathDef(pathNode) {
    const cmds = [];
    const orderedSegments = getChildrenInDocumentOrder(pathNode, 'a:path', PATH_SEGMENT_TAGS);
    for (const segment of orderedSegments) {
        const cmd = parseSegment(segment.tagName, segment.node);
        if (cmd) {
            cmds.push(cmd);
        }
    }
    if (cmds.length === 0)
        return undefined;
    const pathDef = { cmds };
    const wAttr = getXmlAttr(pathNode, 'w');
    const hAttr = getXmlAttr(pathNode, 'h');
    if (wAttr !== undefined)
        pathDef.w = parseFloat(wAttr);
    if (hAttr !== undefined)
        pathDef.h = parseFloat(hAttr);
    const fillAttr = getXmlAttr(pathNode, 'fill');
    if (fillAttr !== undefined && fillAttr !== 'norm') {
        pathDef.fill = fillAttr;
    }
    const strokeAttr = getXmlAttr(pathNode, 'stroke');
    if (strokeAttr === 'false' || strokeAttr === '0') {
        pathDef.stroke = false;
    }
    return pathDef;
}
/**
 * Parses an a:custGeom node into a GeometryDefinition for the geometry
 * engine: adjust value defaults (a:avLst), guide formulas (a:gdLst, in
 * document order), the text rectangle (a:rect), and all paths (a:pathLst)
 * with per-path coordinate spaces and fill/stroke flags.
 *
 * @param custGeom The a:custGeom XML node
 * @returns The geometry definition, or undefined if no drawable path exists
 */
export function parseCustomGeometryDefinition(custGeom) {
    const pathLst = getXmlChild(custGeom, 'a:pathLst');
    if (!pathLst)
        return undefined;
    const paths = [];
    for (const pathNode of getXmlChildren(pathLst, 'a:path')) {
        const pathDef = parsePathDef(pathNode);
        if (pathDef) {
            paths.push(pathDef);
        }
    }
    if (paths.length === 0)
        return undefined;
    const def = { paths };
    // Adjust values: a:avLst entries are always "val N" formulas
    const avLst = getXmlChild(custGeom, 'a:avLst');
    if (avLst) {
        const av = [];
        for (const gd of getXmlChildren(avLst, 'a:gd')) {
            const name = getXmlAttr(gd, 'name');
            const fmla = getXmlAttr(gd, 'fmla');
            const match = fmla ? /^val\s+([-+]?\d+(?:\.\d+)?)$/.exec(fmla) : null;
            if (name && match?.[1]) {
                av.push([name, parseFloat(match[1])]);
            }
        }
        if (av.length > 0)
            def.av = av;
    }
    // Guide formulas in document order (same-tag children preserve order)
    const gdLst = getXmlChild(custGeom, 'a:gdLst');
    if (gdLst) {
        const gd = [];
        for (const gdNode of getXmlChildren(gdLst, 'a:gd')) {
            const name = getXmlAttr(gdNode, 'name');
            const fmla = getXmlAttr(gdNode, 'fmla');
            if (name && fmla) {
                gd.push([name, fmla]);
            }
        }
        if (gd.length > 0)
            def.gd = gd;
    }
    // Text rectangle
    const rectNode = getXmlChild(custGeom, 'a:rect');
    if (rectNode) {
        const l = toGeomValue(getXmlAttr(rectNode, 'l'));
        const t = toGeomValue(getXmlAttr(rectNode, 't'));
        const r = toGeomValue(getXmlAttr(rectNode, 'r'));
        const b = toGeomValue(getXmlAttr(rectNode, 'b'));
        if (l !== undefined && t !== undefined && r !== undefined && b !== undefined) {
            def.rect = [l, t, r, b];
        }
    }
    return def;
}
//# sourceMappingURL=CustomGeometryParser.js.map