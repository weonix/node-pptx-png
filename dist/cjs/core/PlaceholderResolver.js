"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PlaceholderResolver = void 0;
exports.getMasterStyleBucket = getMasterStyleBucket;
const PptxParser_js_1 = require("./PptxParser.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Maps a placeholder type to the master p:txStyles bucket that provides its
 * base text styling (ECMA-376 §19.3.1.36 p:ph and §19.3.1.52 p:txStyles):
 * - title / ctrTitle use p:titleStyle
 * - dt / ftr / sldNum use p:otherStyle
 * - everything else (body, subTitle, obj, and untyped placeholders, which
 *   default to body) uses p:bodyStyle
 *
 * @param type Placeholder type (undefined defaults to body per §19.3.1.36)
 * @returns The master style bucket
 */
function getMasterStyleBucket(type) {
    switch (type) {
        case 'title':
        case 'ctrTitle':
            return 'title';
        case 'dt':
        case 'ftr':
        case 'sldNum':
            return 'other';
        default:
            return 'body';
    }
}
/**
 * Normalizes a placeholder type for matching between slide and layout:
 * ctrTitle is equivalent to title, and a missing type defaults to body.
 */
function normalizePlaceholderType(type) {
    if (type === undefined)
        return 'body';
    return type === 'ctrTitle' ? 'title' : type;
}
/**
 * Gets a child element, preserving the presence of empty elements. The XML
 * parser yields self-closed childless elements (e.g. a bare `<p:ph/>`) as
 * empty strings, which must not be confused with a missing element.
 */
function getPresentChild(node, name) {
    if (!node || typeof node !== 'object')
        return undefined;
    const child = node[name];
    if (child === undefined)
        return undefined;
    return typeof child === 'object' && child !== null ? child : {};
}
/**
 * Resolves placeholder inheritance across the slide -> layout -> master chain.
 *
 * The text-style inheritance chain per ECMA-376 is, from lowest to highest
 * precedence: presentation p:defaultTextStyle -> master p:txStyles bucket ->
 * master placeholder a:lstStyle -> layout placeholder a:lstStyle -> slide
 * placeholder a:lstStyle -> explicit paragraph pPr -> run rPr. This class
 * resolves the inherited (presentation + master + layout) portion of the
 * chain; the slide-local portion is merged by TextParser.
 */
class PlaceholderResolver {
    logger;
    constructor(logger) {
        this.logger = logger ?? (0, Logger_js_1.createLogger)('warn', 'PlaceholderResolver');
    }
    /**
     * Extracts placeholder reference from a shape element.
     */
    extractPlaceholderRef(shapeNode) {
        // Look for nvSpPr -> nvPr -> ph
        const nvSpPr = (0, PptxParser_js_1.getXmlChild)(shapeNode, 'p:nvSpPr');
        if (!nvSpPr)
            return undefined;
        const nvPr = (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr');
        if (!nvPr)
            return undefined;
        const ph = getPresentChild(nvPr, 'p:ph');
        if (!ph)
            return undefined;
        const type = (0, PptxParser_js_1.getXmlAttr)(ph, 'type');
        const idx = (0, PptxParser_js_1.getXmlAttr)(ph, 'idx');
        return {
            type,
            idx: idx !== undefined ? parseInt(idx, 10) : undefined,
        };
    }
    /**
     * Resolves a placeholder by looking up inherited properties: the matching
     * layout placeholder's shape properties and text body, and the master
     * txStyles bucket for the placeholder's type.
     */
    resolvePlaceholder(placeholderRef, _slideNode, layoutNode, masterNode) {
        const layoutPh = layoutNode
            ? this.findPlaceholderShape(layoutNode, placeholderRef.type, placeholderRef.idx)
            : undefined;
        const effectiveType = placeholderRef.type ?? this.getShapePlaceholderType(layoutPh);
        return {
            type: placeholderRef.type,
            idx: placeholderRef.idx,
            inheritedShapeProps: layoutPh ? (0, PptxParser_js_1.getXmlChild)(layoutPh, 'p:spPr') : undefined,
            inheritedTextBody: layoutPh ? (0, PptxParser_js_1.getXmlChild)(layoutPh, 'p:txBody') : undefined,
            inheritedTextStyle: this.getDefaultTextStyle(effectiveType, masterNode),
        };
    }
    /**
     * Resolves the inherited list-style chain for a placeholder, ordered from
     * lowest to highest precedence:
     * 1. The presentation part's p:defaultTextStyle (ECMA-376 §19.2.1.12), the
     *    document-wide base every other layer overrides
     * 2. The master p:txStyles bucket for the placeholder type (title/ctrTitle
     *    -> p:titleStyle; dt/ftr/sldNum -> p:otherStyle; otherwise p:bodyStyle)
     * 3. The matching master placeholder shape's own p:txBody a:lstStyle
     * 4. The matching layout placeholder's a:lstStyle
     *
     * Each returned node contains a:lvl1pPr..a:lvl9pPr children; TextParser
     * merges the paragraph-level properties per level on top of this chain.
     *
     * @param placeholderRef The slide shape's placeholder reference
     * @param layoutNode Slide layout node (p:sldLayout)
     * @param masterNode Slide master node (p:sldMaster)
     * @param defaultTextStyle The presentation part's p:defaultTextStyle node
     * @returns List-style source nodes, lowest precedence first
     */
    resolveTextStyleChain(placeholderRef, layoutNode, masterNode, defaultTextStyle) {
        const chain = [];
        if (defaultTextStyle) {
            chain.push(defaultTextStyle);
        }
        const layoutPh = layoutNode
            ? this.findPlaceholderShape(layoutNode, placeholderRef.type, placeholderRef.idx)
            : undefined;
        // The slide's ph may omit type (idx-only body placeholders are common);
        // fall back to the matched layout placeholder's type for bucket selection.
        const effectiveType = placeholderRef.type ?? this.getShapePlaceholderType(layoutPh);
        const masterStyle = this.getDefaultTextStyle(effectiveType, masterNode);
        if (masterStyle) {
            chain.push(masterStyle);
        }
        // The master placeholder shape's own lstStyle sits between the master
        // txStyles bucket and the layout lstStyle. Matching is by type only:
        // idx values relate a slide ph to its LAYOUT ph (§19.3.1.36), while the
        // layout -> master correspondence is type-based, and master idx values
        // (dt=2/ftr=3/sldNum=4 in stock masters) would collide with layout
        // content-placeholder indices.
        // Body-bucket types without their own master placeholder (subTitle,
        // obj, pic, tbl, ...) inherit from the master's body placeholder.
        let masterPh = masterNode
            ? this.findPlaceholderShape(masterNode, effectiveType, undefined)
            : undefined;
        if (!masterPh &&
            masterNode &&
            effectiveType !== 'body' &&
            getMasterStyleBucket(effectiveType) === 'body') {
            masterPh = this.findPlaceholderShape(masterNode, 'body', undefined);
        }
        const masterTxBody = masterPh ? (0, PptxParser_js_1.getXmlChild)(masterPh, 'p:txBody') : undefined;
        const masterLstStyle = masterTxBody ? (0, PptxParser_js_1.getXmlChild)(masterTxBody, 'a:lstStyle') : undefined;
        if (masterLstStyle) {
            chain.push(masterLstStyle);
        }
        const layoutTxBody = layoutPh ? (0, PptxParser_js_1.getXmlChild)(layoutPh, 'p:txBody') : undefined;
        const layoutLstStyle = layoutTxBody ? (0, PptxParser_js_1.getXmlChild)(layoutTxBody, 'a:lstStyle') : undefined;
        if (layoutLstStyle) {
            chain.push(layoutLstStyle);
        }
        this.logger.debug('Resolved text style chain', {
            type: placeholderRef.type,
            idx: placeholderRef.idx,
            effectiveType,
            hasDefaultTextStyle: !!defaultTextStyle,
            hasMasterStyle: !!masterStyle,
            hasMasterPhLstStyle: !!masterLstStyle,
            hasLayoutLstStyle: !!layoutLstStyle,
        });
        return chain;
    }
    /**
     * Finds a matching placeholder shape in a layout/master node.
     *
     * Matching follows PowerPoint's rules: match by idx first (when the
     * reference carries one), then by type, treating ctrTitle as equivalent to
     * title and a missing type as body.
     *
     * @param containerNode Layout or master root node (p:sldLayout/p:sldMaster)
     * @param type Placeholder type from the referencing shape
     * @param idx Placeholder index from the referencing shape
     * @returns The matching p:sp node, or undefined
     */
    findPlaceholderShape(containerNode, type, idx) {
        const cSld = (0, PptxParser_js_1.getXmlChild)(containerNode, 'p:cSld');
        const spTree = cSld ? (0, PptxParser_js_1.getXmlChild)(cSld, 'p:spTree') : undefined;
        if (!spTree)
            return undefined;
        const shapes = (0, PptxParser_js_1.getXmlChildren)(spTree, 'p:sp');
        const candidates = [];
        for (const shape of shapes) {
            const nvSpPr = (0, PptxParser_js_1.getXmlChild)(shape, 'p:nvSpPr');
            const nvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr') : undefined;
            const ph = getPresentChild(nvPr, 'p:ph');
            if (!ph)
                continue;
            const phIdxAttr = (0, PptxParser_js_1.getXmlAttr)(ph, 'idx');
            candidates.push({
                shape,
                phType: (0, PptxParser_js_1.getXmlAttr)(ph, 'type'),
                phIdx: phIdxAttr !== undefined ? parseInt(phIdxAttr, 10) : undefined,
            });
        }
        // 1. Match by idx
        if (idx !== undefined) {
            const idxMatch = candidates.find((c) => c.phIdx === idx);
            if (idxMatch)
                return idxMatch.shape;
        }
        // 2. Match by (normalized) type
        const normalizedType = normalizePlaceholderType(type);
        const typeMatch = candidates.find((c) => normalizePlaceholderType(c.phType) === normalizedType);
        return typeMatch?.shape;
    }
    /**
     * Gets the default text style bucket for a placeholder type from the slide
     * master's p:txStyles element.
     *
     * @param type Placeholder type (undefined defaults to body)
     * @param masterNode Slide master node (p:sldMaster)
     * @returns The style bucket node (contains a:lvl1pPr..a:lvl9pPr), or undefined
     */
    getDefaultTextStyle(type, masterNode) {
        if (!masterNode)
            return undefined;
        const txStyles = (0, PptxParser_js_1.getXmlChild)(masterNode, 'p:txStyles');
        if (!txStyles)
            return undefined;
        switch (getMasterStyleBucket(type)) {
            case 'title':
                return (0, PptxParser_js_1.getXmlChild)(txStyles, 'p:titleStyle');
            case 'other':
                return (0, PptxParser_js_1.getXmlChild)(txStyles, 'p:otherStyle');
            default:
                return (0, PptxParser_js_1.getXmlChild)(txStyles, 'p:bodyStyle');
        }
    }
    /**
     * Merges shape properties with inherited properties.
     * Local properties take precedence; inheritance of individual shape
     * properties beyond the transform (handled elsewhere) is not yet modeled.
     */
    mergeShapeProperties(localProps, inheritedProps) {
        return localProps ?? inheritedProps;
    }
    /**
     * Determines if a shape should be visible based on placeholder rules.
     */
    isPlaceholderVisible(placeholderRef, hasContent) {
        // Empty placeholders without content are typically not rendered
        // unless they have a custom prompt
        if (!hasContent && !placeholderRef.hasCustomPrompt) {
            // Date, footer, slide number placeholders may still render if configured
            if (placeholderRef.type === 'dt' ||
                placeholderRef.type === 'ftr' ||
                placeholderRef.type === 'sldNum') {
                // These would be controlled by slide/presentation settings
                // For now, return false
                return false;
            }
            return false;
        }
        return true;
    }
    /**
     * Reads the placeholder type attribute from a shape node, if present.
     */
    getShapePlaceholderType(shapeNode) {
        if (!shapeNode)
            return undefined;
        const nvSpPr = (0, PptxParser_js_1.getXmlChild)(shapeNode, 'p:nvSpPr');
        const nvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr') : undefined;
        const ph = getPresentChild(nvPr, 'p:ph');
        return ph ? (0, PptxParser_js_1.getXmlAttr)(ph, 'type') : undefined;
    }
}
exports.PlaceholderResolver = PlaceholderResolver;
