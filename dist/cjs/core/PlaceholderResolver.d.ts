import type { PptxXmlNode } from './PptxParser.js';
import type { PlaceholderType, PlaceholderReference } from '../types/index.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Resolved placeholder with inherited properties.
 */
export interface ResolvedPlaceholder {
    /** Placeholder type */
    type?: PlaceholderType;
    /** Placeholder index */
    idx?: number;
    /** Inherited shape properties from layout/master */
    inheritedShapeProps?: PptxXmlNode;
    /** Inherited text body from layout/master */
    inheritedTextBody?: PptxXmlNode;
    /** Inherited text styles from layout/master */
    inheritedTextStyle?: PptxXmlNode;
}
/**
 * Master text-style buckets from p:txStyles (ECMA-376 §19.3.1.52).
 */
export type MasterStyleBucket = 'title' | 'body' | 'other';
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
export declare function getMasterStyleBucket(type: PlaceholderType | undefined): MasterStyleBucket;
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
export declare class PlaceholderResolver {
    private readonly logger;
    constructor(logger?: ILogger);
    /**
     * Extracts placeholder reference from a shape element.
     */
    extractPlaceholderRef(shapeNode: PptxXmlNode): PlaceholderReference | undefined;
    /**
     * Resolves a placeholder by looking up inherited properties: the matching
     * layout placeholder's shape properties and text body, and the master
     * txStyles bucket for the placeholder's type.
     */
    resolvePlaceholder(placeholderRef: PlaceholderReference, _slideNode: PptxXmlNode, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode): ResolvedPlaceholder;
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
    resolveTextStyleChain(placeholderRef: PlaceholderReference, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode, defaultTextStyle?: PptxXmlNode): PptxXmlNode[];
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
    findPlaceholderShape(containerNode: PptxXmlNode, type: PlaceholderType | undefined, idx: number | undefined): PptxXmlNode | undefined;
    /**
     * Gets the default text style bucket for a placeholder type from the slide
     * master's p:txStyles element.
     *
     * @param type Placeholder type (undefined defaults to body)
     * @param masterNode Slide master node (p:sldMaster)
     * @returns The style bucket node (contains a:lvl1pPr..a:lvl9pPr), or undefined
     */
    getDefaultTextStyle(type: PlaceholderType | undefined, masterNode?: PptxXmlNode): PptxXmlNode | undefined;
    /**
     * Merges shape properties with inherited properties.
     * Local properties take precedence; inheritance of individual shape
     * properties beyond the transform (handled elsewhere) is not yet modeled.
     */
    mergeShapeProperties(localProps: PptxXmlNode | undefined, inheritedProps: PptxXmlNode | undefined): PptxXmlNode | undefined;
    /**
     * Determines if a shape should be visible based on placeholder rules.
     */
    isPlaceholderVisible(placeholderRef: PlaceholderReference, hasContent: boolean): boolean;
    /**
     * Reads the placeholder type attribute from a shape node, if present.
     */
    private getShapePlaceholderType;
}
