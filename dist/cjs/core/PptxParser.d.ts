import type { ILogger } from '../utils/Logger.js';
/**
 * Represents an element in an ordered XML structure.
 * Used by the preserveOrder parser to maintain document order.
 */
export interface OrderedXmlElement {
    /** The tag name of the element */
    tagName: string;
    /** The element's attributes (prefixed with @_) */
    attributes: Record<string, string>;
    /** The element node itself */
    node: PptxXmlNode;
}
/**
 * Raw slide data extracted from PPTX.
 */
export interface SlideData {
    /** Slide index (0-based) */
    index: number;
    /** Slide XML content parsed as object */
    content: PptxXmlNode;
    /** Slide layout relationship ID */
    layoutRelId?: string;
    /** Path to the slide file within the PPTX */
    path: string;
}
/**
 * Raw slide layout data.
 */
export interface SlideLayoutData {
    /** Layout name */
    name?: string;
    /** Layout XML content */
    content: PptxXmlNode;
    /** Master relationship ID */
    masterRelId?: string;
    /** Path to the layout file */
    path: string;
}
/**
 * Raw slide master data.
 */
export interface SlideMasterData {
    /** Master name */
    name?: string;
    /** Master XML content */
    content: PptxXmlNode;
    /** Theme relationship ID */
    themeRelId?: string;
    /** Path to the master file */
    path: string;
}
/**
 * Theme data extracted from PPTX.
 */
export interface ThemeData {
    /** Theme XML content */
    content: PptxXmlNode;
    /** Path to the theme file */
    path: string;
}
/**
 * Presentation-level data.
 */
export interface PresentationData {
    /** Slide width in EMU */
    slideWidth: number;
    /** Slide height in EMU */
    slideHeight: number;
    /** Slide IDs in order */
    slideIds: string[];
    /** Number of slides */
    slideCount: number;
    /** Presentation XML content */
    content: PptxXmlNode;
}
/**
 * Relationship entry from .rels file.
 */
export interface Relationship {
    id: string;
    type: string;
    target: string;
}
/**
 * Style of an embedded font variant, matching the child elements of
 * p:embeddedFont (ECMA-376 §19.2.1.9: p:regular, p:bold, p:italic, p:boldItalic).
 */
export type EmbeddedFontVariantStyle = 'regular' | 'bold' | 'italic' | 'boldItalic';
/**
 * A single embedded font variant (one font part inside the package).
 */
export interface EmbeddedFontVariant {
    /** Which face this part carries (regular/bold/italic/boldItalic) */
    style: EmbeddedFontVariantStyle;
    /** Path of the font part within the PPTX (e.g., ppt/fonts/font1.fntdata) */
    path: string;
    /** Raw bytes of the font part (may be plain TTF/OTF or wrapped/obfuscated) */
    data: Buffer;
}
/**
 * An embedded font entry from p:embeddedFontLst with its extracted variants.
 */
export interface EmbeddedFontData {
    /** Typeface name declared by p:font/@typeface (e.g., 'Open Sans') */
    typeface: string;
    /** Extracted font variants in declaration order */
    variants: EmbeddedFontVariant[];
}
/**
 * Generic XML node type from fast-xml-parser.
 */
export type PptxXmlNode = Record<string, unknown>;
/**
 * Tag value processor that trims text values for all tags except those where
 * whitespace is significant. Used with `trimValues: false` so that a:t runs
 * keep their whitespace while inter-element whitespace (pretty-printed XML)
 * is still discarded for structural tags.
 */
export declare const trimTagValueExceptText: (tagName: string, tagText: string) => string;
/**
 * Attribute value processor that trims attribute values, preserving the
 * behavior `trimValues: true` previously provided for attributes.
 */
export declare const trimAttributeValue: (_attrName: string, attrValue: string) => string;
/**
 * Parser for PPTX files.
 * Handles ZIP extraction and XML parsing.
 *
 * **Caching Behavior:**
 * This parser maintains internal caches for parsed XML content and relationships
 * to avoid redundant parsing of the same files within a PPTX. The caches are:
 * - `xmlCache`: Caches parsed XML nodes by file path
 * - `relationshipCache`: Caches parsed relationship arrays by .rels file path
 *
 * **Important:** Caches are cleared when:
 * - A new PPTX file is opened via `open()`
 * - The parser is explicitly closed via `close()`
 *
 * **Lifecycle:** This class is designed to be short-lived, typically used for
 * a single rendering operation. Create a new instance for each PPTX file you
 * process, and call `close()` when done to release resources and clear caches.
 *
 * @example
 * ```typescript
 * const parser = new PptxParser();
 * try {
 *   await parser.open(pptxBuffer);
 *   const presentation = await parser.getPresentation();
 *   // ... process slides
 * } finally {
 *   parser.close(); // Always close to clear caches
 * }
 * ```
 */
export declare class PptxParser {
    private zip;
    private readonly logger;
    private readonly xmlParser;
    private readonly orderedXmlParser;
    /** Cache for parsed relationship arrays, keyed by .rels file path. */
    private relationshipCache;
    /** Cache for parsed XML content, keyed by file path within the PPTX. */
    private xmlCache;
    /** Cache for raw XML strings, keyed by file path within the PPTX. */
    private rawXmlCache;
    /** Path to the main presentation XML file, discovered from _rels/.rels */
    private presentationPath;
    constructor(logger?: ILogger);
    /**
     * Opens a PPTX file from a file path or Buffer.
     */
    open(input: Buffer | string): Promise<void>;
    /**
     * Ensures the parser has an open PPTX file.
     */
    private ensureOpen;
    /**
     * Reads and parses an XML file from the PPTX.
     */
    private readXml;
    /**
     * Gets the raw XML content for a file path.
     * Returns undefined if the file hasn't been read yet.
     */
    getRawXml(path: string): string | undefined;
    /**
     * Reads and parses an XML file with preserved element order.
     * This is used for z-order sensitive operations.
     */
    readXmlOrdered(path: string): Promise<OrderedXmlOutput>;
    /**
     * Reads a binary file from the PPTX.
     */
    readBinary(path: string): Promise<Buffer>;
    /**
     * Checks if a file exists in the PPTX.
     */
    fileExists(path: string): boolean;
    /**
     * Lists all files in the PPTX.
     */
    listFiles(): string[];
    /**
     * Parses relationships from a .rels file.
     */
    getRelationships(relPath: string): Promise<Relationship[]>;
    /**
     * Gets the relationship target for a given ID.
     */
    getRelationshipTarget(relPath: string, relId: string): Promise<string | undefined>;
    /**
     * Resolves a relative path to an absolute path within the PPTX.
     *
     * Handles '.' and '..' segments with a segment stack, so it terminates on
     * all inputs (including hostile relationship targets from untrusted files).
     * '..' segments that would escape the archive root are clamped at the root.
     */
    resolvePath(basePath: string, relativePath: string): string;
    /**
     * Finds the path to the main presentation XML file by reading _rels/.rels.
     * This handles non-standard PPTX files where the presentation is not at ppt/presentation.xml.
     */
    findPresentationPath(): Promise<string>;
    /**
     * Gets the relationships file path for the presentation.
     */
    private getPresentationRelsPath;
    /**
     * Gets presentation data.
     */
    getPresentation(): Promise<PresentationData>;
    /**
     * Gets the number of slides in the presentation.
     */
    getSlideCount(): Promise<number>;
    /**
     * Extracts fonts embedded in the presentation (p:embeddedFontLst).
     *
     * Each p:embeddedFont carries a p:font typeface name plus up to four
     * variant references (p:regular, p:bold, p:italic, p:boldItalic) whose
     * r:id values resolve through the presentation relationships to font
     * parts (typically ppt/fonts/*.fntdata).
     *
     * The returned variant data is the raw part content: it may be a plain
     * TTF/OTF, an EOT wrapper, or an ODTTF-obfuscated font. Use
     * `decodeFntdata` from the text module to normalize it.
     *
     * @returns Embedded fonts in declaration order; empty when the
     *   presentation embeds no fonts.
     */
    getEmbeddedFonts(): Promise<EmbeddedFontData[]>;
    /**
     * Gets slide data by index (0-based).
     */
    getSlide(index: number): Promise<SlideData>;
    /**
     * Gets slide layout data.
     */
    getSlideLayout(slidePath: string, layoutRelId: string): Promise<SlideLayoutData>;
    /**
     * Gets slide master data.
     */
    getSlideMaster(layoutPath: string, masterRelId: string): Promise<SlideMasterData>;
    /**
     * Gets the theme data referenced by a specific slide master. Each master
     * carries its own theme relationship (ECMA-376 §14.2.7), so decks with
     * multiple masters can have several themes with different color schemes;
     * colors must resolve against the theme of the SLIDE'S OWN master chain,
     * not a single presentation-wide theme.
     *
     * Returns undefined when the master has no theme relationship or the
     * referenced part is missing/invalid (callers fall back to the
     * presentation-default theme).
     */
    getThemeForMaster(masterPath: string): Promise<ThemeData | undefined>;
    /**
     * Gets theme data.
     */
    getTheme(): Promise<ThemeData>;
    /**
     * Reads a theme part via the ordered parser and converts it to a standard
     * node. The conversion attaches ordered sources, so consumers that need
     * document order (e.g. fmtScheme style lists, whose 1-based fillRef
     * indices depend on true element order even across different fill kinds)
     * get the exact order from the file.
     */
    private readThemeNode;
    /**
     * Gets media file by relationship ID.
     */
    getMedia(slidePath: string, relationshipId: string): Promise<Buffer>;
    /**
     * Closes the PPTX file and clears all internal caches.
     *
     * This method should always be called when you are done with the parser
     * to release the ZIP file reference and clear the XML and relationship caches.
     * Failure to call this method may result in memory not being released.
     *
     * After calling `close()`, the parser cannot be used until `open()` is called again.
     */
    close(): void;
}
/**
 * Utility function to extract attribute value from XML node.
 */
export declare function getXmlAttr(node: PptxXmlNode | undefined, attr: string): string | undefined;
/**
 * Utility function to get a child element from XML node.
 */
export declare function getXmlChild(node: PptxXmlNode | undefined, path: string): PptxXmlNode | undefined;
/**
 * Checks whether a child element is present at all. Self-closed elements
 * (e.g. `<a:noFill/>`) parse to an empty string, which is falsy — use this
 * instead of truthiness on getXmlChild when presence itself is the signal.
 */
export declare function hasXmlChild(node: PptxXmlNode | undefined, path: string): boolean;
/**
 * Utility function to get a child element as array.
 */
export declare function getXmlChildren(node: PptxXmlNode | undefined, path: string): PptxXmlNode[];
/**
 * Represents a single element in the ordered XML output from fast-xml-parser.
 * Each element has one key (the tag name) with children as value, and optionally ':@' for attributes.
 */
export interface OrderedXmlNode {
    [tagName: string]: OrderedXmlOutput | string | Record<string, string> | undefined;
}
/**
 * Type representing the output of fast-xml-parser with preserveOrder: true.
 * Returns an array of elements in document order.
 */
export type OrderedXmlOutput = OrderedXmlNode[];
/**
 * Parses an XML string with preserved document order.
 * Returns an array of elements in the order they appear in the document.
 *
 * @param xmlString Raw XML string to parse
 * @returns Parsed XML with preserved order
 */
export declare function parseXmlPreservingOrder(xmlString: string): OrderedXmlOutput;
/**
 * Extracts ordered child elements from an ordered XML node.
 * Filters to only include specified tag names.
 * Fully converts nested ordered XML structures to standard PptxXmlNode format.
 *
 * @param orderedNode Array of ordered XML elements
 * @param tagNames Tag names to include (e.g., ['a:moveTo', 'a:lnTo'])
 * @returns Array of OrderedXmlElement in document order
 */
export declare function getOrderedChildren(orderedNode: OrderedXmlOutput, tagNames: readonly string[]): OrderedXmlElement[];
/**
 * Extracts ordered child elements from an XML node, handling the raw ordered output.
 * Use this for getting children of any element in document order.
 *
 * @param parentChildren The children array from an ordered element
 * @param tagNames Tag names to filter by
 * @returns Ordered elements matching the specified tags
 */
export declare function extractOrderedElements(parentChildren: OrderedXmlOutput | undefined, tagNames: readonly string[]): OrderedXmlElement[];
/**
 * Converts a parsed XML node back to an XML string.
 * This is useful for re-parsing a node with different options (e.g., preserveOrder).
 *
 * @param node The parsed XML node
 * @param wrapperTag Tag name to wrap the node content with
 * @returns XML string representation of the node
 */
export declare function nodeToXmlString(node: PptxXmlNode, wrapperTag: string): string;
/**
 * Gets children of a parsed XML node in document order by re-parsing with preserveOrder.
 * This is the main function to use when you need to iterate over child elements
 * in their original document order (for z-order or path segment order).
 *
 * @param node The parent node containing children to iterate
 * @param wrapperTag The tag name of the parent node (needed for XML serialization)
 * @param childTagNames Array of child tag names to filter and return in order
 * @returns Array of ordered child elements
 *
 * @example
 * ```typescript
 * // Get path segments in document order
 * const segments = getChildrenInDocumentOrder(
 *   pathNode,
 *   'a:path',
 *   ['a:moveTo', 'a:lnTo', 'a:cubicBezTo', 'a:arcTo', 'a:close']
 * );
 *
 * // Get shape tree elements in document order
 * const shapes = getChildrenInDocumentOrder(
 *   spTree,
 *   'p:spTree',
 *   ['p:sp', 'p:cxnSp', 'p:pic', 'p:grpSp']
 * );
 * ```
 */
export declare function getChildrenInDocumentOrder(node: PptxXmlNode, wrapperTag: string, childTagNames: readonly string[]): OrderedXmlElement[];
