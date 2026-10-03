"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.PptxParser = exports.trimAttributeValue = exports.trimTagValueExceptText = void 0;
exports.getXmlAttr = getXmlAttr;
exports.getXmlChild = getXmlChild;
exports.hasXmlChild = hasXmlChild;
exports.getXmlChildren = getXmlChildren;
exports.parseXmlPreservingOrder = parseXmlPreservingOrder;
exports.getOrderedChildren = getOrderedChildren;
exports.extractOrderedElements = extractOrderedElements;
exports.nodeToXmlString = nodeToXmlString;
exports.getChildrenInDocumentOrder = getChildrenInDocumentOrder;
const jszip_1 = __importDefault(require("jszip"));
const fast_xml_parser_1 = require("fast-xml-parser");
const fs = __importStar(require("fs/promises"));
const Logger_js_1 = require("../utils/Logger.js");
/**
 * XML attribute prefix used by fast-xml-parser.
 */
const ATTR_PREFIX = '@_';
/**
 * Common namespace prefixes in PPTX XML.
 */
const NAMESPACES = {
    presentation: 'p:presentation',
    slide: 'p:sld',
    slideLayout: 'p:sldLayout',
    slideMaster: 'p:sldMaster',
    theme: 'a:theme',
    relationships: 'Relationships',
    relationship: 'Relationship',
};
/**
 * XML element names that should always be parsed as arrays.
 * These elements can appear multiple times in PPTX XML.
 */
const ARRAY_ELEMENTS = [
    'p:sp',
    'p:pic',
    'p:grpSp',
    'p:cxnSp',
    'p:graphicFrame',
    'a:p',
    'a:r',
    'a:gs',
    'p:sldId',
    'Relationship',
    'a:path',
    'a:moveTo',
    'a:lnTo',
    'a:cubicBezTo',
    'a:arcTo',
    'a:close',
];
/**
 * Tags whose text content is significant and must never be trimmed.
 * Per ECMA-376, whitespace inside <a:t> is meaningful: PowerPoint splits a
 * sentence into multiple runs whenever formatting changes mid-sentence, so
 * run-boundary spaces (e.g., '<a:t>world </a:t>') carry real spacing, and
 * whitespace-only runs ('<a:t xml:space="preserve"> </a:t>') are valid text.
 */
const PRESERVE_WHITESPACE_TAGS = new Set(['a:t']);
/**
 * Tag value processor that trims text values for all tags except those where
 * whitespace is significant. Used with `trimValues: false` so that a:t runs
 * keep their whitespace while inter-element whitespace (pretty-printed XML)
 * is still discarded for structural tags.
 */
const trimTagValueExceptText = (tagName, tagText) => PRESERVE_WHITESPACE_TAGS.has(tagName) ? tagText : tagText.trim();
exports.trimTagValueExceptText = trimTagValueExceptText;
/**
 * Attribute value processor that trims attribute values, preserving the
 * behavior `trimValues: true` previously provided for attributes.
 */
const trimAttributeValue = (_attrName, attrValue) => attrValue.trim();
exports.trimAttributeValue = trimAttributeValue;
/**
 * Default XML parser options.
 *
 * Note: `trimValues` is false with tag/attribute value processors that trim
 * everything except significant text tags (see PRESERVE_WHITESPACE_TAGS).
 */
const XML_PARSER_OPTIONS = {
    ignoreAttributes: false,
    attributeNamePrefix: ATTR_PREFIX,
    removeNSPrefix: false,
    parseAttributeValue: false,
    trimValues: false,
    tagValueProcessor: exports.trimTagValueExceptText,
    attributeValueProcessor: exports.trimAttributeValue,
    parseTagValue: false,
    isArray: (name) => {
        return ARRAY_ELEMENTS.some((el) => name.endsWith(el) || name === el);
    },
};
/**
 * XML parser options with preserveOrder enabled.
 * This returns elements in document order as an array structure.
 * Format: [{ tagName: [...children], ':@': { attrs } }, ...]
 */
const ORDERED_XML_PARSER_OPTIONS = {
    ignoreAttributes: false,
    attributeNamePrefix: ATTR_PREFIX,
    removeNSPrefix: false,
    parseAttributeValue: false,
    trimValues: false,
    tagValueProcessor: exports.trimTagValueExceptText,
    attributeValueProcessor: exports.trimAttributeValue,
    parseTagValue: false,
    preserveOrder: true,
};
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
class PptxParser {
    zip = null;
    logger;
    xmlParser;
    orderedXmlParser;
    /** Cache for parsed relationship arrays, keyed by .rels file path. */
    relationshipCache = new Map();
    /** Cache for parsed XML content, keyed by file path within the PPTX. */
    xmlCache = new Map();
    /** Cache for raw XML strings, keyed by file path within the PPTX. */
    rawXmlCache = new Map();
    /** Path to the main presentation XML file, discovered from _rels/.rels */
    presentationPath = null;
    constructor(logger) {
        this.logger = logger ?? (0, Logger_js_1.createLogger)('warn', 'PptxParser');
        this.xmlParser = new fast_xml_parser_1.XMLParser(XML_PARSER_OPTIONS);
        this.orderedXmlParser = new fast_xml_parser_1.XMLParser(ORDERED_XML_PARSER_OPTIONS);
    }
    /**
     * Opens a PPTX file from a file path or Buffer.
     */
    async open(input) {
        let data;
        if (typeof input === 'string') {
            this.logger.debug('Opening PPTX from file path', { path: input });
            data = await fs.readFile(input);
        }
        else {
            this.logger.debug('Opening PPTX from buffer', { size: input.length });
            data = input;
        }
        try {
            this.zip = await jszip_1.default.loadAsync(data);
            this.relationshipCache.clear();
            this.xmlCache.clear();
            this.rawXmlCache.clear();
            this.presentationPath = null;
            this.logger.info('PPTX opened successfully');
        }
        catch (error) {
            const message = error instanceof Error ? error.message : 'Unknown error';
            this.logger.error('Failed to open PPTX', { error: message });
            throw new Error(`Failed to open PPTX file: ${message}`, { cause: error });
        }
    }
    /**
     * Ensures the parser has an open PPTX file.
     */
    ensureOpen() {
        if (!this.zip) {
            throw new Error('No PPTX file is open. Call open() first.');
        }
        return this.zip;
    }
    /**
     * Reads and parses an XML file from the PPTX.
     */
    async readXml(path) {
        // Check cache first
        const cached = this.xmlCache.get(path);
        if (cached) {
            return cached;
        }
        const zip = this.ensureOpen();
        const file = zip.file(path);
        if (!file) {
            throw new Error(`File not found in PPTX: ${path}`);
        }
        const content = await file.async('string');
        // Cache the raw XML for potential ordered parsing later
        this.rawXmlCache.set(path, content);
        const parsed = this.xmlParser.parse(content);
        // Cache the parsed result
        this.xmlCache.set(path, parsed);
        return parsed;
    }
    /**
     * Gets the raw XML content for a file path.
     * Returns undefined if the file hasn't been read yet.
     */
    getRawXml(path) {
        return this.rawXmlCache.get(path);
    }
    /**
     * Reads and parses an XML file with preserved element order.
     * This is used for z-order sensitive operations.
     */
    async readXmlOrdered(path) {
        // Check raw cache first
        let content = this.rawXmlCache.get(path);
        if (!content) {
            const zip = this.ensureOpen();
            const file = zip.file(path);
            if (!file) {
                throw new Error(`File not found in PPTX: ${path}`);
            }
            content = await file.async('string');
            this.rawXmlCache.set(path, content);
        }
        return this.orderedXmlParser.parse(content);
    }
    /**
     * Reads a binary file from the PPTX.
     */
    async readBinary(path) {
        const zip = this.ensureOpen();
        const file = zip.file(path);
        if (!file) {
            throw new Error(`File not found in PPTX: ${path}`);
        }
        const data = await file.async('nodebuffer');
        return data;
    }
    /**
     * Checks if a file exists in the PPTX.
     */
    fileExists(path) {
        const zip = this.ensureOpen();
        return zip.file(path) !== null;
    }
    /**
     * Lists all files in the PPTX.
     */
    listFiles() {
        const zip = this.ensureOpen();
        const files = [];
        zip.forEach((relativePath) => {
            files.push(relativePath);
        });
        return files;
    }
    /**
     * Parses relationships from a .rels file.
     */
    async getRelationships(relPath) {
        // Check cache
        const cached = this.relationshipCache.get(relPath);
        if (cached) {
            return cached;
        }
        if (!this.fileExists(relPath)) {
            this.logger.debug('Relationships file not found', { path: relPath });
            return [];
        }
        const xml = await this.readXml(relPath);
        const relationships = [];
        const rels = xml[NAMESPACES.relationships];
        if (!rels) {
            return [];
        }
        const relElements = rels[NAMESPACES.relationship];
        if (!relElements) {
            return [];
        }
        const relArray = Array.isArray(relElements) ? relElements : [relElements];
        for (const rel of relArray) {
            const relNode = rel;
            relationships.push({
                id: getXmlAttr(relNode, 'Id') ?? '',
                type: getXmlAttr(relNode, 'Type') ?? '',
                target: getXmlAttr(relNode, 'Target') ?? '',
            });
        }
        // Cache the result
        this.relationshipCache.set(relPath, relationships);
        return relationships;
    }
    /**
     * Gets the relationship target for a given ID.
     */
    async getRelationshipTarget(relPath, relId) {
        const rels = await this.getRelationships(relPath);
        const rel = rels.find((r) => r.id === relId);
        return rel?.target;
    }
    /**
     * Resolves a relative path to an absolute path within the PPTX.
     *
     * Handles '.' and '..' segments with a segment stack, so it terminates on
     * all inputs (including hostile relationship targets from untrusted files).
     * '..' segments that would escape the archive root are clamped at the root.
     */
    resolvePath(basePath, relativePath) {
        if (relativePath.startsWith('/')) {
            return relativePath.slice(1);
        }
        const baseDir = basePath.substring(0, basePath.lastIndexOf('/') + 1);
        const segments = [];
        for (const segment of (baseDir + relativePath).split('/')) {
            if (segment === '' || segment === '.') {
                continue;
            }
            if (segment === '..') {
                // Pop the previous segment; a no-op at the root (cannot escape the archive)
                segments.pop();
            }
            else {
                segments.push(segment);
            }
        }
        return segments.join('/');
    }
    /**
     * Finds the path to the main presentation XML file by reading _rels/.rels.
     * This handles non-standard PPTX files where the presentation is not at ppt/presentation.xml.
     */
    async findPresentationPath() {
        // Return cached path if already discovered
        if (this.presentationPath) {
            return this.presentationPath;
        }
        const relsPath = '_rels/.rels';
        if (!this.fileExists(relsPath)) {
            this.logger.warn('Root .rels file not found, using default presentation path');
            this.presentationPath = 'ppt/presentation.xml';
            return this.presentationPath;
        }
        const relsContent = await this.readXml(relsPath);
        const relationships = getXmlChildren(relsContent, NAMESPACES.relationships);
        // If Relationships is the root, get its children
        let relElements = [];
        if (relationships.length > 0) {
            relElements = getXmlChildren(relationships[0], NAMESPACES.relationship);
        }
        else {
            // Try direct access if Relationships is the wrapper
            const rels = relsContent[NAMESPACES.relationships];
            if (rels) {
                const children = rels[NAMESPACES.relationship];
                relElements = children
                    ? Array.isArray(children)
                        ? children
                        : [children]
                    : [];
            }
        }
        for (const rel of relElements) {
            const type = getXmlAttr(rel, 'Type') ?? '';
            // Look for the main officeDocument relationship (ends with /officeDocument)
            // This avoids matching extended-properties or other relationships that contain 'officeDocument' in the URL
            if (type.endsWith('/officeDocument')) {
                let target = getXmlAttr(rel, 'Target');
                if (target) {
                    // Remove leading slash if present (some PPTX files use /ppt/presentation.xml)
                    if (target.startsWith('/')) {
                        target = target.substring(1);
                    }
                    this.presentationPath = target;
                    this.logger.info('Found presentation path from .rels', { path: this.presentationPath });
                    return this.presentationPath;
                }
            }
        }
        // Fallback to default path
        this.logger.warn('No officeDocument relationship found, using default presentation path');
        this.presentationPath = 'ppt/presentation.xml';
        return this.presentationPath;
    }
    /**
     * Gets the relationships file path for the presentation.
     */
    getPresentationRelsPath(presPath) {
        // Convert ppt/presentation.xml to ppt/_rels/presentation.xml.rels
        const lastSlash = presPath.lastIndexOf('/');
        if (lastSlash === -1) {
            return `_rels/${presPath}.rels`;
        }
        const dir = presPath.substring(0, lastSlash);
        const filename = presPath.substring(lastSlash + 1);
        return `${dir}/_rels/${filename}.rels`;
    }
    /**
     * Gets presentation data.
     */
    async getPresentation() {
        const presentationPath = await this.findPresentationPath();
        const xml = await this.readXml(presentationPath);
        const presentation = xml[NAMESPACES.presentation];
        if (!presentation) {
            throw new Error('Invalid PPTX: missing presentation element');
        }
        // Get slide size
        const sldSz = presentation['p:sldSz'];
        const slideWidth = sldSz ? parseInt(getXmlAttr(sldSz, 'cx') ?? '9144000', 10) : 9144000;
        const slideHeight = sldSz ? parseInt(getXmlAttr(sldSz, 'cy') ?? '6858000', 10) : 6858000;
        // Get slide IDs
        const sldIdLst = presentation['p:sldIdLst'];
        const slideIds = [];
        if (sldIdLst) {
            const sldIdElements = sldIdLst['p:sldId'];
            if (sldIdElements) {
                const sldIdArray = Array.isArray(sldIdElements) ? sldIdElements : [sldIdElements];
                for (const sldId of sldIdArray) {
                    const idNode = sldId;
                    const rId = getXmlAttr(idNode, 'r:id');
                    if (rId) {
                        slideIds.push(rId);
                    }
                }
            }
        }
        this.logger.info('Presentation data loaded', {
            slideWidth,
            slideHeight,
            slideCount: slideIds.length,
        });
        return {
            slideWidth,
            slideHeight,
            slideIds,
            slideCount: slideIds.length,
            content: presentation,
        };
    }
    /**
     * Gets the number of slides in the presentation.
     */
    async getSlideCount() {
        const presentation = await this.getPresentation();
        return presentation.slideCount;
    }
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
    async getEmbeddedFonts() {
        const presentation = await this.getPresentation();
        const presentationPath = await this.findPresentationPath();
        const presRelsPath = this.getPresentationRelsPath(presentationPath);
        const embeddedFontLst = getXmlChild(presentation.content, 'p:embeddedFontLst');
        if (!embeddedFontLst) {
            return [];
        }
        const variantTags = [
            { tag: 'p:regular', style: 'regular' },
            { tag: 'p:bold', style: 'bold' },
            { tag: 'p:italic', style: 'italic' },
            { tag: 'p:boldItalic', style: 'boldItalic' },
        ];
        const fonts = [];
        for (const embeddedFont of getXmlChildren(embeddedFontLst, 'p:embeddedFont')) {
            const fontEl = getXmlChild(embeddedFont, 'p:font');
            const typeface = getXmlAttr(fontEl, 'typeface');
            if (!typeface) {
                this.logger.warn('Skipping p:embeddedFont without a p:font typeface');
                continue;
            }
            const variants = [];
            for (const { tag, style } of variantTags) {
                const variantEl = getXmlChild(embeddedFont, tag);
                const relId = getXmlAttr(variantEl, 'r:id');
                if (!relId) {
                    continue;
                }
                try {
                    const target = await this.getRelationshipTarget(presRelsPath, relId);
                    if (!target) {
                        this.logger.warn('Embedded font relationship not found', { typeface, style, relId });
                        continue;
                    }
                    const fontPath = this.resolvePath(presentationPath, target);
                    const data = await this.readBinary(fontPath);
                    variants.push({ style, path: fontPath, data });
                }
                catch (error) {
                    this.logger.warn('Failed to extract embedded font variant', {
                        typeface,
                        style,
                        relId,
                        error: error instanceof Error ? error.message : String(error),
                    });
                }
            }
            if (variants.length > 0) {
                fonts.push({ typeface, variants });
            }
            else {
                this.logger.warn('Embedded font has no readable variants', { typeface });
            }
        }
        this.logger.info('Embedded fonts extracted', {
            fontCount: fonts.length,
            variantCount: fonts.reduce((n, f) => n + f.variants.length, 0),
        });
        return fonts;
    }
    /**
     * Gets slide data by index (0-based).
     */
    async getSlide(index) {
        const presentation = await this.getPresentation();
        const presentationPath = await this.findPresentationPath();
        if (index < 0 || index >= presentation.slideCount) {
            throw new Error(`Slide index ${index} out of range (0-${presentation.slideCount - 1})`);
        }
        const slideRelId = presentation.slideIds[index];
        if (!slideRelId) {
            throw new Error(`No relationship ID found for slide ${index}`);
        }
        // Get the slide path from relationships using the dynamic presentation path
        const presRelsPath = this.getPresentationRelsPath(presentationPath);
        const rels = await this.getRelationships(presRelsPath);
        const slideRel = rels.find((r) => r.id === slideRelId);
        if (!slideRel) {
            throw new Error(`Relationship not found for slide ${index}: ${slideRelId}`);
        }
        const slidePath = this.resolvePath(presentationPath, slideRel.target);
        const xml = await this.readXml(slidePath);
        const slide = xml[NAMESPACES.slide];
        if (!slide) {
            throw new Error(`Invalid slide XML: missing slide element in ${slidePath}`);
        }
        // Get layout relationship
        const slideRelsPath = slidePath
            .replace('slides/', 'slides/_rels/')
            .replace('.xml', '.xml.rels');
        const slideRels = await this.getRelationships(slideRelsPath);
        const layoutRel = slideRels.find((r) => r.type.includes('slideLayout'));
        this.logger.debug('Slide loaded', { index, path: slidePath });
        return {
            index,
            content: slide,
            layoutRelId: layoutRel?.id,
            path: slidePath,
        };
    }
    /**
     * Gets slide layout data.
     */
    async getSlideLayout(slidePath, layoutRelId) {
        const slideRelsPath = slidePath
            .replace('slides/', 'slides/_rels/')
            .replace('.xml', '.xml.rels');
        const target = await this.getRelationshipTarget(slideRelsPath, layoutRelId);
        if (!target) {
            throw new Error(`Layout relationship not found: ${layoutRelId}`);
        }
        const layoutPath = this.resolvePath(slidePath, target);
        const xml = await this.readXml(layoutPath);
        const layout = xml[NAMESPACES.slideLayout];
        if (!layout) {
            throw new Error(`Invalid layout XML: missing slideLayout element in ${layoutPath}`);
        }
        // Get master relationship
        const layoutRelsPath = layoutPath
            .replace('slideLayouts/', 'slideLayouts/_rels/')
            .replace('.xml', '.xml.rels');
        const layoutRels = await this.getRelationships(layoutRelsPath);
        const masterRel = layoutRels.find((r) => r.type.includes('slideMaster'));
        return {
            name: layout[`${ATTR_PREFIX}name`],
            content: layout,
            masterRelId: masterRel?.id,
            path: layoutPath,
        };
    }
    /**
     * Gets slide master data.
     */
    async getSlideMaster(layoutPath, masterRelId) {
        const layoutRelsPath = layoutPath
            .replace('slideLayouts/', 'slideLayouts/_rels/')
            .replace('.xml', '.xml.rels');
        const target = await this.getRelationshipTarget(layoutRelsPath, masterRelId);
        if (!target) {
            throw new Error(`Master relationship not found: ${masterRelId}`);
        }
        const masterPath = this.resolvePath(layoutPath, target);
        const xml = await this.readXml(masterPath);
        const master = xml[NAMESPACES.slideMaster];
        if (!master) {
            throw new Error(`Invalid master XML: missing slideMaster element in ${masterPath}`);
        }
        // Get theme relationship
        const masterRelsPath = masterPath
            .replace('slideMasters/', 'slideMasters/_rels/')
            .replace('.xml', '.xml.rels');
        const masterRels = await this.getRelationships(masterRelsPath);
        const themeRel = masterRels.find((r) => r.type.includes('theme'));
        return {
            name: master[`${ATTR_PREFIX}name`],
            content: master,
            themeRelId: themeRel?.id,
            path: masterPath,
        };
    }
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
    async getThemeForMaster(masterPath) {
        const masterRelsPath = masterPath
            .replace('slideMasters/', 'slideMasters/_rels/')
            .replace('.xml', '.xml.rels');
        const masterRels = await this.getRelationships(masterRelsPath);
        const themeRel = masterRels.find((r) => r.type.includes('theme'));
        if (!themeRel) {
            return undefined;
        }
        const themePath = this.resolvePath(masterPath, themeRel.target);
        const theme = await this.readThemeNode(themePath);
        if (!theme) {
            return undefined;
        }
        this.logger.debug('Theme loaded for master', { master: masterPath, path: themePath });
        return { content: theme, path: themePath };
    }
    /**
     * Gets theme data.
     */
    async getTheme() {
        // First, try to get theme from the first slide master
        const presentationPath = await this.findPresentationPath();
        const presRelsPath = this.getPresentationRelsPath(presentationPath);
        const rels = await this.getRelationships(presRelsPath);
        const masterRel = rels.find((r) => r.type.includes('slideMaster'));
        if (masterRel) {
            const masterPath = this.resolvePath(presentationPath, masterRel.target);
            const masterRelsPath = masterPath
                .replace('slideMasters/', 'slideMasters/_rels/')
                .replace('.xml', '.xml.rels');
            const masterRels = await this.getRelationships(masterRelsPath);
            const themeRel = masterRels.find((r) => r.type.includes('theme'));
            if (themeRel) {
                const themePath = this.resolvePath(masterPath, themeRel.target);
                const theme = await this.readThemeNode(themePath);
                if (theme) {
                    this.logger.debug('Theme loaded', { path: themePath });
                    return { content: theme, path: themePath };
                }
            }
        }
        // Fallback: try default theme path
        const defaultThemePath = 'ppt/theme/theme1.xml';
        if (this.fileExists(defaultThemePath)) {
            const theme = await this.readThemeNode(defaultThemePath);
            if (theme) {
                this.logger.debug('Theme loaded from default path', { path: defaultThemePath });
                return { content: theme, path: defaultThemePath };
            }
        }
        throw new Error('No theme found in presentation');
    }
    /**
     * Reads a theme part via the ordered parser and converts it to a standard
     * node. The conversion attaches ordered sources, so consumers that need
     * document order (e.g. fmtScheme style lists, whose 1-based fillRef
     * indices depend on true element order even across different fill kinds)
     * get the exact order from the file.
     */
    async readThemeNode(path) {
        const ordered = await this.readXmlOrdered(path);
        const converted = convertOrderedToStandardNode(ordered);
        return converted[NAMESPACES.theme];
    }
    /**
     * Gets media file by relationship ID.
     */
    async getMedia(slidePath, relationshipId) {
        const slideRelsPath = slidePath
            .replace('slides/', 'slides/_rels/')
            .replace('.xml', '.xml.rels');
        const target = await this.getRelationshipTarget(slideRelsPath, relationshipId);
        if (!target) {
            throw new Error(`Media relationship not found: ${relationshipId}`);
        }
        const mediaPath = this.resolvePath(slidePath, target);
        return this.readBinary(mediaPath);
    }
    /**
     * Closes the PPTX file and clears all internal caches.
     *
     * This method should always be called when you are done with the parser
     * to release the ZIP file reference and clear the XML and relationship caches.
     * Failure to call this method may result in memory not being released.
     *
     * After calling `close()`, the parser cannot be used until `open()` is called again.
     */
    close() {
        this.zip = null;
        this.relationshipCache.clear();
        this.xmlCache.clear();
        this.rawXmlCache.clear();
        this.presentationPath = null;
        this.logger.debug('PPTX closed');
    }
}
exports.PptxParser = PptxParser;
/**
 * Utility function to extract attribute value from XML node.
 */
function getXmlAttr(node, attr) {
    if (!node)
        return undefined;
    const value = node[`${ATTR_PREFIX}${attr}`];
    if (typeof value === 'string')
        return value;
    if (typeof value === 'number' || typeof value === 'boolean')
        return String(value);
    return undefined;
}
/**
 * Utility function to get a child element from XML node.
 */
function getXmlChild(node, path) {
    if (!node)
        return undefined;
    return node[path];
}
/**
 * Checks whether a child element is present at all. Self-closed elements
 * (e.g. `<a:noFill/>`) parse to an empty string, which is falsy — use this
 * instead of truthiness on getXmlChild when presence itself is the signal.
 */
function hasXmlChild(node, path) {
    return node?.[path] !== undefined;
}
/**
 * Utility function to get a child element as array.
 */
function getXmlChildren(node, path) {
    if (!node)
        return [];
    const child = node[path];
    if (!child)
        return [];
    return Array.isArray(child) ? child : [child];
}
/**
 * Ordered XML parser singleton for use by external modules.
 * Used for parsing XML with preserveOrder to maintain document order.
 */
const orderedXmlParserSingleton = new fast_xml_parser_1.XMLParser(ORDERED_XML_PARSER_OPTIONS);
/**
 * XML builder options for converting nodes back to XML.
 */
const XML_BUILDER_OPTIONS = {
    ignoreAttributes: false,
    attributeNamePrefix: ATTR_PREFIX,
    suppressEmptyNode: false,
    format: false,
};
/**
 * XML builder singleton for serializing nodes back to XML strings.
 */
const xmlBuilderSingleton = new fast_xml_parser_1.XMLBuilder(XML_BUILDER_OPTIONS);
/**
 * Parses an XML string with preserved document order.
 * Returns an array of elements in the order they appear in the document.
 *
 * @param xmlString Raw XML string to parse
 * @returns Parsed XML with preserved order
 */
function parseXmlPreservingOrder(xmlString) {
    return orderedXmlParserSingleton.parse(xmlString);
}
/**
 * Recursively converts an ordered XML element to a standard PptxXmlNode structure.
 * This handles arbitrary nesting depth, which is required for shape properties.
 *
 * @param orderedChildren Array of ordered child elements
 * @returns A PptxXmlNode containing all children in standard format
 */
/**
 * Symbol under which a converted node keeps a reference to its original
 * ordered (preserveOrder) children. Non-enumerable, so serialization and
 * key iteration are unaffected; getChildrenInDocumentOrder uses it to
 * recover true document order without a lossy serialize/re-parse round trip.
 */
const ORDERED_SOURCE = Symbol('orderedSource');
function attachOrderedSource(node, ordered) {
    Object.defineProperty(node, ORDERED_SOURCE, { value: ordered, enumerable: false });
}
function getOrderedSource(node) {
    return node[ORDERED_SOURCE];
}
function convertOrderedToStandardNode(orderedChildren) {
    const result = {};
    attachOrderedSource(result, orderedChildren);
    for (const child of orderedChildren) {
        if (typeof child !== 'object' || child === null) {
            continue;
        }
        // Get the attributes (stored in ':@')
        const childAttrs = (child[':@'] ?? {});
        // Get the tag name(s) - each child object has one tag key plus optional ':@'
        const childKeys = Object.keys(child).filter((k) => k !== ':@');
        for (const childTag of childKeys) {
            const childContent = child[childTag];
            // Build the child node starting with its attributes
            const childNode = { ...childAttrs };
            // Handle text content (#text)
            if (childTag === '#text') {
                // Text content is stored directly
                result['#text'] = childContent;
                continue;
            }
            // Recursively process nested children
            if (Array.isArray(childContent) && childContent.length > 0) {
                const nestedResult = convertOrderedToStandardNode(childContent);
                // Merge nested children into the child node. Object.assign does not
                // copy the non-enumerable ordered-source symbol, so re-attach it.
                Object.assign(childNode, nestedResult);
                attachOrderedSource(childNode, childContent);
            }
            // Add the child to the result (handle multiple elements with same tag)
            if (!result[childTag]) {
                result[childTag] = childNode;
            }
            else if (Array.isArray(result[childTag])) {
                result[childTag].push(childNode);
            }
            else {
                result[childTag] = [result[childTag], childNode];
            }
        }
    }
    return result;
}
/**
 * Extracts ordered child elements from an ordered XML node.
 * Filters to only include specified tag names.
 * Fully converts nested ordered XML structures to standard PptxXmlNode format.
 *
 * @param orderedNode Array of ordered XML elements
 * @param tagNames Tag names to include (e.g., ['a:moveTo', 'a:lnTo'])
 * @returns Array of OrderedXmlElement in document order
 */
function getOrderedChildren(orderedNode, tagNames) {
    const result = [];
    const tagSet = new Set(tagNames);
    for (const element of orderedNode) {
        // Each element has one key (tag name) plus optional ':@' for attributes
        const keys = Object.keys(element).filter((k) => k !== ':@');
        for (const tagName of keys) {
            if (tagSet.has(tagName)) {
                const attributes = (element[':@'] ?? {});
                const children = element[tagName];
                // Build a node that matches the PptxXmlNode structure
                // Start with the element's own attributes
                const node = { ...attributes };
                // Recursively convert all children to standard format.
                // Object.assign does not copy the non-enumerable ordered-source
                // symbol, so re-attach it after merging.
                if (Array.isArray(children) && children.length > 0) {
                    const convertedChildren = convertOrderedToStandardNode(children);
                    Object.assign(node, convertedChildren);
                    attachOrderedSource(node, children);
                }
                result.push({ tagName, attributes, node });
            }
        }
    }
    return result;
}
/**
 * Extracts ordered child elements from an XML node, handling the raw ordered output.
 * Use this for getting children of any element in document order.
 *
 * @param parentChildren The children array from an ordered element
 * @param tagNames Tag names to filter by
 * @returns Ordered elements matching the specified tags
 */
function extractOrderedElements(parentChildren, tagNames) {
    if (!parentChildren || !Array.isArray(parentChildren)) {
        return [];
    }
    return getOrderedChildren(parentChildren, tagNames);
}
/**
 * Converts a parsed XML node back to an XML string.
 * This is useful for re-parsing a node with different options (e.g., preserveOrder).
 *
 * @param node The parsed XML node
 * @param wrapperTag Tag name to wrap the node content with
 * @returns XML string representation of the node
 */
function nodeToXmlString(node, wrapperTag) {
    const wrapped = { [wrapperTag]: node };
    return xmlBuilderSingleton.build(wrapped);
}
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
function getChildrenInDocumentOrder(node, wrapperTag, childTagNames) {
    // Fast path: nodes converted from an ordered parse carry their original
    // ordered children, which preserves true interleaving across tag types
    // (e.g. a:r / a:br / a:fld within a paragraph).
    const orderedSource = getOrderedSource(node);
    if (orderedSource) {
        return getOrderedChildren(orderedSource, childTagNames);
    }
    // Fallback: serialize and re-parse with preserveOrder. Same-named children
    // keep their relative order, but interleaving across different tag names
    // collapses to first-appearance key order (a standard-parsed node has
    // already grouped same-named children, so true order is unrecoverable).
    const xmlString = nodeToXmlString(node, wrapperTag);
    // Re-parse with preserveOrder to get elements in document order
    const orderedParsed = orderedXmlParserSingleton.parse(xmlString);
    // The result is an array with one element (the wrapper)
    // We need to get its children
    if (!orderedParsed || orderedParsed.length === 0) {
        return [];
    }
    const wrapper = orderedParsed[0];
    if (!wrapper) {
        return [];
    }
    // Get the children array from the wrapper
    const children = wrapper[wrapperTag];
    if (!children || !Array.isArray(children)) {
        return [];
    }
    // Extract the ordered elements matching our filter
    return getOrderedChildren(children, childTagNames);
}
