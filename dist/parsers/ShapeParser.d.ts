/**
 * Parses shape elements (p:sp, p:cxnSp) from slide XML.
 * Extracts structured shape data for rendering.
 *
 * Note: Fill and stroke parsing is handled by FillRenderer and StrokeRenderer
 * respectively. This parser extracts raw XML nodes for those components.
 */
import type { ShapeElement } from '../types/elements.js';
import type { ResolvedTheme } from '../types/theme.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Configuration for ShapeParser.
 */
export interface ShapeParserConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Parses shape elements from slide XML.
 * Note: Fill and stroke are returned as raw XML nodes (spPrNode).
 * Use FillRenderer.parseFill() and StrokeRenderer.parseStroke() for actual parsing.
 */
export declare class ShapeParser {
    private readonly logger;
    constructor(config: ShapeParserConfig);
    /**
     * Parses a shape element (p:sp) from XML.
     * @param spNode Shape XML node
     * @returns Parsed ShapeElement or undefined if invalid
     */
    parseShape(spNode: PptxXmlNode): ShapeElement | undefined;
    /**
     * Parses a placeholder reference from non-visual properties.
     */
    private parsePlaceholder;
    /**
     * Parses transform (xfrm) from shape properties.
     */
    private parseTransform;
    /**
     * Parses geometry from shape properties.
     *
     * Custom geometry (a:custGeom) is parsed in document order and evaluated
     * through the shared geometry engine (guide formulas, per-path coordinate
     * spaces, arcs). The resulting paths are in the shape's EMU coordinate
     * space (origin at the shape's top-left).
     */
    private parseGeometry;
    /**
     * Parses text body from shape.
     */
    private parseTextBody;
    /**
     * Parses a paragraph.
     */
    private parseParagraph;
    /**
     * Parses a text run.
     */
    private parseTextRun;
}
/**
 * Creates a ShapeParser instance.
 */
export declare function createShapeParser(theme: ResolvedTheme, logger?: ILogger): ShapeParser;
//# sourceMappingURL=ShapeParser.d.ts.map