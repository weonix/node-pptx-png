/**
 * Parses .rels files to resolve relationship IDs to file paths.
 * Handles slide, layout, and master relationships.
 * Resolves image references (r:embed) to media file paths.
 */
import type { PptxParser, Relationship } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Relationship types commonly found in PPTX files.
 */
export declare const RelationshipTypes: {
    readonly image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    readonly slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    readonly slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    readonly slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    readonly theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    readonly hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    readonly chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    readonly oleObject: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    readonly diagramData: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
    /** Pre-rendered SmartArt drawing part (MS-ODRAWXML dsp: namespace) */
    readonly diagramDrawing: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
};
/**
 * Configuration for RelationshipParser.
 */
export interface RelationshipParserConfig {
    /** PPTX parser instance for accessing files */
    parser: PptxParser;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Parser for PPTX relationship files (.rels).
 * Provides methods to resolve relationship IDs to actual file paths within the PPTX.
 */
export declare class RelationshipParser {
    private readonly logger;
    private readonly parser;
    /** Cache of parsed relationships, keyed by source file path */
    private readonly cache;
    constructor(config: RelationshipParserConfig);
    /**
     * Gets the .rels file path for a given source file.
     * For example:
     * - ppt/slides/slide1.xml -> ppt/slides/_rels/slide1.xml.rels
     * - ppt/slideLayouts/slideLayout1.xml -> ppt/slideLayouts/_rels/slideLayout1.xml.rels
     */
    getRelsPath(sourcePath: string): string;
    /**
     * Gets the base directory for resolving relative paths from a source file.
     */
    private getBaseDir;
    /**
     * Loads and caches relationships for a source file.
     */
    private loadRelationships;
    /**
     * Resolves a relationship ID to a full file path within the PPTX.
     *
     * @param sourcePath The source file path (e.g., ppt/slides/slide1.xml)
     * @param relationshipId The relationship ID (e.g., rId1)
     * @returns The resolved file path or undefined if not found
     */
    resolveRelationshipId(sourcePath: string, relationshipId: string): Promise<string | undefined>;
    /**
     * Resolves an image relationship ID to the media file path.
     *
     * @param sourcePath The source file path (e.g., ppt/slides/slide1.xml)
     * @param relationshipId The relationship ID (e.g., rId2)
     * @returns The resolved media file path or undefined if not found
     */
    resolveImageRelationship(sourcePath: string, relationshipId: string): Promise<string | undefined>;
    /**
     * Gets all image relationships for a source file.
     *
     * @param sourcePath The source file path
     * @returns Array of {id, path} for all image relationships
     */
    getImageRelationships(sourcePath: string): Promise<Array<{
        id: string;
        path: string;
    }>>;
    /**
     * Gets a relationship by ID.
     *
     * @param sourcePath The source file path
     * @param relationshipId The relationship ID
     * @returns The relationship object or undefined
     */
    getRelationship(sourcePath: string, relationshipId: string): Promise<Relationship | undefined>;
    /**
     * Gets all relationships for a source file.
     *
     * @param sourcePath The source file path
     * @returns Array of all relationships
     */
    getAllRelationships(sourcePath: string): Promise<Relationship[]>;
    /**
     * Gets relationships of a specific type.
     *
     * @param sourcePath The source file path
     * @param type The relationship type to filter by
     * @returns Array of relationships matching the type
     */
    getRelationshipsByType(sourcePath: string, type: string): Promise<Relationship[]>;
    /**
     * Clears the relationship cache.
     * Call this when switching to a new PPTX file.
     */
    clearCache(): void;
}
/**
 * Creates a RelationshipParser instance.
 */
export declare function createRelationshipParser(parser: PptxParser, logger?: ILogger): RelationshipParser;
