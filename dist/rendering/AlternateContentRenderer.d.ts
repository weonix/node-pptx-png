/**
 * Handles mc:AlternateContent elements (SmartArt, diagrams, etc.).
 * Parses mc:Choice and mc:Fallback elements, rendering fallback content
 * when primary content is not supported.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Configuration for AlternateContentRenderer.
 */
export interface AlternateContentRendererConfig {
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Result of parsing AlternateContent.
 */
export interface AlternateContentResult {
    /** Whether we should render the Choice content */
    useChoice: boolean;
    /** The content to render (either Choice or Fallback) */
    content: PptxXmlNode | undefined;
    /** The namespace requirement (if Choice) */
    requires?: string;
    /** List of child elements to render */
    children: PptxXmlNode[];
}
/**
 * Callback function type for rendering child elements.
 */
export type AlternateContentRenderCallback = (ctx: CanvasRenderingContext2D, tagName: string, node: PptxXmlNode) => Promise<void>;
/**
 * Renders mc:AlternateContent elements by selecting and rendering
 * the appropriate content (Choice or Fallback).
 */
export declare class AlternateContentRenderer {
    private readonly logger;
    constructor(config?: AlternateContentRendererConfig);
    /**
     * Checks if we support rendering the content specified by a namespace.
     * @param requires The namespace requirement string
     * @returns True if we can render this content
     */
    supportsNamespace(requires: string | undefined): boolean;
    /**
     * Parses an mc:AlternateContent element and determines what to render.
     * @param alternateContentNode The mc:AlternateContent XML node
     * @returns The parsed result indicating what content to render
     */
    parseAlternateContent(alternateContentNode: PptxXmlNode): AlternateContentResult;
    /**
     * Extracts renderable child elements from a Choice or Fallback node.
     * @param node The mc:Choice or mc:Fallback node
     * @returns Array of child element nodes
     */
    private extractChildren;
    /**
     * Gets the tag name for a child element.
     * @param node The child node
     * @param fallbackNode The fallback node containing the child
     * @returns The tag name of the child
     */
    getChildTagName(node: PptxXmlNode, fallbackNode: PptxXmlNode): string;
    /**
     * Renders the content of an mc:AlternateContent element.
     * @param ctx Canvas 2D context
     * @param alternateContentNode The mc:AlternateContent XML node
     * @param renderChild Callback to render individual child elements
     */
    renderAlternateContent(ctx: CanvasRenderingContext2D, alternateContentNode: PptxXmlNode, renderChild: AlternateContentRenderCallback): Promise<void>;
    /**
     * Checks if a node is an mc:AlternateContent element.
     * @param tagName The tag name to check
     * @returns True if this is an AlternateContent element
     */
    static isAlternateContent(tagName: string): boolean;
}
/**
 * Creates an AlternateContentRenderer instance.
 */
export declare function createAlternateContentRenderer(logger?: ILogger): AlternateContentRenderer;
//# sourceMappingURL=AlternateContentRenderer.d.ts.map