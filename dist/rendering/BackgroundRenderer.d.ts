import type { Canvas, CanvasRenderingContext2D } from 'skia-canvas';
import type { Rgba, ResolvedTheme, GradientStop } from '../types/index.js';
import type { PptxParser, PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Background type.
 */
export type BackgroundType = 'solid' | 'gradient' | 'pattern' | 'picture' | 'none';
/**
 * Picture fill data for background images.
 */
export interface PictureFillData {
    /** Relationship ID for the embedded image */
    blipRelId: string;
    /** Source node (slide, layout, or master) where the background is defined */
    source: 'slide' | 'layout' | 'master';
}
/**
 * Parsed background fill data.
 */
export interface ParsedBackground {
    type: BackgroundType;
    color?: Rgba;
    gradientStops?: GradientStop[];
    gradientAngle?: number;
    isRadial?: boolean;
    /** Picture fill data for blipFill backgrounds */
    pictureFill?: PictureFillData;
}
/**
 * Renders slide backgrounds.
 */
export declare class BackgroundRenderer {
    private readonly logger;
    private readonly theme;
    private readonly colorResolver;
    /** Cached RelationshipParser instance (created lazily) */
    private relationshipParser;
    /** Cached ImageDecoder instance (created lazily) */
    private imageDecoder;
    /** Cached parser reference for relationship resolution */
    private cachedParser;
    constructor(theme: ResolvedTheme, logger?: ILogger);
    /**
     * Builds a color resolver honoring the effective color map for a slide
     * (slide p:clrMapOvr > layout p:clrMapOvr > master p:clrMap > default).
     */
    private createColorResolverForSlide;
    /**
     * Renders the background for a slide.
     * Follows the inheritance chain: slide -> layout -> master
     */
    renderBackground(ctx: CanvasRenderingContext2D, canvas: Canvas, slideNode: PptxXmlNode, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode, overrideColor?: string): void;
    /**
     * Renders the background for a slide with async support for picture backgrounds.
     * Follows the inheritance chain: slide -> layout -> master
     *
     * @param ctx Canvas 2D rendering context
     * @param canvas Canvas to render to
     * @param slideNode Slide XML node
     * @param parser PPTX parser for accessing resources
     * @param slidePath Path to the slide file (e.g., ppt/slides/slide1.xml)
     * @param layoutNode Optional layout XML node
     * @param layoutPath Optional path to the layout file (e.g., ppt/slideLayouts/slideLayout1.xml)
     * @param masterNode Optional master XML node
     * @param masterPath Optional path to the master file
     * @param overrideColor Optional background color override
     */
    renderBackgroundAsync(ctx: CanvasRenderingContext2D, canvas: Canvas, slideNode: PptxXmlNode, parser: PptxParser, slidePath: string, layoutNode?: PptxXmlNode, layoutPath?: string, masterNode?: PptxXmlNode, masterPath?: string, overrideColor?: string): Promise<void>;
    /**
     * Resolves background from the inheritance chain (slide -> layout -> master).
     * Simple version without path tracking for sync rendering.
     */
    private resolveBackgroundFromChain;
    /**
     * Resolves background from the inheritance chain with path tracking.
     * Required for async rendering to resolve picture relationships from the correct source.
     */
    private resolveBackgroundFromChainWithPath;
    /**
     * Renders background fill (sync version - no picture support).
     */
    private renderBackgroundFill;
    /**
     * Renders background fill (async version - with picture support).
     */
    private renderBackgroundFillAsync;
    /**
     * Renders gradient background (shared between sync and async).
     */
    private renderGradient;
    /**
     * Gets or creates the cached RelationshipParser instance.
     */
    private getRelationshipParser;
    /**
     * Gets or creates the cached ImageDecoder instance.
     */
    private getImageDecoder;
    /**
     * Fills the canvas with a picture background.
     *
     * @param ctx Canvas 2D rendering context
     * @param width Canvas width
     * @param height Canvas height
     * @param pictureFill Picture fill data
     * @param parser PPTX parser for accessing resources
     * @param sourcePath Path to the source file for relationship resolution
     */
    private fillPicture;
    /**
     * Parses background fill from a slide/layout/master node.
     */
    private parseBackground;
    /**
     * Parses explicit background properties.
     */
    private parseBgProperties;
    /**
     * Parses a background reference (p:bgRef) by resolving its index into the
     * theme style matrix (ECMA-376 §20.1.4.2.10 idx semantics: 0/1000 = no
     * fill, 1-999 = fillStyleLst, 1001+ = bgFillStyleLst). The bgRef child
     * color substitutes phClr placeholders, so template gradient backgrounds
     * render as real gradients.
     */
    private parseBgReference;
    /**
     * Converts a theme style-matrix fill into a ParsedBackground, substituting
     * phClr placeholders with the bgRef color.
     */
    private themeFillToBackground;
    /**
     * Parses gradient fill properties.
     */
    private parseGradientFill;
    /**
     * Fills the canvas with a solid color.
     */
    private fillSolid;
    /**
     * Fills the canvas with a linear gradient.
     */
    private fillLinearGradient;
    /**
     * Fills the canvas with a radial gradient.
     */
    private fillRadialGradient;
    /**
     * Gets the background color if it's a solid fill (for contrast calculations).
     */
    getBackgroundColor(slideNode: PptxXmlNode, layoutNode?: PptxXmlNode, masterNode?: PptxXmlNode): Rgba | undefined;
}
//# sourceMappingURL=BackgroundRenderer.d.ts.map