/**
 * Parses DrawingML effect lists (a:effectLst) from shape properties.
 * Plain outer shadows are rendered by ShapeRenderer via canvas shadow
 * state; perspective/picture-fill shadows, inner shadow, glow, soft edge,
 * and reflection are rendered by EffectRenderer from these records.
 */
import type { Rgba } from '../types/geometry.js';
import type { ResolvedTheme, ColorMap, ThemeOuterShadow } from '../types/theme.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Outer shadow effect parsed from a:outerShdw (ECMA-376 §20.1.8.45).
 */
export interface OuterShadowEffect {
    type: 'outerShadow';
    /** Blur radius in EMU (default 0) */
    blurRadius: number;
    /** Shadow distance from the shape in EMU (default 0) */
    distance: number;
    /**
     * Shadow direction in degrees, clockwise from the positive x-axis in
     * y-down slide space (default 0): 0 = right, 90 = down.
     */
    direction: number;
    /** Horizontal shadow scale factor (1 = 100%, default 1) */
    scaleX: number;
    /** Vertical shadow scale factor (1 = 100%, default 1) */
    scaleY: number;
    /** Horizontal skew angle in degrees (perspective presets, default 0) */
    skewX: number;
    /** Vertical skew angle in degrees (default 0) */
    skewY: number;
    /** Whether the shadow rotates with the shape (default true) */
    rotateWithShape: boolean;
    /** Shadow alignment (tl, t, tr, l, ctr, r, bl, b, br) */
    alignment?: string;
    /** Shadow color with alpha applied */
    color: Rgba;
}
/**
 * Inner shadow effect parsed from a:innerShdw (ECMA-376 §20.1.8.40).
 * Rendered by EffectRenderer.
 */
export interface InnerShadowEffect {
    type: 'innerShadow';
    /** Blur radius in EMU (default 0) */
    blurRadius: number;
    /** Shadow distance in EMU (default 0) */
    distance: number;
    /** Shadow direction in degrees (default 0) */
    direction: number;
    /** Shadow color with alpha applied */
    color: Rgba;
}
/**
 * Glow effect parsed from a:glow (ECMA-376 §20.1.8.32).
 * Rendered by EffectRenderer.
 */
export interface GlowEffect {
    type: 'glow';
    /** Glow radius in EMU (default 0) */
    radius: number;
    /** Glow color with alpha applied */
    color: Rgba;
}
/**
 * Soft edge effect parsed from a:softEdge (ECMA-376 §20.1.8.53).
 * Rendered by EffectRenderer.
 */
export interface SoftEdgeEffect {
    type: 'softEdge';
    /** Soft edge radius in EMU */
    radius: number;
}
/**
 * Reflection effect parsed from a:reflection (ECMA-376 §20.1.8.50).
 * Rendered by EffectRenderer.
 */
export interface ReflectionEffect {
    type: 'reflection';
    /** Blur radius in EMU (default 0) */
    blurRadius: number;
    /** Reflection distance in EMU (default 0) */
    distance: number;
    /** Reflection direction in degrees (default 0) */
    direction: number;
    /** Fade direction in degrees (default 90) */
    fadeDirection: number;
    /** Starting alpha as a decimal 0-1 (default 1) */
    startAlpha: number;
    /** Ending alpha as a decimal 0-1 (default 0) */
    endAlpha: number;
    /** Start position of the alpha gradient as a decimal 0-1 (default 0) */
    startPosition: number;
    /** End position of the alpha gradient as a decimal 0-1 (default 1) */
    endPosition: number;
    /** Horizontal scale factor (1 = 100%, default 1) */
    scaleX: number;
    /** Vertical scale factor (1 = 100%, default 1; -1 flips downward) */
    scaleY: number;
    /** Horizontal skew angle in degrees (default 0) */
    skewX: number;
    /** Vertical skew angle in degrees (default 0) */
    skewY: number;
    /** Whether the reflection rotates with the shape (default true) */
    rotateWithShape: boolean;
    /** Reflection alignment (default b) */
    alignment?: string;
}
/**
 * The effects parsed from one a:effectLst. An empty object is meaningful:
 * it records an explicit (possibly empty) effect list, which per ECMA-376
 * overrides any style-referenced effects (a:effectRef).
 */
export interface ParsedEffectList {
    /** Outer shadow (a:outerShdw) */
    outerShadow?: OuterShadowEffect;
    /** Inner shadow (a:innerShdw) */
    innerShadow?: InnerShadowEffect;
    /** Glow (a:glow) */
    glow?: GlowEffect;
    /** Soft edge (a:softEdge) */
    softEdge?: SoftEdgeEffect;
    /** Reflection (a:reflection) */
    reflection?: ReflectionEffect;
}
/**
 * Computes the x/y offset of a shadow from its distance and direction.
 * Direction is measured clockwise from the positive x-axis in y-down slide
 * space (matching a:lin gradient angles), so the offset maps directly to
 * (cos, sin) without adjustment.
 * @param distance Shadow distance (any length unit; the result is in the same unit)
 * @param directionDegrees Shadow direction in degrees
 * @returns Shadow offset in the same unit as the distance
 */
export declare function computeShadowOffset(distance: number, directionDegrees: number): {
    dx: number;
    dy: number;
};
/**
 * Configuration for EffectParser.
 */
export interface EffectParserConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Parses a:effectLst elements into typed effect records.
 */
export declare class EffectParser {
    private readonly logger;
    private readonly colorResolver;
    constructor(config: EffectParserConfig);
    /**
     * Parses the a:effectLst of a shape properties node (p:spPr).
     * @param spPr Shape properties node
     * @returns The parsed effects, an empty record for an explicit empty
     *          a:effectLst, or undefined when the node has no a:effectLst
     *          (so style-referenced effects may still apply)
     */
    parseShapeEffects(spPr: PptxXmlNode | undefined): ParsedEffectList | undefined;
    /**
     * Parses an a:effectLst node into typed effect records.
     * @param effectLst The a:effectLst node (an empty node yields an empty record)
     */
    parseEffectList(effectLst: PptxXmlNode | undefined): ParsedEffectList;
    /**
     * Resolves a theme effect-style outer shadow (from a:effectStyleLst) into
     * a renderable effect, substituting phClr placeholders in the shadow color
     * with the color supplied by the style reference (a:effectRef child color).
     * @param shadow The theme outer shadow definition
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved outer shadow effect
     */
    resolveThemeOuterShadow(shadow: ThemeOuterShadow, phClr?: Rgba): OuterShadowEffect;
    /**
     * Parses an a:outerShdw element (defaults per ECMA-376 §20.1.8.45).
     */
    private parseOuterShadow;
    /**
     * Parses an a:innerShdw element (defaults per ECMA-376 §20.1.8.40).
     */
    private parseInnerShadow;
    /**
     * Parses an a:glow element (defaults per ECMA-376 §20.1.8.32).
     */
    private parseGlow;
    /**
     * Parses an a:softEdge element (ECMA-376 §20.1.8.53).
     */
    private parseSoftEdge;
    /**
     * Parses an a:reflection element (defaults per ECMA-376 §20.1.8.50).
     */
    private parseReflection;
    /**
     * Resolves the color child of an effect element (transforms, including
     * a:alpha, are applied by the color resolver). Falls back to opaque black.
     */
    private parseEffectColor;
}
/**
 * Default effect parser factory.
 */
export declare function createEffectParser(theme: ResolvedTheme, logger?: ILogger): EffectParser;
