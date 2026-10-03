import type { ColorTransform, ResolvedColorScheme, SchemeColorType, ColorMap, ThemeStyleColor } from '../types/theme.js';
import type { Rgba } from '../types/geometry.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
/**
 * Parses a color mapping element (p:clrMap or a:overrideClrMapping) into a
 * ColorMap. Missing or invalid attributes fall back to the default mapping.
 * @param clrMapNode The p:clrMap / a:overrideClrMapping node
 * @returns The parsed color map, or undefined when no node is given
 */
export declare function parseColorMap(clrMapNode: PptxXmlNode | undefined): ColorMap | undefined;
/**
 * Computes the effective color map for a slide per ECMA-376: the slide's
 * a:overrideClrMapping wins, then the layout's, then the master's p:clrMap,
 * then the default Office mapping.
 * @param masterNode The p:sldMaster element (carries p:clrMap)
 * @param layoutNode The p:sldLayout element (may carry p:clrMapOvr)
 * @param slideNode The p:sld element (may carry p:clrMapOvr)
 * @returns The effective color map (never undefined)
 */
export declare function resolveEffectiveColorMap(masterNode?: PptxXmlNode, layoutNode?: PptxXmlNode, slideNode?: PptxXmlNode): ColorMap;
/**
 * Resolves colors from OpenXML color definitions.
 */
export declare class ColorResolver {
    private readonly colorScheme;
    private readonly colorMap;
    constructor(colorScheme?: ResolvedColorScheme, colorMap?: ColorMap);
    /**
     * Resolves a scheme color reference to an RGBA color.
     * Mappable names (bg1/tx1/bg2/tx2, accents, hyperlinks) resolve through
     * the color map; direct slot names (dk1/lt1/dk2/lt2) bypass it.
     * @param schemeColorType The scheme color name from a:schemeClr val
     * @param phClr Substitution color for val="phClr" (style placeholder)
     */
    resolveSchemeColor(schemeColorType: SchemeColorType, phClr?: Rgba): Rgba;
    /**
     * Resolves a concrete theme color slot to its scheme color.
     */
    private resolveColorSlot;
    /**
     * Parses a hex color string to RGBA.
     */
    parseHexColor(hex: string): Rgba;
    /**
     * Resolves a color from an OpenXML color element.
     * @param node Parent node containing a color child (a:schemeClr, a:srgbClr, ...)
     * @param phClr Substitution color for a:schemeClr val="phClr" placeholders
     */
    resolveColorElement(node: PptxXmlNode | undefined, phClr?: Rgba): Rgba | undefined;
    /**
     * Extracts color transforms from an XML node.
     */
    extractTransforms(node: PptxXmlNode): ColorTransform;
    /**
     * Applies color transforms to a base color.
     */
    applyTransforms(baseColor: Rgba, node: PptxXmlNode): Rgba;
    /**
     * Resolves a theme style color (from a fmtScheme style list) to RGBA.
     * Concrete colors are returned as-is; phClr placeholders are substituted
     * with the supplied color and their stored transforms applied.
     * @param styleColor The theme style color
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved color, or undefined for a phClr placeholder with
     *          no substitution color available
     */
    resolveThemeStyleColor(styleColor: ThemeStyleColor, phClr?: Rgba): Rgba | undefined;
    /**
     * Applies color transforms to a base color.
     */
    applyColorTransforms(baseColor: Rgba, transforms: ColorTransform): Rgba;
    /**
     * Converts an sRGB channel value (0-255) to linear-light (0-1) using the
     * IEC 61966-2-1 transfer function. Tint/shade transforms operate in this
     * space (ECMA-376 §20.1.2.3.32/§20.1.2.3.31 as implemented by PowerPoint).
     */
    private static srgbToLinear;
    /**
     * Converts a linear-light value (0-1) back to an sRGB channel (0-255),
     * clamping out-of-range inputs.
     */
    private static linearToSrgb;
    /**
     * Converts RGBA to HSL.
     */
    rgbaToHsl(color: Rgba): {
        h: number;
        s: number;
        l: number;
    };
    /**
     * Converts HSL to RGBA.
     */
    hslToRgba(h: number, s: number, l: number): Rgba;
    private hueToRgb;
    /**
     * Resolves a preset color name to RGBA.
     */
    resolvePresetColor(name: string): Rgba;
    /**
     * Resolves a system color name to RGBA.
     */
    resolveSystemColor(name: string): Rgba;
    /**
     * Converts RGBA to CSS color string.
     */
    rgbaToCss(color: Rgba): string;
    /**
     * Converts RGBA to hex string.
     */
    rgbaToHex(color: Rgba, includeAlpha?: boolean): string;
    /**
     * Calculates relative luminance for contrast calculations.
     */
    calculateLuminance(color: Rgba): number;
    /**
     * Determines if a color is dark (for contrast purposes).
     */
    isDarkColor(color: Rgba): boolean;
}
