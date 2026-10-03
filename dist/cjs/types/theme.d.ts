import type { Rgba } from './geometry.js';
/**
 * Resolved color scheme with all 12 standard theme colors.
 * Values are fully computed RGBA colors.
 */
export interface ResolvedColorScheme {
    /** Dark 1 - typically black or near-black (tx1/bg1 alternate) */
    dark1: Rgba;
    /** Light 1 - typically white or near-white (tx2/bg2 alternate) */
    light1: Rgba;
    /** Dark 2 - secondary dark color */
    dark2: Rgba;
    /** Light 2 - secondary light color */
    light2: Rgba;
    /** Accent 1 - primary accent color */
    accent1: Rgba;
    /** Accent 2 */
    accent2: Rgba;
    /** Accent 3 */
    accent3: Rgba;
    /** Accent 4 */
    accent4: Rgba;
    /** Accent 5 */
    accent5: Rgba;
    /** Accent 6 */
    accent6: Rgba;
    /** Hyperlink color */
    hyperlink: Rgba;
    /** Followed hyperlink color */
    followedHyperlink: Rgba;
}
/**
 * Resolved font scheme with heading and body fonts.
 */
export interface ResolvedFontScheme {
    /** Major font family for headings/titles */
    majorFont: string;
    /** Minor font family for body text */
    minorFont: string;
    /** East Asian major font (optional) */
    majorFontEastAsian?: string;
    /** East Asian minor font (optional) */
    minorFontEastAsian?: string;
    /** Complex script major font (optional) */
    majorFontComplexScript?: string;
    /** Complex script minor font (optional) */
    minorFontComplexScript?: string;
}
/**
 * Fully resolved theme with computed colors and fonts.
 */
export interface ResolvedTheme {
    /** Resolved color scheme */
    colors: ResolvedColorScheme;
    /** Resolved font scheme */
    fonts: ResolvedFontScheme;
    /**
     * Background fill style colors flattened to representative colors
     * (legacy field; prefer formatScheme.backgroundFillStyles).
     */
    backgroundFillStyles?: Rgba[];
    /** Full a:fmtScheme style matrix (fill/line/effect/background style lists) */
    formatScheme?: ThemeFormatScheme;
}
/**
 * Scheme color references used in OpenXML.
 */
export type SchemeColorType = 'dk1' | 'lt1' | 'dk2' | 'lt2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink' | 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'phClr';
/**
 * Color transform types that can be applied to a base color.
 */
export interface ColorTransform {
    /** Tint - lighten color toward white (0-100000 = 0-100%) */
    tint?: number;
    /** Shade - darken color toward black (0-100000 = 0-100%) */
    shade?: number;
    /** Saturation modulation (0-100000 = 0-100% scale) */
    satMod?: number;
    /** Luminance modulation (0-100000 = 0-100% scale) */
    lumMod?: number;
    /** Luminance offset (-100000 to 100000 = -100% to 100%) */
    lumOff?: number;
    /** Hue modulation (0-100000 = 0-100% scale) */
    hueMod?: number;
    /** Hue offset (-360 to 360 degrees, in 60000ths) */
    hueOff?: number;
    /** Alpha/transparency (0-100000 = 0-100%) */
    alpha?: number;
}
/**
 * A concrete slot in the theme color scheme that scheme color names map to.
 */
export type ThemeColorSlot = 'dk1' | 'lt1' | 'dk2' | 'lt2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink';
/**
 * Color map from p:clrMap / a:overrideClrMapping (ECMA-376 §19.3.1.6).
 * Maps the mappable scheme color names (bg1/tx1/bg2/tx2/accents/hyperlinks)
 * to concrete theme color slots.
 */
export interface ColorMap {
    /** Background 1 mapping (default lt1) */
    bg1: ThemeColorSlot;
    /** Text 1 mapping (default dk1) */
    tx1: ThemeColorSlot;
    /** Background 2 mapping (default lt2) */
    bg2: ThemeColorSlot;
    /** Text 2 mapping (default dk2) */
    tx2: ThemeColorSlot;
    /** Accent 1 mapping */
    accent1: ThemeColorSlot;
    /** Accent 2 mapping */
    accent2: ThemeColorSlot;
    /** Accent 3 mapping */
    accent3: ThemeColorSlot;
    /** Accent 4 mapping */
    accent4: ThemeColorSlot;
    /** Accent 5 mapping */
    accent5: ThemeColorSlot;
    /** Accent 6 mapping */
    accent6: ThemeColorSlot;
    /** Hyperlink mapping */
    hlink: ThemeColorSlot;
    /** Followed hyperlink mapping */
    folHlink: ThemeColorSlot;
}
/**
 * Default color map used when a master declares no p:clrMap.
 * Matches the standard Office mapping (and the library's previous
 * hardcoded bg1->lt1 / tx1->dk1 behavior).
 */
export declare const DEFAULT_COLOR_MAP: ColorMap;
/**
 * A color inside a theme style list entry. Either a concrete resolved color
 * (transforms already applied) or the symbolic a:schemeClr val="phClr"
 * placeholder whose transforms are applied to the substituted color at use
 * time (the color supplied by a style reference such as a:fillRef).
 */
export interface ThemeStyleColor {
    /** True when this is the phClr placeholder resolved at use time */
    isPhClr?: boolean;
    /** Concrete RGBA color (present when not phClr) */
    color?: Rgba;
    /** Transforms to apply to the substituted placeholder color */
    transforms?: ColorTransform;
}
/**
 * A gradient stop in a theme style fill.
 */
export interface ThemeGradientStop {
    /** Stop position 0-1 */
    position: number;
    /** Stop color (may be a phClr placeholder) */
    color: ThemeStyleColor;
}
/**
 * Solid fill variant in a theme style list.
 */
export interface ThemeSolidFill {
    type: 'solid';
    /** Fill color (may be a phClr placeholder) */
    color: ThemeStyleColor;
}
/**
 * Gradient fill variant in a theme style list.
 */
export interface ThemeGradientFill {
    type: 'gradient';
    /** Gradient stops sorted by position */
    stops: ThemeGradientStop[];
    /** Linear gradient angle in degrees (clockwise from positive x-axis) */
    angle?: number;
    /** Whether this is a path/radial gradient */
    isRadial?: boolean;
    /** Center point for radial gradients (0-1) */
    centerX?: number;
    /** Center point for radial gradients (0-1) */
    centerY?: number;
}
/**
 * Pattern fill variant in a theme style list.
 */
export interface ThemePatternFill {
    type: 'pattern';
    /** Pattern preset name */
    preset: string;
    /** Foreground color (may be a phClr placeholder) */
    foregroundColor: ThemeStyleColor;
    /** Background color (may be a phClr placeholder) */
    backgroundColor: ThemeStyleColor;
}
/**
 * No-fill variant in a theme style list.
 */
export interface ThemeNoFill {
    type: 'none';
}
/**
 * A fill definition from the theme style matrix (fillStyleLst /
 * bgFillStyleLst), preserving the fill variant and phClr placeholders.
 */
export type ThemeFill = ThemeSolidFill | ThemeGradientFill | ThemePatternFill | ThemeNoFill;
/**
 * A line style from the theme lnStyleLst.
 */
export interface ThemeLineStyle {
    /** Line width in EMU */
    width?: number;
    /** Raw OOXML cap value ('flat' | 'rnd' | 'sq') */
    cap?: string;
    /** Compound line type (sng, dbl, ...) */
    compound?: string;
    /** Pen alignment (ctr, in) */
    align?: string;
    /** Preset dash value (solid, dash, dot, ...) */
    dash?: string;
    /** Line join type */
    join?: 'miter' | 'round' | 'bevel';
    /** Line fill (usually solid with phClr) */
    fill?: ThemeFill;
}
/**
 * Parsed a:outerShdw from a theme effect style.
 */
export interface ThemeOuterShadow {
    /** Blur radius in EMU */
    blurRadius?: number;
    /** Shadow distance in EMU */
    distance?: number;
    /** Shadow direction in degrees */
    direction?: number;
    /** Shadow alignment (tl, ctr, ...) */
    alignment?: string;
    /** Whether the shadow rotates with the shape */
    rotateWithShape?: boolean;
    /** Shadow color (may be a phClr placeholder) */
    color?: ThemeStyleColor;
}
/**
 * An effect style from the theme effectStyleLst.
 * Parsed and stored for style references; effect rendering is out of scope.
 */
export interface ThemeEffectStyle {
    /** Outer shadow effect, when present */
    outerShadow?: ThemeOuterShadow;
}
/**
 * The theme format scheme (a:fmtScheme style matrix, ECMA-376 §20.1.4.1.14).
 * Style references (a:fillRef/a:lnRef/a:effectRef, §20.1.4.2) index into
 * these lists 1-based; fill indices 1001+ index backgroundFillStyles.
 */
export interface ThemeFormatScheme {
    /** Fill style list (a:fillStyleLst) */
    fillStyles: ThemeFill[];
    /** Line style list (a:lnStyleLst) */
    lineStyles: ThemeLineStyle[];
    /** Effect style list (a:effectStyleLst) */
    effectStyles: ThemeEffectStyle[];
    /** Background fill style list (a:bgFillStyleLst) */
    backgroundFillStyles: ThemeFill[];
}
/**
 * Default Office theme colors.
 */
export declare const DEFAULT_OFFICE_COLORS: ResolvedColorScheme;
/**
 * Default font scheme.
 */
export declare const DEFAULT_FONT_SCHEME: ResolvedFontScheme;
/**
 * Default resolved theme.
 */
export declare const DEFAULT_THEME: ResolvedTheme;
