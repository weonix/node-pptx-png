import type { ResolvedTheme, ResolvedColorScheme } from '../types/index.js';
import type { ThemeFormatScheme } from '../types/theme.js';
import type { PptxParser, PptxXmlNode, ThemeData } from '../core/PptxParser.js';
import { ColorResolver } from './ColorResolver.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Resolves theme colors, fonts, and effects from PPTX theme data.
 */
export declare class ThemeResolver {
    private readonly logger;
    private colorResolver;
    constructor(logger?: ILogger);
    /**
     * Resolves the complete theme from PPTX parser.
     */
    resolveTheme(parser: PptxParser): Promise<ResolvedTheme>;
    /**
     * Parses theme data into resolved theme.
     */
    parseTheme(themeData: ThemeData): ResolvedTheme;
    /**
     * Parses the a:fmtScheme style matrix into full style lists.
     * Style references (a:fillRef/a:lnRef/a:effectRef) index into these lists;
     * phClr placeholders are kept symbolic for substitution at use time.
     */
    parseFormatScheme(fmtScheme: PptxXmlNode | undefined, colorScheme: ResolvedColorScheme): ThemeFormatScheme | undefined;
    /**
     * Parses a fill style list (a:fillStyleLst / a:bgFillStyleLst) preserving
     * document order, which style reference indices depend on.
     */
    private parseFillStyleList;
    /**
     * Parses a single fill element from a style list into a ThemeFill.
     * Unsupported variants (blip/group fills) degrade to a solid phClr fill so
     * list positions stay aligned with style reference indices.
     */
    private parseThemeFill;
    /**
     * Parses a gradient fill from a style list, keeping phClr stops symbolic.
     */
    private parseThemeGradientFill;
    /**
     * Parses a color child of a style list element into a ThemeStyleColor.
     * a:schemeClr val="phClr" is kept symbolic with its transforms; all other
     * colors are resolved to concrete RGBA at parse time.
     */
    private parseThemeStyleColor;
    /**
     * Parses the line style list (a:lnStyleLst).
     */
    private parseLineStyleList;
    /**
     * Parses the effect style list (a:effectStyleLst). Effects are parsed and
     * stored for style references; rendering them is out of scope.
     */
    private parseEffectStyleList;
    /**
     * Resolves the color scheme from theme XML.
     */
    private resolveColorScheme;
    /**
     * Extracts a color from a theme color element.
     */
    private extractThemeColor;
    /**
     * Resolves the font scheme from theme XML.
     */
    private resolveFontScheme;
    /**
     * Extracts a font family from a font element.
     */
    private extractFontFamily;
    /**
     * Resolves background fill styles from the format scheme.
     */
    private resolveBackgroundFillStyles;
    /**
     * Gets the color resolver for this theme.
     */
    getColorResolver(): ColorResolver;
    /**
     * Creates a color resolver with the given color scheme.
     */
    createColorResolver(colorScheme: ResolvedColorScheme): ColorResolver;
}
