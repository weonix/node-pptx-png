/**
 * Parses ppt/tableStyles.xml (a:tblStyleLst, ECMA-376 §20.1.4.2.27) into
 * resolved table styles keyed by their styleId GUID.
 *
 * Each a:tblStyle carries a whole-table background (a:tblBg) and conditional
 * formatting parts (wholeTbl, band1H, band2H, band1V, band2V, firstRow,
 * lastRow, firstCol, lastCol, and the corner parts nwCell/neCell/swCell/
 * seCell). A part consists of:
 * - a:tcStyle: cell fill (a:fill or theme-relative a:fillRef) and cell
 *   borders (a:tcBdr with left/right/top/bottom/insideH/insideV plus the
 *   diagonals tl2br/tr2bl, each a direct a:ln or a theme-relative a:lnRef;
 *   the legacy lnL/lnR/lnT/lnB/lnTlToBr/lnBlToTr spellings are accepted too)
 * - a:tcTxStyle: text bold/italic, color, and font reference
 *
 * Theme-relative references (a:fillRef, a:lnRef, phClr placeholders) are
 * resolved against the theme's a:fmtScheme style matrix at parse time, so
 * consumers receive concrete colors. Linear gradient fills are preserved as
 * gradients; radial/path gradients are approximated by their lowest-position
 * stop.
 */
import type { Rgba } from '../types/geometry.js';
import type { GradientStop } from '../types/elements.js';
import type { ResolvedTheme, ColorMap } from '../types/theme.js';
import type { PptxParser, PptxXmlNode } from '../core/PptxParser.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Fixed part path of the table styles part inside a PPTX package.
 */
export declare const TABLE_STYLES_PART_PATH = "ppt/tableStyles.xml";
/**
 * A border defined by a table style part.
 * `none` marks a border the style explicitly removes (a noFill line), which
 * must override any border a lower-precedence part defines.
 */
export interface TableStyleBorder {
    /** True when the style explicitly removes this border */
    none?: boolean;
    /** Border width in EMU */
    width?: number;
    /** Border color */
    color?: Rgba;
    /** Preset dash value (solid, dash, dot, ...) */
    dash?: string;
}
/**
 * The borders a table style part may define (a:tcBdr).
 * left/right/top/bottom apply at the edges of the part's region;
 * insideH/insideV apply to edges between cells inside the region.
 */
export interface TableStyleBorders {
    /** Left edge border (a:lnL) */
    left?: TableStyleBorder;
    /** Right edge border (a:lnR) */
    right?: TableStyleBorder;
    /** Top edge border (a:lnT) */
    top?: TableStyleBorder;
    /** Bottom edge border (a:lnB) */
    bottom?: TableStyleBorder;
    /** Inner horizontal borders (a:insideH) */
    insideH?: TableStyleBorder;
    /** Inner vertical borders (a:insideV) */
    insideV?: TableStyleBorder;
    /** Diagonal border from top-left to bottom-right (a:tl2br) */
    tlToBr?: TableStyleBorder;
    /** Diagonal border from bottom-left to top-right (a:tr2bl) */
    blToTr?: TableStyleBorder;
}
/**
 * A linear gradient cell fill with resolved stops.
 */
export interface TableStyleGradientFill {
    type: 'gradient';
    /** Gradient stops sorted ascending by position (0-1) */
    stops: GradientStop[];
    /** Gradient angle in degrees, clockwise from the positive x-axis */
    angle: number;
}
/**
 * A cell fill from a table style part: a concrete solid color, a linear
 * gradient, or an explicit no-fill that must override lower-precedence
 * fills. Radial/path gradient and pattern style fills are flattened to a
 * representative solid.
 */
export type TableStyleFill = {
    type: 'solid';
    color: Rgba;
} | TableStyleGradientFill | {
    type: 'none';
};
/**
 * Text formatting from a part's a:tcTxStyle.
 */
export interface TableStyleTextStyle {
    /** Bold flag (b="on"/"off"; undefined when unspecified) */
    bold?: boolean;
    /** Italic flag (i="on"/"off"; undefined when unspecified) */
    italic?: boolean;
    /** Text color (from the direct color child or the a:fontRef color) */
    color?: Rgba;
    /** Font reference index ('major' | 'minor' | 'none') from a:fontRef */
    fontRefIdx?: string;
    /** Font family resolved from fontRefIdx against the theme font scheme */
    fontFamily?: string;
}
/**
 * One conditional formatting part of a table style (CT_TablePartStyle).
 */
export interface TablePartStyle {
    /** Cell fill for cells in this part's region */
    fill?: TableStyleFill;
    /** Cell borders for this part's region */
    borders: TableStyleBorders;
    /** Text formatting for cells in this part's region */
    text: TableStyleTextStyle;
}
/**
 * A parsed table style (a:tblStyle) with its conditional formatting parts.
 */
export interface TableStyle {
    /** Style GUID from the styleId attribute, e.g. {5C22544A-...} */
    styleId: string;
    /** Human-readable style name */
    styleName?: string;
    /** Whole-table background fill (a:tblBg), painted beneath cell fills */
    tblBg?: TableStyleFill;
    /** Formatting for the whole table (lowest precedence) */
    wholeTbl?: TablePartStyle;
    /** First horizontal band (odd data rows, 1-based) */
    band1H?: TablePartStyle;
    /** Second horizontal band (even data rows, 1-based) */
    band2H?: TablePartStyle;
    /** First vertical band (odd data columns, 1-based) */
    band1V?: TablePartStyle;
    /** Second vertical band (even data columns, 1-based) */
    band2V?: TablePartStyle;
    /** Header row formatting */
    firstRow?: TablePartStyle;
    /** Total row formatting */
    lastRow?: TablePartStyle;
    /** First column formatting */
    firstCol?: TablePartStyle;
    /** Last column formatting */
    lastCol?: TablePartStyle;
    /** Top-left corner cell formatting (layers above firstRow/firstCol) */
    nwCell?: TablePartStyle;
    /** Top-right corner cell formatting (layers above firstRow/lastCol) */
    neCell?: TablePartStyle;
    /** Bottom-left corner cell formatting (layers above lastRow/firstCol) */
    swCell?: TablePartStyle;
    /** Bottom-right corner cell formatting (layers above lastRow/lastCol) */
    seCell?: TablePartStyle;
}
/**
 * All table styles of a presentation, keyed by styleId GUID.
 */
export interface TableStyleCollection {
    /** Parsed styles by styleId GUID */
    styles: Map<string, TableStyle>;
    /** Default style GUID from the tblStyleLst def attribute */
    defaultStyleId?: string;
}
/**
 * Configuration for TableStyleParser.
 */
export interface TableStyleParserConfig {
    /** Resolved theme for color and fmtScheme resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Parses the table styles part of a PPTX into resolved table styles.
 */
export declare class TableStyleParser {
    private readonly logger;
    private readonly theme;
    private readonly colorResolver;
    constructor(config: TableStyleParserConfig);
    /**
     * Loads and parses ppt/tableStyles.xml from an opened PPTX.
     * A missing or unreadable part yields an empty collection.
     * @param parser The opened PPTX parser
     * @returns The parsed table style collection
     */
    load(parser: PptxParser): Promise<TableStyleCollection>;
    /**
     * Parses an a:tblStyleLst node into a table style collection.
     * @param tblStyleLst The a:tblStyleLst XML node
     * @returns The parsed table style collection
     */
    parseTableStyleList(tblStyleLst: PptxXmlNode): TableStyleCollection;
    /**
     * Parses a single a:tblStyle node.
     */
    private parseTableStyle;
    /**
     * Parses one conditional formatting part (CT_TablePartStyle).
     */
    private parsePartStyle;
    /**
     * Parses a part's cell fill: either a direct a:fill or a theme-relative
     * a:fillRef (EG_ThemeableFillStyle).
     */
    private parseThemeableFill;
    /**
     * Parses a direct a:fill wrapper (noFill/solidFill/gradFill/...).
     */
    private parseDirectFill;
    /**
     * Resolves an a:fillRef against the theme fill style matrix
     * (ECMA-376 §20.1.4.2.10): idx 0/1000 means no fill, 1-999 indexes
     * fillStyleLst, 1001+ indexes bgFillStyleLst. The ref's child color
     * substitutes phClr placeholders.
     */
    private resolveFillRef;
    /**
     * Resolves a theme gradient fill to a table style fill: linear gradients
     * keep their stops (phClr substituted) and angle; radial/path gradients
     * are approximated by their lowest-position stop.
     */
    private resolveThemeGradient;
    /**
     * Flattens a theme fill to a representative solid color, substituting
     * phClr placeholders (gradients use their first stop, patterns their
     * foreground color).
     */
    private flattenThemeFill;
    /**
     * Parses a themeable border line (CT_ThemeableLineStyle): a direct a:ln
     * or a theme-relative a:lnRef.
     */
    private parseThemeableLine;
    /**
     * Parses a direct a:ln border definition (width, color, dash).
     */
    private parseDirectLine;
    /**
     * Resolves an a:lnRef against the theme lnStyleLst (ECMA-376
     * §20.1.4.2.19; 1-based index, 0 means no line). The ref's child color
     * substitutes phClr placeholders.
     */
    private resolveLineRef;
    /**
     * Flattens a theme line style to a table style border, substituting phClr
     * placeholders in the line's fill.
     */
    private flattenThemeLineStyle;
    /**
     * Parses a part's a:tcTxStyle (bold/italic flags, color, font reference).
     */
    private parseTcTxStyle;
}
/**
 * Parses an a:gradFill element into a table cell fill. Linear gradients
 * (a:lin, or no explicit shading child) keep their stops sorted ascending
 * by position with the angle from a:lin@ang; radial/path gradients (a:path)
 * are approximated by their lowest-position stop as a solid.
 * @param gradFill The a:gradFill XML node
 * @param colorResolver Resolver for the stop color elements
 * @param logger Optional logger for approximation diagnostics
 * @returns The parsed fill, or undefined when no stop color resolves
 */
export declare function parseTableGradientFill(gradFill: PptxXmlNode, colorResolver: ColorResolver, logger?: ILogger): {
    type: 'solid';
    color: Rgba;
} | TableStyleGradientFill | undefined;
/**
 * Creates a TableStyleParser instance.
 */
export declare function createTableStyleParser(theme: ResolvedTheme, colorMap?: ColorMap, logger?: ILogger): TableStyleParser;
