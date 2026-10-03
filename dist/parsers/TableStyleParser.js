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
import { Colors } from '../types/geometry.js';
import { getXmlAttr, getXmlChild, getXmlChildren, hasXmlChild, getOrderedChildren, } from '../core/PptxParser.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import { percentageToDecimal } from '../core/UnitConverter.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Fixed part path of the table styles part inside a PPTX package.
 */
export const TABLE_STYLES_PART_PATH = 'ppt/tableStyles.xml';
/**
 * Default border width in EMU (12700 EMU = 1 point) used when a style
 * border omits an explicit width.
 */
const DEFAULT_STYLE_BORDER_WIDTH_EMU = 12700;
/**
 * The style parts a tblStyle may define, mapped from their XML tag names.
 */
const STYLE_PART_TAGS = [
    ['a:wholeTbl', 'wholeTbl'],
    ['a:band1H', 'band1H'],
    ['a:band2H', 'band2H'],
    ['a:band1V', 'band1V'],
    ['a:band2V', 'band2V'],
    ['a:firstRow', 'firstRow'],
    ['a:lastRow', 'lastRow'],
    ['a:firstCol', 'firstCol'],
    ['a:lastCol', 'lastCol'],
    ['a:nwCell', 'nwCell'],
    ['a:neCell', 'neCell'],
    ['a:swCell', 'swCell'],
    ['a:seCell', 'seCell'],
];
/**
 * Parses the table styles part of a PPTX into resolved table styles.
 */
export class TableStyleParser {
    logger;
    theme;
    colorResolver;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'TableStyleParser');
        this.theme = config.theme;
        this.colorResolver = new ColorResolver(config.theme.colors, config.colorMap);
    }
    /**
     * Loads and parses ppt/tableStyles.xml from an opened PPTX.
     * A missing or unreadable part yields an empty collection.
     * @param parser The opened PPTX parser
     * @returns The parsed table style collection
     */
    async load(parser) {
        const empty = { styles: new Map() };
        let root;
        try {
            if (!parser.fileExists(TABLE_STYLES_PART_PATH)) {
                this.logger.debug('No tableStyles.xml in package');
                return empty;
            }
            const ordered = await parser.readXmlOrdered(TABLE_STYLES_PART_PATH);
            root = getOrderedChildren(ordered, ['a:tblStyleLst'])[0]?.node;
        }
        catch (error) {
            this.logger.debug('Failed to read tableStyles.xml', {
                error: error instanceof Error ? error.message : String(error),
            });
            return empty;
        }
        if (!root) {
            this.logger.debug('tableStyles.xml has no a:tblStyleLst root');
            return empty;
        }
        return this.parseTableStyleList(root);
    }
    /**
     * Parses an a:tblStyleLst node into a table style collection.
     * @param tblStyleLst The a:tblStyleLst XML node
     * @returns The parsed table style collection
     */
    parseTableStyleList(tblStyleLst) {
        const styles = new Map();
        const defaultStyleId = getXmlAttr(tblStyleLst, 'def');
        for (const styleNode of getXmlChildren(tblStyleLst, 'a:tblStyle')) {
            const style = this.parseTableStyle(styleNode);
            if (style) {
                styles.set(style.styleId, style);
            }
        }
        this.logger.debug('Parsed table styles', { count: styles.size, defaultStyleId });
        return { styles, defaultStyleId };
    }
    /**
     * Parses a single a:tblStyle node.
     */
    parseTableStyle(styleNode) {
        const styleId = getXmlAttr(styleNode, 'styleId');
        if (!styleId) {
            this.logger.debug('Skipping table style without styleId');
            return undefined;
        }
        const style = {
            styleId,
            styleName: getXmlAttr(styleNode, 'styleName'),
        };
        // a:tblBg (CT_TableBackgroundStyle) carries the same fill/fillRef
        // choice as a:tcStyle
        const tblBg = getXmlChild(styleNode, 'a:tblBg');
        if (tblBg && typeof tblBg === 'object') {
            style.tblBg = this.parseThemeableFill(tblBg);
        }
        for (const [tag, key] of STYLE_PART_TAGS) {
            const partNode = getXmlChild(styleNode, tag);
            const part = this.parsePartStyle(partNode);
            if (part) {
                style[key] = part;
            }
        }
        return style;
    }
    /**
     * Parses one conditional formatting part (CT_TablePartStyle).
     */
    parsePartStyle(partNode) {
        if (!partNode || typeof partNode !== 'object')
            return undefined;
        const part = { borders: {}, text: {} };
        const tcStyle = getXmlChild(partNode, 'a:tcStyle');
        if (tcStyle && typeof tcStyle === 'object') {
            part.fill = this.parseThemeableFill(tcStyle);
            const tcBdr = getXmlChild(tcStyle, 'a:tcBdr');
            if (tcBdr && typeof tcBdr === 'object') {
                // ECMA-376 names the tcBdr children left/right/top/bottom/tl2br/
                // tr2bl; the tcPr-style lnL/lnR/... spellings occur in the wild
                const pick = (specTag, altTag) => getXmlChild(tcBdr, specTag) ?? getXmlChild(tcBdr, altTag);
                part.borders = {
                    left: this.parseThemeableLine(pick('a:left', 'a:lnL')),
                    right: this.parseThemeableLine(pick('a:right', 'a:lnR')),
                    top: this.parseThemeableLine(pick('a:top', 'a:lnT')),
                    bottom: this.parseThemeableLine(pick('a:bottom', 'a:lnB')),
                    insideH: this.parseThemeableLine(getXmlChild(tcBdr, 'a:insideH')),
                    insideV: this.parseThemeableLine(getXmlChild(tcBdr, 'a:insideV')),
                    tlToBr: this.parseThemeableLine(pick('a:tl2br', 'a:lnTlToBr')),
                    blToTr: this.parseThemeableLine(pick('a:tr2bl', 'a:lnBlToTr')),
                };
            }
        }
        part.text = this.parseTcTxStyle(getXmlChild(partNode, 'a:tcTxStyle'));
        return part;
    }
    /**
     * Parses a part's cell fill: either a direct a:fill or a theme-relative
     * a:fillRef (EG_ThemeableFillStyle).
     */
    parseThemeableFill(tcStyle) {
        const fillNode = getXmlChild(tcStyle, 'a:fill');
        if (fillNode !== undefined) {
            return this.parseDirectFill(fillNode);
        }
        const fillRef = getXmlChild(tcStyle, 'a:fillRef');
        if (fillRef !== undefined) {
            return this.resolveFillRef(fillRef);
        }
        return undefined;
    }
    /**
     * Parses a direct a:fill wrapper (noFill/solidFill/gradFill/...).
     */
    parseDirectFill(fillNode) {
        if (typeof fillNode !== 'object')
            return undefined;
        if (hasXmlChild(fillNode, 'a:noFill')) {
            return { type: 'none' };
        }
        const solidFill = getXmlChild(fillNode, 'a:solidFill');
        if (solidFill) {
            const color = this.colorResolver.resolveColorElement(solidFill);
            return color ? { type: 'solid', color } : undefined;
        }
        // Linear gradients are preserved; radial/path gradients are
        // approximated by their lowest-position stop
        const gradFill = getXmlChild(fillNode, 'a:gradFill');
        if (gradFill && typeof gradFill === 'object') {
            return parseTableGradientFill(gradFill, this.colorResolver, this.logger);
        }
        this.logger.debug('Unsupported table style fill type ignored');
        return undefined;
    }
    /**
     * Resolves an a:fillRef against the theme fill style matrix
     * (ECMA-376 §20.1.4.2.10): idx 0/1000 means no fill, 1-999 indexes
     * fillStyleLst, 1001+ indexes bgFillStyleLst. The ref's child color
     * substitutes phClr placeholders.
     */
    resolveFillRef(fillRef) {
        if (typeof fillRef !== 'object')
            return undefined;
        const idx = parseInt(getXmlAttr(fillRef, 'idx') ?? '0', 10);
        const phClr = this.colorResolver.resolveColorElement(fillRef);
        // A malformed idx leaves the part's fill unspecified (so lower layers
        // still apply); only a valid 0/1000 idx is an explicit removal.
        if (!Number.isFinite(idx)) {
            this.logger.debug('Ignoring malformed a:fillRef idx');
            return undefined;
        }
        if (idx <= 0 || idx === 1000) {
            return { type: 'none' };
        }
        const formatScheme = this.theme.formatScheme;
        const themeFill = idx < 1000
            ? formatScheme?.fillStyles[idx - 1]
            : formatScheme?.backgroundFillStyles[idx - 1001];
        if (!themeFill) {
            // No parsed style matrix (or out-of-range index): degrade to a solid
            // fill of the reference color so the style still renders
            this.logger.debug('Table style fillRef index not found in theme, using ref color', { idx });
            return phClr ? { type: 'solid', color: phClr } : undefined;
        }
        if (themeFill.type === 'none')
            return { type: 'none' };
        if (themeFill.type === 'gradient') {
            return this.resolveThemeGradient(themeFill, phClr);
        }
        const color = this.flattenThemeFill(themeFill, phClr);
        return color ? { type: 'solid', color } : undefined;
    }
    /**
     * Resolves a theme gradient fill to a table style fill: linear gradients
     * keep their stops (phClr substituted) and angle; radial/path gradients
     * are approximated by their lowest-position stop.
     */
    resolveThemeGradient(themeFill, phClr) {
        const stops = [];
        for (const stop of themeFill.stops) {
            const color = this.colorResolver.resolveThemeStyleColor(stop.color, phClr);
            if (color) {
                stops.push({ position: stop.position, color });
            }
        }
        if (stops.length === 0)
            return undefined;
        stops.sort((a, b) => a.position - b.position);
        const first = stops[0];
        if (!first)
            return undefined;
        if (themeFill.isRadial || stops.length === 1) {
            this.logger.debug('Table style radial gradient approximated by lowest-pos stop');
            return { type: 'solid', color: first.color };
        }
        return { type: 'gradient', stops, angle: themeFill.angle ?? 0 };
    }
    /**
     * Flattens a theme fill to a representative solid color, substituting
     * phClr placeholders (gradients use their first stop, patterns their
     * foreground color).
     */
    flattenThemeFill(themeFill, phClr) {
        switch (themeFill.type) {
            case 'solid':
                return this.colorResolver.resolveThemeStyleColor(themeFill.color, phClr);
            case 'gradient': {
                const firstStop = themeFill.stops[0];
                return firstStop
                    ? this.colorResolver.resolveThemeStyleColor(firstStop.color, phClr)
                    : undefined;
            }
            case 'pattern':
                return this.colorResolver.resolveThemeStyleColor(themeFill.foregroundColor, phClr);
            case 'none':
                return undefined;
        }
    }
    /**
     * Parses a themeable border line (CT_ThemeableLineStyle): a direct a:ln
     * or a theme-relative a:lnRef.
     */
    parseThemeableLine(lineParent) {
        if (!lineParent || typeof lineParent !== 'object')
            return undefined;
        if (hasXmlChild(lineParent, 'a:ln')) {
            return this.parseDirectLine(getXmlChild(lineParent, 'a:ln'));
        }
        if (hasXmlChild(lineParent, 'a:lnRef')) {
            return this.resolveLineRef(getXmlChild(lineParent, 'a:lnRef'));
        }
        return undefined;
    }
    /**
     * Parses a direct a:ln border definition (width, color, dash).
     */
    parseDirectLine(lnNode) {
        // A self-closed <a:ln/> parses to '' — treat as a default-width border
        if (!lnNode || typeof lnNode !== 'object') {
            return { width: DEFAULT_STYLE_BORDER_WIDTH_EMU, color: { ...Colors.black } };
        }
        if (hasXmlChild(lnNode, 'a:noFill')) {
            return { none: true };
        }
        const wAttr = getXmlAttr(lnNode, 'w');
        const width = wAttr !== undefined ? parseInt(wAttr, 10) : DEFAULT_STYLE_BORDER_WIDTH_EMU;
        const solidFill = getXmlChild(lnNode, 'a:solidFill');
        const color = solidFill
            ? this.colorResolver.resolveColorElement(solidFill)
            : { ...Colors.black };
        const prstDash = getXmlChild(lnNode, 'a:prstDash');
        const dash = prstDash ? getXmlAttr(prstDash, 'val') : undefined;
        return { width, color: color ?? { ...Colors.black }, dash };
    }
    /**
     * Resolves an a:lnRef against the theme lnStyleLst (ECMA-376
     * §20.1.4.2.19; 1-based index, 0 means no line). The ref's child color
     * substitutes phClr placeholders.
     */
    resolveLineRef(lnRef) {
        if (!lnRef || typeof lnRef !== 'object')
            return undefined;
        const idx = parseInt(getXmlAttr(lnRef, 'idx') ?? '0', 10);
        const phClr = this.colorResolver.resolveColorElement(lnRef);
        if (!Number.isFinite(idx) || idx <= 0) {
            return { none: true };
        }
        const lineStyle = this.theme.formatScheme?.lineStyles[idx - 1];
        if (!lineStyle) {
            this.logger.debug('Table style lnRef index not found in theme, using ref color', { idx });
            return phClr ? { width: DEFAULT_STYLE_BORDER_WIDTH_EMU, color: phClr } : undefined;
        }
        return this.flattenThemeLineStyle(lineStyle, phClr);
    }
    /**
     * Flattens a theme line style to a table style border, substituting phClr
     * placeholders in the line's fill.
     */
    flattenThemeLineStyle(style, phClr) {
        if (style.fill?.type === 'none') {
            return { none: true };
        }
        const color = style.fill ? this.flattenThemeFill(style.fill, phClr) : undefined;
        return {
            width: style.width ?? DEFAULT_STYLE_BORDER_WIDTH_EMU,
            color: color ?? phClr ?? { ...Colors.black },
            dash: style.dash,
        };
    }
    /**
     * Parses a part's a:tcTxStyle (bold/italic flags, color, font reference).
     */
    parseTcTxStyle(txStyleNode) {
        if (!txStyleNode || typeof txStyleNode !== 'object')
            return {};
        const text = {
            bold: parseOnOff(getXmlAttr(txStyleNode, 'b')),
            italic: parseOnOff(getXmlAttr(txStyleNode, 'i')),
            // EG_ColorChoice: the color element is a direct child of tcTxStyle
            color: this.colorResolver.resolveColorElement(txStyleNode),
        };
        const fontRef = getXmlChild(txStyleNode, 'a:fontRef');
        if (fontRef && typeof fontRef === 'object') {
            text.fontRefIdx = getXmlAttr(fontRef, 'idx');
            if (text.fontRefIdx === 'major') {
                text.fontFamily = this.theme.fonts.majorFont;
            }
            else if (text.fontRefIdx === 'minor') {
                text.fontFamily = this.theme.fonts.minorFont;
            }
            // The fontRef's child color is a fallback text color
            text.color ??= this.colorResolver.resolveColorElement(fontRef);
        }
        return text;
    }
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
export function parseTableGradientFill(gradFill, colorResolver, logger) {
    const gsLst = getXmlChild(gradFill, 'a:gsLst');
    const stops = [];
    for (const gs of getXmlChildren(gsLst, 'a:gs')) {
        const posAttr = getXmlAttr(gs, 'pos');
        const rawPos = posAttr !== undefined ? parseInt(posAttr, 10) : 0;
        // Malformed pos values must not produce NaN stop positions (they crash
        // canvas addColorStop and drop the whole table)
        const position = Number.isFinite(rawPos)
            ? Math.min(1, Math.max(0, percentageToDecimal(rawPos)))
            : 0;
        const color = colorResolver.resolveColorElement(gs);
        if (color) {
            stops.push({ position, color });
        }
    }
    if (stops.length === 0)
        return undefined;
    stops.sort((a, b) => a.position - b.position);
    const first = stops[0];
    if (!first)
        return undefined;
    if (stops.length === 1) {
        return { type: 'solid', color: first.color };
    }
    // Radial/path gradients are approximated by their lowest-position stop
    if (hasXmlChild(gradFill, 'a:path')) {
        logger?.debug('Table radial/path gradient approximated by lowest-pos stop');
        return { type: 'solid', color: first.color };
    }
    const lin = getXmlChild(gradFill, 'a:lin');
    const angAttr = lin ? getXmlAttr(lin, 'ang') : undefined;
    const rawAng = angAttr !== undefined ? parseInt(angAttr, 10) / 60000 : 0;
    const angle = Number.isFinite(rawAng) ? rawAng : 0;
    return { type: 'gradient', stops, angle };
}
/**
 * Parses an ST_OnOffStyleType attribute value ('on'/'off'/'def').
 * @param value The raw attribute value
 * @returns true for 'on', false for 'off', undefined otherwise
 */
function parseOnOff(value) {
    if (value === 'on')
        return true;
    if (value === 'off')
        return false;
    return undefined;
}
/**
 * Creates a TableStyleParser instance.
 */
export function createTableStyleParser(theme, colorMap, logger) {
    return new TableStyleParser({ theme, colorMap, logger });
}
//# sourceMappingURL=TableStyleParser.js.map