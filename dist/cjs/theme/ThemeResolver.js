"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ThemeResolver = void 0;
const index_js_1 = require("../types/index.js");
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const ColorResolver_js_1 = require("./ColorResolver.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Fill element tag names that may appear in a fmtScheme fill style list,
 * in the order they are checked.
 */
const FILL_LIST_TAGS = [
    'a:noFill',
    'a:solidFill',
    'a:gradFill',
    'a:blipFill',
    'a:pattFill',
    'a:grpFill',
];
/**
 * Resolves theme colors, fonts, and effects from PPTX theme data.
 */
class ThemeResolver {
    logger;
    colorResolver;
    constructor(logger) {
        this.logger = logger ?? (0, Logger_js_1.createLogger)('warn', 'ThemeResolver');
        this.colorResolver = new ColorResolver_js_1.ColorResolver();
    }
    /**
     * Resolves the complete theme from PPTX parser.
     */
    async resolveTheme(parser) {
        try {
            const themeData = await parser.getTheme();
            return this.parseTheme(themeData);
        }
        catch (error) {
            this.logger.warn('Failed to load theme, using default', {
                error: error instanceof Error ? error.message : String(error),
            });
            return { ...index_js_1.DEFAULT_THEME };
        }
    }
    /**
     * Parses theme data into resolved theme.
     */
    parseTheme(themeData) {
        const themeElements = (0, PptxParser_js_1.getXmlChild)(themeData.content, 'a:themeElements');
        if (!themeElements) {
            this.logger.debug('No theme elements found, using default theme');
            return { ...index_js_1.DEFAULT_THEME };
        }
        const colorScheme = this.resolveColorScheme((0, PptxParser_js_1.getXmlChild)(themeElements, 'a:clrScheme'));
        const fontScheme = this.resolveFontScheme((0, PptxParser_js_1.getXmlChild)(themeElements, 'a:fontScheme'));
        const fmtScheme = (0, PptxParser_js_1.getXmlChild)(themeElements, 'a:fmtScheme');
        const backgroundFillStyles = this.resolveBackgroundFillStyles(fmtScheme, colorScheme);
        const formatScheme = this.parseFormatScheme(fmtScheme, colorScheme);
        // Update the color resolver with the new color scheme
        this.colorResolver = new ColorResolver_js_1.ColorResolver(colorScheme);
        return {
            colors: colorScheme,
            fonts: fontScheme,
            backgroundFillStyles,
            formatScheme,
        };
    }
    /**
     * Parses the a:fmtScheme style matrix into full style lists.
     * Style references (a:fillRef/a:lnRef/a:effectRef) index into these lists;
     * phClr placeholders are kept symbolic for substitution at use time.
     */
    parseFormatScheme(fmtScheme, colorScheme) {
        if (!fmtScheme)
            return undefined;
        const resolver = new ColorResolver_js_1.ColorResolver(colorScheme);
        const formatScheme = {
            fillStyles: this.parseFillStyleList((0, PptxParser_js_1.getXmlChild)(fmtScheme, 'a:fillStyleLst'), 'a:fillStyleLst', resolver),
            lineStyles: this.parseLineStyleList((0, PptxParser_js_1.getXmlChild)(fmtScheme, 'a:lnStyleLst'), resolver),
            effectStyles: this.parseEffectStyleList((0, PptxParser_js_1.getXmlChild)(fmtScheme, 'a:effectStyleLst'), resolver),
            backgroundFillStyles: this.parseFillStyleList((0, PptxParser_js_1.getXmlChild)(fmtScheme, 'a:bgFillStyleLst'), 'a:bgFillStyleLst', resolver),
        };
        this.logger.debug('Parsed format scheme', {
            fillStyles: formatScheme.fillStyles.length,
            lineStyles: formatScheme.lineStyles.length,
            effectStyles: formatScheme.effectStyles.length,
            backgroundFillStyles: formatScheme.backgroundFillStyles.length,
        });
        return formatScheme;
    }
    /**
     * Parses a fill style list (a:fillStyleLst / a:bgFillStyleLst) preserving
     * document order, which style reference indices depend on.
     */
    parseFillStyleList(listNode, wrapperTag, resolver) {
        if (!listNode)
            return [];
        const entries = (0, PptxParser_js_1.getChildrenInDocumentOrder)(listNode, wrapperTag, FILL_LIST_TAGS);
        return entries.map(({ tagName, node }) => this.parseThemeFill(tagName, node, resolver));
    }
    /**
     * Parses a single fill element from a style list into a ThemeFill.
     * Unsupported variants (blip/group fills) degrade to a solid phClr fill so
     * list positions stay aligned with style reference indices.
     */
    parseThemeFill(tagName, node, resolver) {
        switch (tagName) {
            case 'a:noFill':
                return { type: 'none' };
            case 'a:solidFill':
                return { type: 'solid', color: this.parseThemeStyleColor(node, resolver) };
            case 'a:gradFill':
                return this.parseThemeGradientFill(node, resolver);
            case 'a:pattFill': {
                const fgClr = (0, PptxParser_js_1.getXmlChild)(node, 'a:fgClr');
                const bgClr = (0, PptxParser_js_1.getXmlChild)(node, 'a:bgClr');
                return {
                    type: 'pattern',
                    preset: (0, PptxParser_js_1.getXmlAttr)(node, 'prst') ?? 'solid',
                    foregroundColor: fgClr
                        ? this.parseThemeStyleColor(fgClr, resolver)
                        : { color: { r: 0, g: 0, b: 0, a: 255 } },
                    backgroundColor: bgClr
                        ? this.parseThemeStyleColor(bgClr, resolver)
                        : { color: { r: 255, g: 255, b: 255, a: 255 } },
                };
            }
            default:
                // a:blipFill / a:grpFill: unsupported in the style matrix; degrade to
                // a plain phClr solid so the reference still renders a sensible color
                this.logger.debug('Unsupported theme fill variant, degrading to solid phClr', {
                    tag: tagName,
                });
                return { type: 'solid', color: { isPhClr: true } };
        }
    }
    /**
     * Parses a gradient fill from a style list, keeping phClr stops symbolic.
     */
    parseThemeGradientFill(gradFill, resolver) {
        const stops = [];
        const gsLst = (0, PptxParser_js_1.getXmlChild)(gradFill, 'a:gsLst');
        if (gsLst) {
            for (const gs of (0, PptxParser_js_1.getXmlChildren)(gsLst, 'a:gs')) {
                const pos = (0, PptxParser_js_1.getXmlAttr)(gs, 'pos');
                const position = pos !== undefined ? (0, UnitConverter_js_1.percentageToDecimal)(parseInt(pos, 10)) : 0;
                stops.push({ position, color: this.parseThemeStyleColor(gs, resolver) });
            }
        }
        stops.sort((a, b) => a.position - b.position);
        const lin = (0, PptxParser_js_1.getXmlChild)(gradFill, 'a:lin');
        if (lin) {
            const ang = (0, PptxParser_js_1.getXmlAttr)(lin, 'ang');
            const angle = ang !== undefined ? parseInt(ang, 10) / 60000 : 0;
            return { type: 'gradient', stops, angle, isRadial: false };
        }
        const path = (0, PptxParser_js_1.getXmlChild)(gradFill, 'a:path');
        if (path) {
            let centerX = 0.5;
            let centerY = 0.5;
            const fillToRect = (0, PptxParser_js_1.getXmlChild)(path, 'a:fillToRect');
            if (fillToRect) {
                const left = (0, UnitConverter_js_1.percentageToDecimal)(parseInt((0, PptxParser_js_1.getXmlAttr)(fillToRect, 'l') ?? '0', 10));
                const top = (0, UnitConverter_js_1.percentageToDecimal)(parseInt((0, PptxParser_js_1.getXmlAttr)(fillToRect, 't') ?? '0', 10));
                const right = (0, UnitConverter_js_1.percentageToDecimal)(parseInt((0, PptxParser_js_1.getXmlAttr)(fillToRect, 'r') ?? '100000', 10));
                const bottom = (0, UnitConverter_js_1.percentageToDecimal)(parseInt((0, PptxParser_js_1.getXmlAttr)(fillToRect, 'b') ?? '100000', 10));
                centerX = (left + right) / 2;
                centerY = (top + bottom) / 2;
            }
            return { type: 'gradient', stops, isRadial: true, centerX, centerY };
        }
        return { type: 'gradient', stops, angle: 0, isRadial: false };
    }
    /**
     * Parses a color child of a style list element into a ThemeStyleColor.
     * a:schemeClr val="phClr" is kept symbolic with its transforms; all other
     * colors are resolved to concrete RGBA at parse time.
     */
    parseThemeStyleColor(container, resolver) {
        const schemeClr = (0, PptxParser_js_1.getXmlChild)(container, 'a:schemeClr');
        if (schemeClr && (0, PptxParser_js_1.getXmlAttr)(schemeClr, 'val') === 'phClr') {
            const transforms = resolver.extractTransforms(schemeClr);
            return Object.keys(transforms).length > 0 ? { isPhClr: true, transforms } : { isPhClr: true };
        }
        const color = resolver.resolveColorElement(container);
        if (color) {
            return { color };
        }
        // No recognizable color: treat as a bare placeholder
        return { isPhClr: true };
    }
    /**
     * Parses the line style list (a:lnStyleLst).
     */
    parseLineStyleList(listNode, resolver) {
        if (!listNode)
            return [];
        return (0, PptxParser_js_1.getXmlChildren)(listNode, 'a:ln').map((ln) => {
            const style = {};
            const w = (0, PptxParser_js_1.getXmlAttr)(ln, 'w');
            if (w !== undefined)
                style.width = parseInt(w, 10);
            const cap = (0, PptxParser_js_1.getXmlAttr)(ln, 'cap');
            if (cap !== undefined)
                style.cap = cap;
            const cmpd = (0, PptxParser_js_1.getXmlAttr)(ln, 'cmpd');
            if (cmpd !== undefined)
                style.compound = cmpd;
            const algn = (0, PptxParser_js_1.getXmlAttr)(ln, 'algn');
            if (algn !== undefined)
                style.align = algn;
            const prstDash = (0, PptxParser_js_1.getXmlChild)(ln, 'a:prstDash');
            const dashVal = prstDash ? (0, PptxParser_js_1.getXmlAttr)(prstDash, 'val') : undefined;
            if (dashVal !== undefined)
                style.dash = dashVal;
            if ((0, PptxParser_js_1.getXmlChild)(ln, 'a:round'))
                style.join = 'round';
            else if ((0, PptxParser_js_1.getXmlChild)(ln, 'a:bevel'))
                style.join = 'bevel';
            else if ((0, PptxParser_js_1.getXmlChild)(ln, 'a:miter'))
                style.join = 'miter';
            for (const tag of FILL_LIST_TAGS) {
                const fillNode = (0, PptxParser_js_1.getXmlChild)(ln, tag);
                if (fillNode) {
                    style.fill = this.parseThemeFill(tag, fillNode, resolver);
                    break;
                }
            }
            return style;
        });
    }
    /**
     * Parses the effect style list (a:effectStyleLst). Effects are parsed and
     * stored for style references; rendering them is out of scope.
     */
    parseEffectStyleList(listNode, resolver) {
        if (!listNode)
            return [];
        return (0, PptxParser_js_1.getXmlChildren)(listNode, 'a:effectStyle').map((effectStyle) => {
            const parsed = {};
            const effectLst = (0, PptxParser_js_1.getXmlChild)(effectStyle, 'a:effectLst');
            const outerShdw = effectLst ? (0, PptxParser_js_1.getXmlChild)(effectLst, 'a:outerShdw') : undefined;
            if (outerShdw) {
                const blurRad = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'blurRad');
                const dist = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'dist');
                const dir = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'dir');
                const algn = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'algn');
                const rotWithShape = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'rotWithShape');
                parsed.outerShadow = {
                    ...(blurRad !== undefined && { blurRadius: parseInt(blurRad, 10) }),
                    ...(dist !== undefined && { distance: parseInt(dist, 10) }),
                    ...(dir !== undefined && { direction: parseInt(dir, 10) / 60000 }),
                    ...(algn !== undefined && { alignment: algn }),
                    ...(rotWithShape !== undefined && { rotateWithShape: rotWithShape !== '0' }),
                    color: this.parseThemeStyleColor(outerShdw, resolver),
                };
            }
            return parsed;
        });
    }
    /**
     * Resolves the color scheme from theme XML.
     */
    resolveColorScheme(clrScheme) {
        if (!clrScheme) {
            this.logger.debug('No color scheme found, using default');
            return { ...index_js_1.DEFAULT_OFFICE_COLORS };
        }
        return {
            dark1: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:dk1')) ?? index_js_1.DEFAULT_OFFICE_COLORS.dark1,
            light1: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:lt1')) ?? index_js_1.DEFAULT_OFFICE_COLORS.light1,
            dark2: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:dk2')) ?? index_js_1.DEFAULT_OFFICE_COLORS.dark2,
            light2: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:lt2')) ?? index_js_1.DEFAULT_OFFICE_COLORS.light2,
            accent1: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent1')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent1,
            accent2: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent2')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent2,
            accent3: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent3')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent3,
            accent4: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent4')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent4,
            accent5: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent5')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent5,
            accent6: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:accent6')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.accent6,
            hyperlink: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:hlink')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.hyperlink,
            followedHyperlink: this.extractThemeColor((0, PptxParser_js_1.getXmlChild)(clrScheme, 'a:folHlink')) ??
                index_js_1.DEFAULT_OFFICE_COLORS.followedHyperlink,
        };
    }
    /**
     * Extracts a color from a theme color element.
     */
    extractThemeColor(node) {
        if (!node)
            return undefined;
        // Check for sRGB color (most common)
        const srgbClr = (0, PptxParser_js_1.getXmlChild)(node, 'a:srgbClr');
        if (srgbClr) {
            const val = (0, PptxParser_js_1.getXmlAttr)(srgbClr, 'val');
            if (val) {
                return this.colorResolver.parseHexColor(val);
            }
        }
        // Check for system color
        const sysClr = (0, PptxParser_js_1.getXmlChild)(node, 'a:sysClr');
        if (sysClr) {
            const lastClr = (0, PptxParser_js_1.getXmlAttr)(sysClr, 'lastClr');
            if (lastClr) {
                return this.colorResolver.parseHexColor(lastClr);
            }
            const val = (0, PptxParser_js_1.getXmlAttr)(sysClr, 'val');
            if (val) {
                return this.colorResolver.resolveSystemColor(val);
            }
        }
        return undefined;
    }
    /**
     * Resolves the font scheme from theme XML.
     */
    resolveFontScheme(fontScheme) {
        if (!fontScheme) {
            this.logger.debug('No font scheme found, using default');
            return { ...index_js_1.DEFAULT_FONT_SCHEME };
        }
        const majorFont = (0, PptxParser_js_1.getXmlChild)(fontScheme, 'a:majorFont');
        const minorFont = (0, PptxParser_js_1.getXmlChild)(fontScheme, 'a:minorFont');
        return {
            majorFont: this.extractFontFamily(majorFont, 'latin') ?? index_js_1.DEFAULT_FONT_SCHEME.majorFont,
            minorFont: this.extractFontFamily(minorFont, 'latin') ?? index_js_1.DEFAULT_FONT_SCHEME.minorFont,
            majorFontEastAsian: this.extractFontFamily(majorFont, 'ea'),
            minorFontEastAsian: this.extractFontFamily(minorFont, 'ea'),
            majorFontComplexScript: this.extractFontFamily(majorFont, 'cs'),
            minorFontComplexScript: this.extractFontFamily(minorFont, 'cs'),
        };
    }
    /**
     * Extracts a font family from a font element.
     */
    extractFontFamily(fontNode, type) {
        if (!fontNode)
            return undefined;
        const elementName = type === 'latin' ? 'a:latin' : type === 'ea' ? 'a:ea' : 'a:cs';
        const fontElement = (0, PptxParser_js_1.getXmlChild)(fontNode, elementName);
        if (fontElement) {
            return (0, PptxParser_js_1.getXmlAttr)(fontElement, 'typeface');
        }
        return undefined;
    }
    /**
     * Resolves background fill styles from the format scheme.
     */
    resolveBackgroundFillStyles(fmtScheme, colorScheme) {
        if (!fmtScheme)
            return undefined;
        const bgFillStyleLst = (0, PptxParser_js_1.getXmlChild)(fmtScheme, 'a:bgFillStyleLst');
        if (!bgFillStyleLst)
            return undefined;
        const styles = [];
        // Create a color resolver with the current color scheme for fill resolution
        const resolver = new ColorResolver_js_1.ColorResolver(colorScheme);
        // Look for solid fills
        const solidFills = (0, PptxParser_js_1.getXmlChildren)(bgFillStyleLst, 'a:solidFill');
        for (const fill of solidFills) {
            const color = resolver.resolveColorElement(fill);
            if (color) {
                styles.push(color);
            }
        }
        // Look for gradient fills (use first stop color as representative)
        const gradFills = (0, PptxParser_js_1.getXmlChildren)(bgFillStyleLst, 'a:gradFill');
        for (const fill of gradFills) {
            const gsLst = (0, PptxParser_js_1.getXmlChild)(fill, 'a:gsLst');
            if (gsLst) {
                const stops = (0, PptxParser_js_1.getXmlChildren)(gsLst, 'a:gs');
                if (stops.length > 0 && stops[0]) {
                    const color = resolver.resolveColorElement(stops[0]);
                    if (color) {
                        styles.push(color);
                    }
                }
            }
        }
        this.logger.debug('Extracted background fill styles', { count: styles.length });
        return styles.length > 0 ? styles : undefined;
    }
    /**
     * Gets the color resolver for this theme.
     */
    getColorResolver() {
        return this.colorResolver;
    }
    /**
     * Creates a color resolver with the given color scheme.
     */
    createColorResolver(colorScheme) {
        return new ColorResolver_js_1.ColorResolver(colorScheme);
    }
}
exports.ThemeResolver = ThemeResolver;
