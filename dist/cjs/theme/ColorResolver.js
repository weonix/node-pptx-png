"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ColorResolver = void 0;
exports.parseColorMap = parseColorMap;
exports.resolveEffectiveColorMap = resolveEffectiveColorMap;
const theme_js_1 = require("../types/theme.js");
const geometry_js_1 = require("../types/geometry.js");
const PptxParser_js_1 = require("../core/PptxParser.js");
/**
 * The set of valid theme color slot names a color map may reference.
 */
const THEME_COLOR_SLOTS = new Set([
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
]);
/**
 * Parses a color mapping element (p:clrMap or a:overrideClrMapping) into a
 * ColorMap. Missing or invalid attributes fall back to the default mapping.
 * @param clrMapNode The p:clrMap / a:overrideClrMapping node
 * @returns The parsed color map, or undefined when no node is given
 */
function parseColorMap(clrMapNode) {
    if (!clrMapNode)
        return undefined;
    const readSlot = (attr) => {
        const value = (0, PptxParser_js_1.getXmlAttr)(clrMapNode, attr);
        if (value !== undefined && THEME_COLOR_SLOTS.has(value)) {
            return value;
        }
        return theme_js_1.DEFAULT_COLOR_MAP[attr];
    };
    return {
        bg1: readSlot('bg1'),
        tx1: readSlot('tx1'),
        bg2: readSlot('bg2'),
        tx2: readSlot('tx2'),
        accent1: readSlot('accent1'),
        accent2: readSlot('accent2'),
        accent3: readSlot('accent3'),
        accent4: readSlot('accent4'),
        accent5: readSlot('accent5'),
        accent6: readSlot('accent6'),
        hlink: readSlot('hlink'),
        folHlink: readSlot('folHlink'),
    };
}
/**
 * Reads an a:overrideClrMapping from a node's p:clrMapOvr child, if present.
 * An a:masterClrMapping child (or a missing p:clrMapOvr) means "inherit".
 */
function parseColorMapOverride(node) {
    const clrMapOvr = node ? (0, PptxParser_js_1.getXmlChild)(node, 'p:clrMapOvr') : undefined;
    const override = clrMapOvr ? (0, PptxParser_js_1.getXmlChild)(clrMapOvr, 'a:overrideClrMapping') : undefined;
    return parseColorMap(override);
}
/**
 * Computes the effective color map for a slide per ECMA-376: the slide's
 * a:overrideClrMapping wins, then the layout's, then the master's p:clrMap,
 * then the default Office mapping.
 * @param masterNode The p:sldMaster element (carries p:clrMap)
 * @param layoutNode The p:sldLayout element (may carry p:clrMapOvr)
 * @param slideNode The p:sld element (may carry p:clrMapOvr)
 * @returns The effective color map (never undefined)
 */
function resolveEffectiveColorMap(masterNode, layoutNode, slideNode) {
    const slideOverride = parseColorMapOverride(slideNode);
    if (slideOverride)
        return slideOverride;
    const layoutOverride = parseColorMapOverride(layoutNode);
    if (layoutOverride)
        return layoutOverride;
    const masterMap = parseColorMap(masterNode ? (0, PptxParser_js_1.getXmlChild)(masterNode, 'p:clrMap') : undefined);
    return masterMap ?? { ...theme_js_1.DEFAULT_COLOR_MAP };
}
/**
 * Resolves colors from OpenXML color definitions.
 */
class ColorResolver {
    colorScheme;
    colorMap;
    constructor(colorScheme = theme_js_1.DEFAULT_OFFICE_COLORS, colorMap = theme_js_1.DEFAULT_COLOR_MAP) {
        this.colorScheme = colorScheme;
        this.colorMap = colorMap;
    }
    /**
     * Resolves a scheme color reference to an RGBA color.
     * Mappable names (bg1/tx1/bg2/tx2, accents, hyperlinks) resolve through
     * the color map; direct slot names (dk1/lt1/dk2/lt2) bypass it.
     * @param schemeColorType The scheme color name from a:schemeClr val
     * @param phClr Substitution color for val="phClr" (style placeholder)
     */
    resolveSchemeColor(schemeColorType, phClr) {
        switch (schemeColorType) {
            case 'dk1':
                return { ...this.colorScheme.dark1 };
            case 'lt1':
                return { ...this.colorScheme.light1 };
            case 'dk2':
                return { ...this.colorScheme.dark2 };
            case 'lt2':
                return { ...this.colorScheme.light2 };
            case 'tx1':
            case 'bg1':
            case 'tx2':
            case 'bg2':
            case 'accent1':
            case 'accent2':
            case 'accent3':
            case 'accent4':
            case 'accent5':
            case 'accent6':
            case 'hlink':
            case 'folHlink':
                return this.resolveColorSlot(this.colorMap[schemeColorType]);
            case 'phClr':
                // Placeholder color - substituted by the referencing style when known
                return phClr ? { ...phClr } : { ...geometry_js_1.Colors.black };
            default:
                return { ...geometry_js_1.Colors.black };
        }
    }
    /**
     * Resolves a concrete theme color slot to its scheme color.
     */
    resolveColorSlot(slot) {
        switch (slot) {
            case 'dk1':
                return { ...this.colorScheme.dark1 };
            case 'lt1':
                return { ...this.colorScheme.light1 };
            case 'dk2':
                return { ...this.colorScheme.dark2 };
            case 'lt2':
                return { ...this.colorScheme.light2 };
            case 'accent1':
                return { ...this.colorScheme.accent1 };
            case 'accent2':
                return { ...this.colorScheme.accent2 };
            case 'accent3':
                return { ...this.colorScheme.accent3 };
            case 'accent4':
                return { ...this.colorScheme.accent4 };
            case 'accent5':
                return { ...this.colorScheme.accent5 };
            case 'accent6':
                return { ...this.colorScheme.accent6 };
            case 'hlink':
                return { ...this.colorScheme.hyperlink };
            case 'folHlink':
                return { ...this.colorScheme.followedHyperlink };
            default:
                return { ...geometry_js_1.Colors.black };
        }
    }
    /**
     * Parses a hex color string to RGBA.
     */
    parseHexColor(hex) {
        // Remove # if present and handle 3-char shorthand
        hex = hex.replace('#', '');
        if (hex.length === 3) {
            const c0 = hex[0] ?? '0';
            const c1 = hex[1] ?? '0';
            const c2 = hex[2] ?? '0';
            hex = c0 + c0 + c1 + c1 + c2 + c2;
        }
        if (hex.length === 6) {
            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);
            return { r, g, b, a: 255 };
        }
        if (hex.length === 8) {
            const r = parseInt(hex.substring(0, 2), 16);
            const g = parseInt(hex.substring(2, 4), 16);
            const b = parseInt(hex.substring(4, 6), 16);
            const a = parseInt(hex.substring(6, 8), 16);
            return { r, g, b, a };
        }
        return { ...geometry_js_1.Colors.black };
    }
    /**
     * Resolves a color from an OpenXML color element.
     * @param node Parent node containing a color child (a:schemeClr, a:srgbClr, ...)
     * @param phClr Substitution color for a:schemeClr val="phClr" placeholders
     */
    resolveColorElement(node, phClr) {
        if (!node)
            return undefined;
        // Check for scheme color
        const schemeColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:schemeClr');
        if (schemeColor) {
            const val = (0, PptxParser_js_1.getXmlAttr)(schemeColor, 'val');
            if (val) {
                const baseColor = this.resolveSchemeColor(val, phClr);
                return this.applyTransforms(baseColor, schemeColor);
            }
        }
        // Check for sRGB color (hex)
        const srgbColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:srgbClr');
        if (srgbColor) {
            const val = (0, PptxParser_js_1.getXmlAttr)(srgbColor, 'val');
            if (val) {
                const baseColor = this.parseHexColor(val);
                return this.applyTransforms(baseColor, srgbColor);
            }
        }
        // Check for RGB percentage color
        const scrgbColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:scrgbClr');
        if (scrgbColor) {
            const r = (parseFloat((0, PptxParser_js_1.getXmlAttr)(scrgbColor, 'r') ?? '0') / 100000) * 255;
            const g = (parseFloat((0, PptxParser_js_1.getXmlAttr)(scrgbColor, 'g') ?? '0') / 100000) * 255;
            const b = (parseFloat((0, PptxParser_js_1.getXmlAttr)(scrgbColor, 'b') ?? '0') / 100000) * 255;
            const baseColor = { r: Math.round(r), g: Math.round(g), b: Math.round(b), a: 255 };
            return this.applyTransforms(baseColor, scrgbColor);
        }
        // Check for HSL color
        const hslColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:hslClr');
        if (hslColor) {
            const h = parseFloat((0, PptxParser_js_1.getXmlAttr)(hslColor, 'hue') ?? '0') / 60000;
            const s = parseFloat((0, PptxParser_js_1.getXmlAttr)(hslColor, 'sat') ?? '0') / 100000;
            const l = parseFloat((0, PptxParser_js_1.getXmlAttr)(hslColor, 'lum') ?? '0') / 100000;
            const baseColor = this.hslToRgba(h, s, l);
            return this.applyTransforms(baseColor, hslColor);
        }
        // Check for preset color
        const prstColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:prstClr');
        if (prstColor) {
            const val = (0, PptxParser_js_1.getXmlAttr)(prstColor, 'val');
            if (val) {
                const baseColor = this.resolvePresetColor(val);
                return this.applyTransforms(baseColor, prstColor);
            }
        }
        // Check for system color
        const sysColor = (0, PptxParser_js_1.getXmlChild)(node, 'a:sysClr');
        if (sysColor) {
            const lastClr = (0, PptxParser_js_1.getXmlAttr)(sysColor, 'lastClr');
            if (lastClr) {
                const baseColor = this.parseHexColor(lastClr);
                return this.applyTransforms(baseColor, sysColor);
            }
            const val = (0, PptxParser_js_1.getXmlAttr)(sysColor, 'val');
            if (val) {
                const baseColor = this.resolveSystemColor(val);
                return this.applyTransforms(baseColor, sysColor);
            }
        }
        return undefined;
    }
    /**
     * Extracts color transforms from an XML node.
     */
    extractTransforms(node) {
        const transforms = {};
        const tint = (0, PptxParser_js_1.getXmlChild)(node, 'a:tint');
        if (tint) {
            transforms.tint = parseInt((0, PptxParser_js_1.getXmlAttr)(tint, 'val') ?? '0', 10);
        }
        const shade = (0, PptxParser_js_1.getXmlChild)(node, 'a:shade');
        if (shade) {
            transforms.shade = parseInt((0, PptxParser_js_1.getXmlAttr)(shade, 'val') ?? '0', 10);
        }
        const satMod = (0, PptxParser_js_1.getXmlChild)(node, 'a:satMod');
        if (satMod) {
            transforms.satMod = parseInt((0, PptxParser_js_1.getXmlAttr)(satMod, 'val') ?? '100000', 10);
        }
        const lumMod = (0, PptxParser_js_1.getXmlChild)(node, 'a:lumMod');
        if (lumMod) {
            transforms.lumMod = parseInt((0, PptxParser_js_1.getXmlAttr)(lumMod, 'val') ?? '100000', 10);
        }
        const lumOff = (0, PptxParser_js_1.getXmlChild)(node, 'a:lumOff');
        if (lumOff) {
            transforms.lumOff = parseInt((0, PptxParser_js_1.getXmlAttr)(lumOff, 'val') ?? '0', 10);
        }
        const hueMod = (0, PptxParser_js_1.getXmlChild)(node, 'a:hueMod');
        if (hueMod) {
            transforms.hueMod = parseInt((0, PptxParser_js_1.getXmlAttr)(hueMod, 'val') ?? '100000', 10);
        }
        const hueOff = (0, PptxParser_js_1.getXmlChild)(node, 'a:hueOff');
        if (hueOff) {
            transforms.hueOff = parseInt((0, PptxParser_js_1.getXmlAttr)(hueOff, 'val') ?? '0', 10);
        }
        const alpha = (0, PptxParser_js_1.getXmlChild)(node, 'a:alpha');
        if (alpha) {
            transforms.alpha = parseInt((0, PptxParser_js_1.getXmlAttr)(alpha, 'val') ?? '100000', 10);
        }
        return transforms;
    }
    /**
     * Applies color transforms to a base color.
     */
    applyTransforms(baseColor, node) {
        const transforms = this.extractTransforms(node);
        return this.applyColorTransforms(baseColor, transforms);
    }
    /**
     * Resolves a theme style color (from a fmtScheme style list) to RGBA.
     * Concrete colors are returned as-is; phClr placeholders are substituted
     * with the supplied color and their stored transforms applied.
     * @param styleColor The theme style color
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved color, or undefined for a phClr placeholder with
     *          no substitution color available
     */
    resolveThemeStyleColor(styleColor, phClr) {
        if (styleColor.isPhClr) {
            if (!phClr)
                return undefined;
            return styleColor.transforms
                ? this.applyColorTransforms(phClr, styleColor.transforms)
                : { ...phClr };
        }
        return styleColor.color ? { ...styleColor.color } : undefined;
    }
    /**
     * Applies color transforms to a base color.
     */
    applyColorTransforms(baseColor, transforms) {
        // Apply transforms in order: tint/shade -> satMod -> lumMod/lumOff -> hueMod/hueOff -> alpha
        // Tint and shade are per-channel operations in linear-light sRGB
        // (PowerPoint semantics; matches PowerPoint's own exports and
        // LibreOffice's oox implementation): shade multiplies each linearized
        // channel by val, tint interpolates each linearized channel toward
        // white by (1 - val). Doing these in HSL instead shifts colors
        // noticeably (e.g. theme lnRef strokes darken too much).
        let working = baseColor;
        if (transforms.tint !== undefined || transforms.shade !== undefined) {
            const tint = transforms.tint !== undefined
                ? Math.min(1, Math.max(0, transforms.tint / 100000))
                : undefined;
            const shade = transforms.shade !== undefined
                ? Math.min(1, Math.max(0, transforms.shade / 100000))
                : undefined;
            const applyChannel = (c) => {
                let lin = ColorResolver.srgbToLinear(c);
                if (tint !== undefined)
                    lin = lin * tint + (1 - tint);
                if (shade !== undefined)
                    lin = lin * shade;
                return ColorResolver.linearToSrgb(lin);
            };
            working = {
                r: applyChannel(baseColor.r),
                g: applyChannel(baseColor.g),
                b: applyChannel(baseColor.b),
                a: baseColor.a,
            };
        }
        // Convert to HSL for the remaining transformations
        let { h, s, l } = this.rgbaToHsl(working);
        let alpha = working.a / 255;
        // Saturation modulation
        if (transforms.satMod !== undefined) {
            s = Math.min(1, Math.max(0, s * (transforms.satMod / 100000)));
        }
        // Luminance modulation
        if (transforms.lumMod !== undefined) {
            l = Math.min(1, Math.max(0, l * (transforms.lumMod / 100000)));
        }
        // Luminance offset
        if (transforms.lumOff !== undefined) {
            l = Math.min(1, Math.max(0, l + transforms.lumOff / 100000));
        }
        // Hue modulation
        if (transforms.hueMod !== undefined) {
            h = (h * (transforms.hueMod / 100000)) % 360;
        }
        // Hue offset
        if (transforms.hueOff !== undefined) {
            h = (h + transforms.hueOff / 60000) % 360;
            if (h < 0)
                h += 360;
        }
        // Alpha
        if (transforms.alpha !== undefined) {
            alpha = transforms.alpha / 100000;
        }
        // Convert back to RGBA
        const result = this.hslToRgba(h, s, l);
        result.a = Math.round(alpha * 255);
        return result;
    }
    /**
     * Converts an sRGB channel value (0-255) to linear-light (0-1) using the
     * IEC 61966-2-1 transfer function. Tint/shade transforms operate in this
     * space (ECMA-376 §20.1.2.3.32/§20.1.2.3.31 as implemented by PowerPoint).
     */
    static srgbToLinear(channel) {
        const v = channel / 255;
        return v <= 0.04045 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    }
    /**
     * Converts a linear-light value (0-1) back to an sRGB channel (0-255),
     * clamping out-of-range inputs.
     */
    static linearToSrgb(linear) {
        const v = Math.min(1, Math.max(0, linear));
        const c = v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
        return Math.round(c * 255);
    }
    /**
     * Converts RGBA to HSL.
     */
    rgbaToHsl(color) {
        const r = color.r / 255;
        const g = color.g / 255;
        const b = color.b / 255;
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        const delta = max - min;
        let h = 0;
        let s = 0;
        const l = (max + min) / 2;
        if (delta !== 0) {
            s = l > 0.5 ? delta / (2 - max - min) : delta / (max + min);
            if (max === r) {
                h = ((g - b) / delta + (g < b ? 6 : 0)) * 60;
            }
            else if (max === g) {
                h = ((b - r) / delta + 2) * 60;
            }
            else {
                h = ((r - g) / delta + 4) * 60;
            }
        }
        return { h, s, l };
    }
    /**
     * Converts HSL to RGBA.
     */
    hslToRgba(h, s, l) {
        let r, g, b;
        if (s === 0) {
            r = g = b = l;
        }
        else {
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = this.hueToRgb(p, q, h / 360 + 1 / 3);
            g = this.hueToRgb(p, q, h / 360);
            b = this.hueToRgb(p, q, h / 360 - 1 / 3);
        }
        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255),
            a: 255,
        };
    }
    hueToRgb(p, q, t) {
        if (t < 0)
            t += 1;
        if (t > 1)
            t -= 1;
        if (t < 1 / 6)
            return p + (q - p) * 6 * t;
        if (t < 1 / 2)
            return q;
        if (t < 2 / 3)
            return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    }
    /**
     * Resolves a preset color name to RGBA.
     */
    resolvePresetColor(name) {
        const presetColors = {
            black: { r: 0, g: 0, b: 0, a: 255 },
            white: { r: 255, g: 255, b: 255, a: 255 },
            red: { r: 255, g: 0, b: 0, a: 255 },
            green: { r: 0, g: 128, b: 0, a: 255 },
            blue: { r: 0, g: 0, b: 255, a: 255 },
            yellow: { r: 255, g: 255, b: 0, a: 255 },
            cyan: { r: 0, g: 255, b: 255, a: 255 },
            magenta: { r: 255, g: 0, b: 255, a: 255 },
            gray: { r: 128, g: 128, b: 128, a: 255 },
            ltGray: { r: 211, g: 211, b: 211, a: 255 },
            dkGray: { r: 169, g: 169, b: 169, a: 255 },
            dkRed: { r: 139, g: 0, b: 0, a: 255 },
            dkGreen: { r: 0, g: 100, b: 0, a: 255 },
            dkBlue: { r: 0, g: 0, b: 139, a: 255 },
            orange: { r: 255, g: 165, b: 0, a: 255 },
            pink: { r: 255, g: 192, b: 203, a: 255 },
            purple: { r: 128, g: 0, b: 128, a: 255 },
            brown: { r: 165, g: 42, b: 42, a: 255 },
            navy: { r: 0, g: 0, b: 128, a: 255 },
            teal: { r: 0, g: 128, b: 128, a: 255 },
            olive: { r: 128, g: 128, b: 0, a: 255 },
            silver: { r: 192, g: 192, b: 192, a: 255 },
            maroon: { r: 128, g: 0, b: 0, a: 255 },
            aqua: { r: 0, g: 255, b: 255, a: 255 },
            lime: { r: 0, g: 255, b: 0, a: 255 },
            fuchsia: { r: 255, g: 0, b: 255, a: 255 },
        };
        return presetColors[name] ?? { ...geometry_js_1.Colors.black };
    }
    /**
     * Resolves a system color name to RGBA.
     */
    resolveSystemColor(name) {
        const systemColors = {
            windowText: { r: 0, g: 0, b: 0, a: 255 },
            window: { r: 255, g: 255, b: 255, a: 255 },
            highlightText: { r: 255, g: 255, b: 255, a: 255 },
            highlight: { r: 0, g: 120, b: 215, a: 255 },
            grayText: { r: 128, g: 128, b: 128, a: 255 },
            btnFace: { r: 240, g: 240, b: 240, a: 255 },
            btnText: { r: 0, g: 0, b: 0, a: 255 },
            captionText: { r: 0, g: 0, b: 0, a: 255 },
            menuText: { r: 0, g: 0, b: 0, a: 255 },
            scrollBar: { r: 200, g: 200, b: 200, a: 255 },
        };
        return systemColors[name] ?? { ...geometry_js_1.Colors.black };
    }
    /**
     * Converts RGBA to CSS color string.
     */
    rgbaToCss(color) {
        if (color.a === 255) {
            return `rgb(${color.r}, ${color.g}, ${color.b})`;
        }
        return `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a / 255).toFixed(3)})`;
    }
    /**
     * Converts RGBA to hex string.
     */
    rgbaToHex(color, includeAlpha = false) {
        const r = color.r.toString(16).padStart(2, '0');
        const g = color.g.toString(16).padStart(2, '0');
        const b = color.b.toString(16).padStart(2, '0');
        if (includeAlpha) {
            const a = color.a.toString(16).padStart(2, '0');
            return `#${r}${g}${b}${a}`;
        }
        return `#${r}${g}${b}`;
    }
    /**
     * Calculates relative luminance for contrast calculations.
     */
    calculateLuminance(color) {
        const r = color.r / 255;
        const g = color.g / 255;
        const b = color.b / 255;
        const rs = r <= 0.03928 ? r / 12.92 : Math.pow((r + 0.055) / 1.055, 2.4);
        const gs = g <= 0.03928 ? g / 12.92 : Math.pow((g + 0.055) / 1.055, 2.4);
        const bs = b <= 0.03928 ? b / 12.92 : Math.pow((b + 0.055) / 1.055, 2.4);
        return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
    }
    /**
     * Determines if a color is dark (for contrast purposes).
     */
    isDarkColor(color) {
        return this.calculateLuminance(color) < 0.5;
    }
}
exports.ColorResolver = ColorResolver;
