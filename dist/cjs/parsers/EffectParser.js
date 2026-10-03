"use strict";
/**
 * Parses DrawingML effect lists (a:effectLst) from shape properties.
 * Plain outer shadows are rendered by ShapeRenderer via canvas shadow
 * state; perspective/picture-fill shadows, inner shadow, glow, soft edge,
 * and reflection are rendered by EffectRenderer from these records.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.EffectParser = void 0;
exports.computeShadowOffset = computeShadowOffset;
exports.createEffectParser = createEffectParser;
const PptxParser_js_1 = require("../core/PptxParser.js");
const ColorResolver_js_1 = require("../theme/ColorResolver.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Default shadow color when an effect carries no color child
 * (opaque black, matching the PowerPoint default).
 */
const DEFAULT_SHADOW_COLOR = { r: 0, g: 0, b: 0, a: 255 };
/**
 * Computes the x/y offset of a shadow from its distance and direction.
 * Direction is measured clockwise from the positive x-axis in y-down slide
 * space (matching a:lin gradient angles), so the offset maps directly to
 * (cos, sin) without adjustment.
 * @param distance Shadow distance (any length unit; the result is in the same unit)
 * @param directionDegrees Shadow direction in degrees
 * @returns Shadow offset in the same unit as the distance
 */
function computeShadowOffset(distance, directionDegrees) {
    const rad = (directionDegrees * Math.PI) / 180;
    return {
        dx: distance * Math.cos(rad),
        dy: distance * Math.sin(rad),
    };
}
/**
 * Reads an integer attribute with a fallback for missing/invalid values.
 */
function parseIntAttr(node, name, fallback) {
    const raw = (0, PptxParser_js_1.getXmlAttr)(node, name);
    if (raw === undefined)
        return fallback;
    const value = parseInt(raw, 10);
    return Number.isFinite(value) ? value : fallback;
}
/**
 * Reads an OOXML angle attribute (60,000ths of a degree) as degrees.
 */
function parseAngleAttr(node, name, fallbackDegrees) {
    return parseIntAttr(node, name, fallbackDegrees * 60000) / 60000;
}
/**
 * Reads an OOXML percentage attribute (100000 = 100%) as a decimal.
 */
function parsePercentAttr(node, name, fallbackDecimal) {
    return (0, UnitConverter_js_1.percentageToDecimal)(parseIntAttr(node, name, fallbackDecimal * 100000));
}
/**
 * Reads an OOXML boolean attribute ('0'/'false' are false).
 */
function parseBoolAttr(node, name, fallback) {
    const raw = (0, PptxParser_js_1.getXmlAttr)(node, name);
    if (raw === undefined)
        return fallback;
    return raw !== '0' && raw !== 'false';
}
/**
 * Gets a child element, preserving the presence of empty elements
 * (self-closed childless elements parse to empty strings, which must not
 * be confused with a missing element).
 */
function getPresentChild(node, name) {
    if (!node || typeof node !== 'object')
        return undefined;
    const child = node[name];
    if (child === undefined)
        return undefined;
    return typeof child === 'object' && child !== null ? child : {};
}
/**
 * Parses a:effectLst elements into typed effect records.
 */
class EffectParser {
    logger;
    colorResolver;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'EffectParser');
        this.colorResolver = new ColorResolver_js_1.ColorResolver(config.theme.colors, config.colorMap);
    }
    /**
     * Parses the a:effectLst of a shape properties node (p:spPr).
     * @param spPr Shape properties node
     * @returns The parsed effects, an empty record for an explicit empty
     *          a:effectLst, or undefined when the node has no a:effectLst
     *          (so style-referenced effects may still apply)
     */
    parseShapeEffects(spPr) {
        if (!spPr || !(0, PptxParser_js_1.hasXmlChild)(spPr, 'a:effectLst')) {
            return undefined;
        }
        return this.parseEffectList(getPresentChild(spPr, 'a:effectLst'));
    }
    /**
     * Parses an a:effectLst node into typed effect records.
     * @param effectLst The a:effectLst node (an empty node yields an empty record)
     */
    parseEffectList(effectLst) {
        const effects = {};
        if (!effectLst) {
            return effects;
        }
        const outerShdw = getPresentChild(effectLst, 'a:outerShdw');
        if (outerShdw) {
            effects.outerShadow = this.parseOuterShadow(outerShdw);
        }
        const innerShdw = getPresentChild(effectLst, 'a:innerShdw');
        if (innerShdw) {
            effects.innerShadow = this.parseInnerShadow(innerShdw);
        }
        const glow = getPresentChild(effectLst, 'a:glow');
        if (glow) {
            effects.glow = this.parseGlow(glow);
        }
        const softEdge = getPresentChild(effectLst, 'a:softEdge');
        if (softEdge) {
            effects.softEdge = this.parseSoftEdge(softEdge);
        }
        const reflection = getPresentChild(effectLst, 'a:reflection');
        if (reflection) {
            effects.reflection = this.parseReflection(reflection);
        }
        this.logger.debug('Parsed effect list', {
            outerShadow: !!effects.outerShadow,
            innerShadow: !!effects.innerShadow,
            glow: !!effects.glow,
            softEdge: !!effects.softEdge,
            reflection: !!effects.reflection,
        });
        return effects;
    }
    /**
     * Resolves a theme effect-style outer shadow (from a:effectStyleLst) into
     * a renderable effect, substituting phClr placeholders in the shadow color
     * with the color supplied by the style reference (a:effectRef child color).
     * @param shadow The theme outer shadow definition
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved outer shadow effect
     */
    resolveThemeOuterShadow(shadow, phClr) {
        const color = shadow.color
            ? this.colorResolver.resolveThemeStyleColor(shadow.color, phClr)
            : undefined;
        return {
            type: 'outerShadow',
            blurRadius: shadow.blurRadius ?? 0,
            distance: shadow.distance ?? 0,
            direction: shadow.direction ?? 0,
            scaleX: 1,
            scaleY: 1,
            skewX: 0,
            skewY: 0,
            rotateWithShape: shadow.rotateWithShape ?? true,
            ...(shadow.alignment !== undefined && { alignment: shadow.alignment }),
            color: color ?? { ...DEFAULT_SHADOW_COLOR },
        };
    }
    /**
     * Parses an a:outerShdw element (defaults per ECMA-376 §20.1.8.45).
     */
    parseOuterShadow(outerShdw) {
        const alignment = (0, PptxParser_js_1.getXmlAttr)(outerShdw, 'algn');
        return {
            type: 'outerShadow',
            blurRadius: parseIntAttr(outerShdw, 'blurRad', 0),
            distance: parseIntAttr(outerShdw, 'dist', 0),
            direction: parseAngleAttr(outerShdw, 'dir', 0),
            scaleX: parsePercentAttr(outerShdw, 'sx', 1),
            scaleY: parsePercentAttr(outerShdw, 'sy', 1),
            skewX: parseAngleAttr(outerShdw, 'kx', 0),
            skewY: parseAngleAttr(outerShdw, 'ky', 0),
            rotateWithShape: parseBoolAttr(outerShdw, 'rotWithShape', true),
            ...(alignment !== undefined && { alignment }),
            color: this.parseEffectColor(outerShdw),
        };
    }
    /**
     * Parses an a:innerShdw element (defaults per ECMA-376 §20.1.8.40).
     */
    parseInnerShadow(innerShdw) {
        return {
            type: 'innerShadow',
            blurRadius: parseIntAttr(innerShdw, 'blurRad', 0),
            distance: parseIntAttr(innerShdw, 'dist', 0),
            direction: parseAngleAttr(innerShdw, 'dir', 0),
            color: this.parseEffectColor(innerShdw),
        };
    }
    /**
     * Parses an a:glow element (defaults per ECMA-376 §20.1.8.32).
     */
    parseGlow(glow) {
        return {
            type: 'glow',
            radius: parseIntAttr(glow, 'rad', 0),
            color: this.parseEffectColor(glow),
        };
    }
    /**
     * Parses an a:softEdge element (ECMA-376 §20.1.8.53).
     */
    parseSoftEdge(softEdge) {
        return {
            type: 'softEdge',
            radius: parseIntAttr(softEdge, 'rad', 0),
        };
    }
    /**
     * Parses an a:reflection element (defaults per ECMA-376 §20.1.8.50).
     */
    parseReflection(reflection) {
        const alignment = (0, PptxParser_js_1.getXmlAttr)(reflection, 'algn');
        return {
            type: 'reflection',
            blurRadius: parseIntAttr(reflection, 'blurRad', 0),
            distance: parseIntAttr(reflection, 'dist', 0),
            direction: parseAngleAttr(reflection, 'dir', 0),
            fadeDirection: parseAngleAttr(reflection, 'fadeDir', 90),
            startAlpha: parsePercentAttr(reflection, 'stA', 1),
            endAlpha: parsePercentAttr(reflection, 'endA', 0),
            startPosition: parsePercentAttr(reflection, 'stPos', 0),
            endPosition: parsePercentAttr(reflection, 'endPos', 1),
            scaleX: parsePercentAttr(reflection, 'sx', 1),
            scaleY: parsePercentAttr(reflection, 'sy', 1),
            skewX: parseAngleAttr(reflection, 'kx', 0),
            skewY: parseAngleAttr(reflection, 'ky', 0),
            rotateWithShape: parseBoolAttr(reflection, 'rotWithShape', true),
            ...(alignment !== undefined && { alignment }),
        };
    }
    /**
     * Resolves the color child of an effect element (transforms, including
     * a:alpha, are applied by the color resolver). Falls back to opaque black.
     */
    parseEffectColor(effectNode) {
        return this.colorResolver.resolveColorElement(effectNode) ?? { ...DEFAULT_SHADOW_COLOR };
    }
}
exports.EffectParser = EffectParser;
/**
 * Default effect parser factory.
 */
function createEffectParser(theme, logger) {
    return new EffectParser({ theme, logger });
}
