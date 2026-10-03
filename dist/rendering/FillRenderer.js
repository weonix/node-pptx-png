/**
 * Renders fills (solid, gradient, pattern, picture) to canvas context.
 */
import { getXmlChild, getXmlChildren, getXmlAttr, hasXmlChild } from '../core/PptxParser.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import { percentageToDecimal } from '../core/UnitConverter.js';
import { applyPathToContext } from '../geometry/PathBuilder.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Computes the start/end points of a linear gradient line through a center point.
 * Per ECMA-376 (a:lin@ang), the gradient angle is measured clockwise from the
 * positive x-axis in y-down slide space: 0 = left-to-right, 90 = top-to-bottom,
 * 180 = right-to-left, 270 = bottom-to-top. Canvas coordinates are also y-down,
 * so the angle maps directly to (cos, sin) without adjustment.
 * @param centerX Center x of the fill area
 * @param centerY Center y of the fill area
 * @param halfDiagonal Half the diagonal length of the fill area (ensures coverage)
 * @param angleDegrees Gradient angle in degrees, clockwise from the positive x-axis
 * @returns Gradient line start and end points
 */
export function computeLinearGradientPoints(centerX, centerY, halfDiagonal, angleDegrees) {
    const angleRad = (angleDegrees * Math.PI) / 180;
    const dx = Math.cos(angleRad) * halfDiagonal;
    const dy = Math.sin(angleRad) * halfDiagonal;
    return {
        x0: centerX - dx,
        y0: centerY - dy,
        x1: centerX + dx,
        y1: centerY + dy,
    };
}
/**
 * Renders fills for shapes.
 */
export class FillRenderer {
    logger;
    colorResolver;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'FillRenderer');
        this.colorResolver = new ColorResolver(config.theme.colors, config.colorMap);
    }
    /**
     * Renders a fill to the canvas for the given path.
     * For picture fills, use renderFillAsync instead.
     * @param ctx Canvas 2D context
     * @param path Path to fill
     * @param fill Fill definition
     * @param bounds Bounding rectangle of the shape
     */
    renderFill(ctx, path, fill, bounds) {
        if (fill.type === 'none') {
            return;
        }
        ctx.save();
        // Apply path as clip region
        applyPathToContext(ctx, path, true);
        switch (fill.type) {
            case 'solid':
                this.renderSolidFill(ctx, fill);
                break;
            case 'gradient':
                this.renderGradientFill(ctx, fill, bounds);
                break;
            case 'pattern':
                // Pattern fills are complex - fallback to solid with foreground color
                this.logger.debug('Pattern fill rendered as solid (patterns not fully implemented)');
                ctx.fillStyle = this.colorResolver.rgbaToCss(fill.foregroundColor);
                ctx.fill();
                break;
            case 'picture':
                // Picture fills require async rendering - log a warning if called synchronously
                this.logger.debug('Picture fill requires async rendering, use renderFillAsync');
                break;
        }
        ctx.restore();
    }
    /**
     * Renders a fill to the canvas for the given path, with async support for picture fills.
     * @param ctx Canvas 2D context
     * @param path Path to fill
     * @param fill Fill definition
     * @param bounds Bounding rectangle of the shape
     * @param imageRenderer Optional ImageRenderer for picture fills
     */
    async renderFillAsync(ctx, path, fill, bounds, imageRenderer) {
        if (fill.type === 'none') {
            return;
        }
        ctx.save();
        // Apply path as clip region
        applyPathToContext(ctx, path, true);
        switch (fill.type) {
            case 'solid':
                this.renderSolidFill(ctx, fill);
                break;
            case 'gradient':
                this.renderGradientFill(ctx, fill, bounds);
                break;
            case 'pattern':
                this.logger.debug('Pattern fill rendered as solid (patterns not fully implemented)');
                ctx.fillStyle = this.colorResolver.rgbaToCss(fill.foregroundColor);
                ctx.fill();
                break;
            case 'picture':
                // Clip to the shape path so the image is confined to non-rectangular
                // geometry (drawImage ignores the current path, unlike ctx.fill())
                ctx.clip();
                await this.renderPictureFill(ctx, fill, bounds, imageRenderer);
                break;
        }
        ctx.restore();
    }
    /**
     * Renders a picture fill.
     * Delegates all parsing and rendering to ImageRenderer.
     */
    async renderPictureFill(ctx, fill, bounds, imageRenderer) {
        if (!imageRenderer) {
            this.logger.warn('Picture fill requires ImageRenderer');
            return;
        }
        if (!fill.blipFillNode) {
            this.logger.warn('Picture fill has no blipFillNode');
            return;
        }
        try {
            // Delegate to ImageRenderer for parsing and rendering
            await imageRenderer.renderPictureFill(ctx, fill.blipFillNode, bounds);
            this.logger.debug('Rendered picture fill', {
                relId: fill.relationshipId,
            });
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to render picture fill', { error: message });
        }
    }
    /**
     * Renders a solid fill.
     */
    renderSolidFill(ctx, fill) {
        ctx.fillStyle = this.colorResolver.rgbaToCss(fill.color);
        ctx.fill();
        this.logger.debug('Rendered solid fill', {
            color: this.colorResolver.rgbaToHex(fill.color),
        });
    }
    /**
     * Renders a gradient fill.
     */
    renderGradientFill(ctx, fill, bounds) {
        if (!fill.stops || fill.stops.length < 2) {
            // Need at least 2 stops for a gradient
            if (fill.stops?.length === 1 && fill.stops[0]) {
                ctx.fillStyle = this.colorResolver.rgbaToCss(fill.stops[0].color);
                ctx.fill();
            }
            return;
        }
        let gradient;
        if (fill.isRadial) {
            gradient = this.createRadialGradient(ctx, bounds, fill);
        }
        else {
            gradient = this.createLinearGradient(ctx, bounds, fill);
        }
        // Add color stops
        for (const stop of fill.stops) {
            gradient.addColorStop(stop.position, this.colorResolver.rgbaToCss(stop.color));
        }
        ctx.fillStyle = gradient;
        ctx.fill();
        this.logger.debug('Rendered gradient fill', {
            type: fill.isRadial ? 'radial' : 'linear',
            angle: fill.angle,
            stopCount: fill.stops.length,
        });
    }
    /**
     * Creates a linear gradient for the given bounds and fill.
     */
    createLinearGradient(ctx, bounds, fill) {
        const centerX = bounds.x + bounds.width / 2;
        const centerY = bounds.y + bounds.height / 2;
        // Calculate diagonal length to ensure gradient covers entire shape
        const diagonal = Math.sqrt(bounds.width * bounds.width + bounds.height * bounds.height) / 2;
        // ECMA-376 a:lin@ang is clockwise from the positive x-axis in y-down space:
        // 0 = left-to-right, 90 = top-to-bottom
        const { x0, y0, x1, y1 } = computeLinearGradientPoints(centerX, centerY, diagonal, fill.angle ?? 0);
        return ctx.createLinearGradient(x0, y0, x1, y1);
    }
    /**
     * Creates a radial gradient for the given bounds and fill.
     */
    createRadialGradient(ctx, bounds, fill) {
        const centerX = bounds.x + bounds.width * (fill.centerX ?? 0.5);
        const centerY = bounds.y + bounds.height * (fill.centerY ?? 0.5);
        const radius = Math.max(bounds.width, bounds.height) / 2;
        return ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
    }
    /**
     * Parses fill properties from a shape properties node.
     * @param spPr Shape properties node
     * @returns Parsed fill or undefined if no fill specified
     */
    parseFill(spPr) {
        if (!spPr)
            return undefined;
        // Check for no fill
        if (hasXmlChild(spPr, 'a:noFill')) {
            return { type: 'none' };
        }
        // Check for solid fill
        const solidFill = getXmlChild(spPr, 'a:solidFill');
        if (solidFill) {
            return this.parseSolidFill(solidFill);
        }
        // Check for gradient fill
        const gradFill = getXmlChild(spPr, 'a:gradFill');
        if (gradFill) {
            return this.parseGradientFill(gradFill);
        }
        // Check for pattern fill
        const pattFill = getXmlChild(spPr, 'a:pattFill');
        if (pattFill) {
            return this.parsePatternFill(pattFill);
        }
        // Check for picture fill (blipFill)
        const blipFill = getXmlChild(spPr, 'a:blipFill');
        if (blipFill) {
            return this.parsePictureFill(blipFill);
        }
        return undefined;
    }
    /**
     * Parses a solid fill element.
     */
    parseSolidFill(solidFill) {
        const color = this.colorResolver.resolveColorElement(solidFill);
        if (!color)
            return undefined;
        return {
            type: 'solid',
            color,
        };
    }
    /**
     * Parses a gradient fill element.
     */
    parseGradientFill(gradFill) {
        const stops = [];
        // Get gradient stops
        const gsLst = getXmlChild(gradFill, 'a:gsLst');
        if (gsLst) {
            const gsNodes = getXmlChildren(gsLst, 'a:gs');
            for (const gs of gsNodes) {
                const pos = getXmlAttr(gs, 'pos');
                const position = pos !== undefined ? percentageToDecimal(parseInt(pos, 10)) : 0;
                const color = this.colorResolver.resolveColorElement(gs);
                if (color) {
                    stops.push({ position, color });
                }
            }
        }
        if (stops.length < 2) {
            return undefined;
        }
        // Sort stops by position
        stops.sort((a, b) => a.position - b.position);
        // Check for linear gradient
        const lin = getXmlChild(gradFill, 'a:lin');
        if (lin) {
            const ang = getXmlAttr(lin, 'ang');
            const angle = ang !== undefined ? parseInt(ang, 10) / 60000 : 0;
            return {
                type: 'gradient',
                stops,
                angle,
                isRadial: false,
            };
        }
        // Check for path/radial gradient
        const path = getXmlChild(gradFill, 'a:path');
        if (path) {
            const fillToRect = getXmlChild(path, 'a:fillToRect');
            let centerX = 0.5;
            let centerY = 0.5;
            if (fillToRect) {
                const l = getXmlAttr(fillToRect, 'l');
                const t = getXmlAttr(fillToRect, 't');
                const r = getXmlAttr(fillToRect, 'r');
                const b = getXmlAttr(fillToRect, 'b');
                const left = l !== undefined ? percentageToDecimal(parseInt(l, 10)) : 0;
                const top = t !== undefined ? percentageToDecimal(parseInt(t, 10)) : 0;
                const right = r !== undefined ? percentageToDecimal(parseInt(r, 10)) : 1;
                const bottom = b !== undefined ? percentageToDecimal(parseInt(b, 10)) : 1;
                centerX = (left + right) / 2;
                centerY = (top + bottom) / 2;
            }
            // Reverse stops for radial (OpenXML radial goes edge to center)
            const reversedStops = stops
                .map((s) => ({ position: 1 - s.position, color: s.color }))
                .sort((a, b) => a.position - b.position);
            return {
                type: 'gradient',
                stops: reversedStops,
                isRadial: true,
                centerX,
                centerY,
            };
        }
        // Default to horizontal linear gradient
        return {
            type: 'gradient',
            stops,
            angle: 0,
            isRadial: false,
        };
    }
    /**
     * Parses a pattern fill element.
     */
    parsePatternFill(pattFill) {
        // Get foreground and background colors
        const fgClr = getXmlChild(pattFill, 'a:fgClr');
        const bgClr = getXmlChild(pattFill, 'a:bgClr');
        const foregroundColor = fgClr
            ? (this.colorResolver.resolveColorElement(fgClr) ?? { r: 0, g: 0, b: 0, a: 255 })
            : { r: 0, g: 0, b: 0, a: 255 };
        const backgroundColor = bgClr
            ? (this.colorResolver.resolveColorElement(bgClr) ?? { r: 255, g: 255, b: 255, a: 255 })
            : { r: 255, g: 255, b: 255, a: 255 };
        const preset = getXmlAttr(pattFill, 'prst') ?? 'solid';
        return {
            type: 'pattern',
            preset,
            foregroundColor,
            backgroundColor,
        };
    }
    /**
     * Parses a picture fill element.
     * Note: Full parsing of srcRect, tile, fillRect is deferred to ImageRenderer
     * which will parse the blipFillNode when rendering. This avoids duplicate parsing logic.
     */
    parsePictureFill(blipFill) {
        const blip = getXmlChild(blipFill, 'a:blip');
        const embedId = blip ? (getXmlAttr(blip, 'r:embed') ?? '') : '';
        // Store the blipFillNode - ImageRenderer.parseBlipFill() will handle
        // the detailed parsing of srcRect, tile, fillRect, and stretch settings
        return {
            type: 'picture',
            relationshipId: embedId,
            blipFillNode: blipFill,
        };
    }
    /**
     * Resolves a theme style-matrix fill (from fillStyleLst/bgFillStyleLst)
     * into a renderable Fill, substituting phClr placeholders with the color
     * supplied by the style reference (a:fillRef child color).
     * @param themeFill The theme fill definition
     * @param phClr Substitution color for phClr placeholders
     * @returns The resolved fill, or undefined when required colors are missing
     */
    resolveThemeFill(themeFill, phClr) {
        switch (themeFill.type) {
            case 'none':
                return { type: 'none' };
            case 'solid': {
                const color = this.colorResolver.resolveThemeStyleColor(themeFill.color, phClr);
                return color ? { type: 'solid', color } : undefined;
            }
            case 'gradient': {
                const stops = [];
                for (const stop of themeFill.stops) {
                    const color = this.colorResolver.resolveThemeStyleColor(stop.color, phClr);
                    if (color) {
                        stops.push({ position: stop.position, color });
                    }
                }
                if (stops.length === 0)
                    return undefined;
                if (stops.length === 1 && stops[0]) {
                    return { type: 'solid', color: stops[0].color };
                }
                if (themeFill.isRadial) {
                    // OpenXML radial gradients run edge-to-center; reverse for canvas
                    const reversed = stops
                        .map((s) => ({ position: 1 - s.position, color: s.color }))
                        .sort((a, b) => a.position - b.position);
                    return {
                        type: 'gradient',
                        stops: reversed,
                        isRadial: true,
                        centerX: themeFill.centerX,
                        centerY: themeFill.centerY,
                    };
                }
                return {
                    type: 'gradient',
                    stops,
                    angle: themeFill.angle ?? 0,
                    isRadial: false,
                };
            }
            case 'pattern': {
                const foregroundColor = this.colorResolver.resolveThemeStyleColor(themeFill.foregroundColor, phClr) ?? { r: 0, g: 0, b: 0, a: 255 };
                const backgroundColor = this.colorResolver.resolveThemeStyleColor(themeFill.backgroundColor, phClr) ?? { r: 255, g: 255, b: 255, a: 255 };
                return {
                    type: 'pattern',
                    preset: themeFill.preset,
                    foregroundColor,
                    backgroundColor,
                };
            }
        }
    }
    /**
     * Gets fill style string for a solid color (utility method).
     */
    getFillStyle(color) {
        return this.colorResolver.rgbaToCss(color);
    }
}
/**
 * Default fill renderer factory.
 */
export function createFillRenderer(theme, logger) {
    return new FillRenderer({ theme, logger });
}
//# sourceMappingURL=FillRenderer.js.map