import { Colors } from '../types/index.js';
import { getXmlAttr, getXmlChild, getXmlChildren } from '../core/PptxParser.js';
import { ColorResolver, resolveEffectiveColorMap } from '../theme/ColorResolver.js';
import { percentageToDecimal } from '../core/UnitConverter.js';
import { RelationshipParser } from '../parsers/RelationshipParser.js';
import { ImageDecoder } from '../utils/ImageDecoder.js';
import { computeLinearGradientPoints } from './FillRenderer.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Renders slide backgrounds.
 */
export class BackgroundRenderer {
    logger;
    theme;
    colorResolver;
    /** Cached RelationshipParser instance (created lazily) */
    relationshipParser = null;
    /** Cached ImageDecoder instance (created lazily) */
    imageDecoder = null;
    /** Cached parser reference for relationship resolution */
    cachedParser = null;
    constructor(theme, logger) {
        this.logger = logger ?? createLogger('warn', 'BackgroundRenderer');
        this.theme = theme;
        this.colorResolver = new ColorResolver(theme.colors);
    }
    /**
     * Builds a color resolver honoring the effective color map for a slide
     * (slide p:clrMapOvr > layout p:clrMapOvr > master p:clrMap > default).
     */
    createColorResolverForSlide(slideNode, layoutNode, masterNode) {
        const colorMap = resolveEffectiveColorMap(masterNode, layoutNode, slideNode);
        return new ColorResolver(this.theme.colors, colorMap);
    }
    /**
     * Renders the background for a slide.
     * Follows the inheritance chain: slide -> layout -> master
     */
    renderBackground(ctx, canvas, slideNode, layoutNode, masterNode, overrideColor) {
        const width = canvas.width;
        const height = canvas.height;
        // If override color is specified, use it
        if (overrideColor) {
            const color = this.colorResolver.parseHexColor(overrideColor);
            this.fillSolid(ctx, width, height, color);
            return;
        }
        // Resolve background from inheritance chain (without path tracking for sync method)
        const resolver = this.createColorResolverForSlide(slideNode, layoutNode, masterNode);
        const background = this.resolveBackgroundFromChain(slideNode, layoutNode, masterNode, resolver);
        // Render the background (sync version - no picture support)
        this.renderBackgroundFill(ctx, width, height, background, false);
    }
    /**
     * Renders the background for a slide with async support for picture backgrounds.
     * Follows the inheritance chain: slide -> layout -> master
     *
     * @param ctx Canvas 2D rendering context
     * @param canvas Canvas to render to
     * @param slideNode Slide XML node
     * @param parser PPTX parser for accessing resources
     * @param slidePath Path to the slide file (e.g., ppt/slides/slide1.xml)
     * @param layoutNode Optional layout XML node
     * @param layoutPath Optional path to the layout file (e.g., ppt/slideLayouts/slideLayout1.xml)
     * @param masterNode Optional master XML node
     * @param masterPath Optional path to the master file
     * @param overrideColor Optional background color override
     */
    async renderBackgroundAsync(ctx, canvas, slideNode, parser, slidePath, layoutNode, layoutPath, masterNode, masterPath, overrideColor) {
        const width = canvas.width;
        const height = canvas.height;
        // If override color is specified, use it
        if (overrideColor) {
            const color = this.colorResolver.parseHexColor(overrideColor);
            this.fillSolid(ctx, width, height, color);
            return;
        }
        // Resolve background from inheritance chain with path tracking
        const resolver = this.createColorResolverForSlide(slideNode, layoutNode, masterNode);
        const { background, sourcePath } = this.resolveBackgroundFromChainWithPath(slideNode, slidePath, resolver, layoutNode, layoutPath, masterNode, masterPath);
        // Render the background (async version - with picture support)
        await this.renderBackgroundFillAsync(ctx, width, height, background, parser, sourcePath);
    }
    /**
     * Resolves background from the inheritance chain (slide -> layout -> master).
     * Simple version without path tracking for sync rendering.
     */
    resolveBackgroundFromChain(slideNode, layoutNode, masterNode, resolver = this.colorResolver) {
        // Try slide first
        let background = this.parseBackground(slideNode, resolver);
        if (background)
            return background;
        // Try layout
        if (layoutNode) {
            background = this.parseBackground(layoutNode, resolver);
            if (background)
                return background;
        }
        // Try master
        if (masterNode) {
            background = this.parseBackground(masterNode, resolver);
            if (background)
                return background;
        }
        return undefined;
    }
    /**
     * Resolves background from the inheritance chain with path tracking.
     * Required for async rendering to resolve picture relationships from the correct source.
     */
    resolveBackgroundFromChainWithPath(slideNode, slidePath, resolver, layoutNode, layoutPath, masterNode, masterPath) {
        // Try slide first
        let background = this.parseBackground(slideNode, resolver);
        if (background) {
            if (background.pictureFill) {
                background.pictureFill.source = 'slide';
            }
            return { background, sourcePath: slidePath };
        }
        // Try layout
        if (layoutNode) {
            background = this.parseBackground(layoutNode, resolver);
            if (background) {
                if (background.pictureFill) {
                    background.pictureFill.source = 'layout';
                }
                return { background, sourcePath: layoutPath ?? slidePath };
            }
        }
        // Try master
        if (masterNode) {
            background = this.parseBackground(masterNode, resolver);
            if (background) {
                if (background.pictureFill) {
                    background.pictureFill.source = 'master';
                }
                return { background, sourcePath: masterPath ?? slidePath };
            }
        }
        return { background: undefined, sourcePath: slidePath };
    }
    /**
     * Renders background fill (sync version - no picture support).
     */
    renderBackgroundFill(ctx, width, height, background, _supportPicture) {
        if (!background) {
            this.logger.debug('No background found, using white default');
            this.fillSolid(ctx, width, height, Colors.white);
            return;
        }
        switch (background.type) {
            case 'solid':
                this.fillSolid(ctx, width, height, background.color ?? Colors.white);
                break;
            case 'gradient':
                this.renderGradient(ctx, width, height, background);
                break;
            case 'pattern':
                this.logger.debug('Pattern background not yet supported, using solid');
                this.fillSolid(ctx, width, height, background.color ?? Colors.white);
                break;
            case 'picture':
                // Picture backgrounds require async rendering - this sync method falls back to white
                this.logger.debug('Picture background detected, use renderBackgroundAsync for image support');
                this.fillSolid(ctx, width, height, Colors.white);
                break;
            default:
                this.fillSolid(ctx, width, height, Colors.white);
        }
    }
    /**
     * Renders background fill (async version - with picture support).
     */
    async renderBackgroundFillAsync(ctx, width, height, background, parser, sourcePath) {
        if (!background) {
            this.logger.debug('No background found, using white default');
            this.fillSolid(ctx, width, height, Colors.white);
            return;
        }
        switch (background.type) {
            case 'solid':
                this.fillSolid(ctx, width, height, background.color ?? Colors.white);
                break;
            case 'gradient':
                this.renderGradient(ctx, width, height, background);
                break;
            case 'pattern':
                this.logger.debug('Pattern background not yet supported, using solid');
                this.fillSolid(ctx, width, height, background.color ?? Colors.white);
                break;
            case 'picture':
                if (background.pictureFill) {
                    await this.fillPicture(ctx, width, height, background.pictureFill, parser, sourcePath);
                }
                else {
                    this.fillSolid(ctx, width, height, Colors.white);
                }
                break;
            default:
                this.fillSolid(ctx, width, height, Colors.white);
        }
    }
    /**
     * Renders gradient background (shared between sync and async).
     */
    renderGradient(ctx, width, height, background) {
        if (background.gradientStops && background.gradientStops.length >= 2) {
            if (background.isRadial) {
                this.fillRadialGradient(ctx, width, height, background.gradientStops);
            }
            else {
                this.fillLinearGradient(ctx, width, height, background.gradientStops, background.gradientAngle ?? 0);
            }
        }
        else {
            this.fillSolid(ctx, width, height, Colors.white);
        }
    }
    /**
     * Gets or creates the cached RelationshipParser instance.
     */
    getRelationshipParser(parser) {
        // Invalidate cache if parser changed
        if (this.cachedParser !== parser) {
            this.cachedParser = parser;
            this.relationshipParser = null;
        }
        this.relationshipParser ??= new RelationshipParser({
            parser,
            logger: this.logger.child?.('RelParser'),
        });
        return this.relationshipParser;
    }
    /**
     * Gets or creates the cached ImageDecoder instance.
     */
    getImageDecoder() {
        this.imageDecoder ??= new ImageDecoder({
            logger: this.logger.child?.('Decoder'),
        });
        return this.imageDecoder;
    }
    /**
     * Fills the canvas with a picture background.
     *
     * @param ctx Canvas 2D rendering context
     * @param width Canvas width
     * @param height Canvas height
     * @param pictureFill Picture fill data
     * @param parser PPTX parser for accessing resources
     * @param sourcePath Path to the source file for relationship resolution
     */
    async fillPicture(ctx, width, height, pictureFill, parser, sourcePath) {
        try {
            // Use cached RelationshipParser
            const relationshipParser = this.getRelationshipParser(parser);
            const mediaPath = await relationshipParser.resolveImageRelationship(sourcePath, pictureFill.blipRelId);
            if (!mediaPath) {
                this.logger.warn('Could not resolve background image relationship', {
                    relId: pictureFill.blipRelId,
                    source: sourcePath,
                });
                this.fillSolid(ctx, width, height, Colors.white);
                return;
            }
            // Load the image data from the PPTX
            const buffer = await parser.readBinary(mediaPath);
            // Use cached ImageDecoder
            const imageDecoder = this.getImageDecoder();
            const decoded = await imageDecoder.decode(buffer);
            // Draw the image stretched to fill the entire canvas
            ctx.drawImage(decoded.image, 0, 0, width, height);
            this.logger.debug('Filled picture background', {
                relId: pictureFill.blipRelId,
                source: pictureFill.source,
                mediaPath,
                imageWidth: decoded.width,
                imageHeight: decoded.height,
                canvasWidth: width,
                canvasHeight: height,
            });
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to load background image', {
                relId: pictureFill.blipRelId,
                error: message,
            });
            // Fallback to white background
            this.fillSolid(ctx, width, height, Colors.white);
        }
    }
    /**
     * Parses background fill from a slide/layout/master node.
     */
    parseBackground(node, resolver = this.colorResolver) {
        // Look for cSld -> bg -> bgPr or cSld -> bg -> bgRef
        const cSld = getXmlChild(node, 'p:cSld');
        if (!cSld)
            return undefined;
        const bg = getXmlChild(cSld, 'p:bg');
        if (!bg)
            return undefined;
        // Check for background properties (explicit fill)
        const bgPr = getXmlChild(bg, 'p:bgPr');
        if (bgPr) {
            return this.parseBgProperties(bgPr, resolver);
        }
        // Check for background reference (theme style)
        const bgRef = getXmlChild(bg, 'p:bgRef');
        if (bgRef) {
            return this.parseBgReference(bgRef, resolver);
        }
        return undefined;
    }
    /**
     * Parses explicit background properties.
     */
    parseBgProperties(bgPr, resolver) {
        // Check for solid fill
        const solidFill = getXmlChild(bgPr, 'a:solidFill');
        if (solidFill) {
            const color = resolver.resolveColorElement(solidFill);
            if (color) {
                return { type: 'solid', color };
            }
        }
        // Check for gradient fill
        const gradFill = getXmlChild(bgPr, 'a:gradFill');
        if (gradFill) {
            return this.parseGradientFill(gradFill, resolver);
        }
        // Check for pattern fill
        const pattFill = getXmlChild(bgPr, 'a:pattFill');
        if (pattFill) {
            // Get foreground color as fallback
            const fgClr = getXmlChild(pattFill, 'a:fgClr');
            const color = fgClr ? resolver.resolveColorElement(fgClr) : undefined;
            return { type: 'pattern', color };
        }
        // Check for picture fill
        const blipFill = getXmlChild(bgPr, 'a:blipFill');
        if (blipFill) {
            const blip = getXmlChild(blipFill, 'a:blip');
            const blipRelId = blip ? getXmlAttr(blip, 'r:embed') : undefined;
            if (blipRelId) {
                return {
                    type: 'picture',
                    pictureFill: {
                        blipRelId,
                        source: 'slide', // Will be updated by resolveBackgroundFromChainWithPath
                    },
                };
            }
            // No valid blip reference, fallback to no fill
            return { type: 'none' };
        }
        // Check for no fill
        const noFill = getXmlChild(bgPr, 'a:noFill');
        if (noFill) {
            return { type: 'none' };
        }
        return undefined;
    }
    /**
     * Parses a background reference (p:bgRef) by resolving its index into the
     * theme style matrix (ECMA-376 §20.1.4.2.10 idx semantics: 0/1000 = no
     * fill, 1-999 = fillStyleLst, 1001+ = bgFillStyleLst). The bgRef child
     * color substitutes phClr placeholders, so template gradient backgrounds
     * render as real gradients.
     */
    parseBgReference(bgRef, resolver) {
        const idx = parseInt(getXmlAttr(bgRef, 'idx') ?? '0', 10);
        const phClr = resolver.resolveColorElement(bgRef);
        this.logger.debug('Background reference to theme style', { idx });
        if (!Number.isFinite(idx) || idx <= 0 || idx === 1000) {
            return { type: 'none' };
        }
        const formatScheme = this.theme.formatScheme;
        const themeFill = idx < 1000
            ? formatScheme?.fillStyles[idx - 1]
            : formatScheme?.backgroundFillStyles[idx - 1001];
        if (!themeFill) {
            // No parsed style matrix (or out-of-range index): fall back to the
            // reference color as a solid background
            return phClr ? { type: 'solid', color: phClr } : undefined;
        }
        return this.themeFillToBackground(themeFill, phClr, resolver);
    }
    /**
     * Converts a theme style-matrix fill into a ParsedBackground, substituting
     * phClr placeholders with the bgRef color.
     */
    themeFillToBackground(themeFill, phClr, resolver) {
        switch (themeFill.type) {
            case 'none':
                return { type: 'none' };
            case 'solid': {
                const color = resolver.resolveThemeStyleColor(themeFill.color, phClr);
                return color ? { type: 'solid', color } : undefined;
            }
            case 'gradient': {
                const stops = [];
                for (const stop of themeFill.stops) {
                    const color = resolver.resolveThemeStyleColor(stop.color, phClr);
                    if (color) {
                        stops.push({ position: stop.position, color });
                    }
                }
                if (stops.length >= 2) {
                    return {
                        type: 'gradient',
                        gradientStops: stops,
                        gradientAngle: themeFill.angle ?? 0,
                        isRadial: themeFill.isRadial ?? false,
                    };
                }
                if (stops.length === 1 && stops[0]) {
                    return { type: 'solid', color: stops[0].color };
                }
                return undefined;
            }
            case 'pattern': {
                const color = resolver.resolveThemeStyleColor(themeFill.foregroundColor, phClr) ?? phClr;
                return { type: 'pattern', color };
            }
        }
    }
    /**
     * Parses gradient fill properties.
     */
    parseGradientFill(gradFill, resolver) {
        const stops = [];
        // Get gradient stops
        const gsLst = getXmlChild(gradFill, 'a:gsLst');
        if (gsLst) {
            const gsNodes = getXmlChildren(gsLst, 'a:gs');
            for (const gs of gsNodes) {
                const pos = getXmlAttr(gs, 'pos');
                const position = pos !== undefined ? percentageToDecimal(parseInt(pos, 10)) : 0;
                const color = resolver.resolveColorElement(gs);
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
                gradientStops: stops,
                gradientAngle: angle,
                isRadial: false,
            };
        }
        // Check for path gradient (radial)
        const path = getXmlChild(gradFill, 'a:path');
        if (path) {
            return {
                type: 'gradient',
                gradientStops: stops,
                isRadial: true,
            };
        }
        // Default to horizontal linear gradient
        return {
            type: 'gradient',
            gradientStops: stops,
            gradientAngle: 0,
            isRadial: false,
        };
    }
    /**
     * Fills the canvas with a solid color.
     */
    fillSolid(ctx, width, height, color) {
        ctx.fillStyle = this.colorResolver.rgbaToCss(color);
        ctx.fillRect(0, 0, width, height);
        this.logger.debug('Filled solid background', { color: this.colorResolver.rgbaToHex(color) });
    }
    /**
     * Fills the canvas with a linear gradient.
     */
    fillLinearGradient(ctx, width, height, stops, angleDegrees) {
        const centerX = width / 2;
        const centerY = height / 2;
        // Calculate the diagonal length to ensure gradient covers entire canvas
        const diagonal = Math.sqrt(width * width + height * height) / 2;
        // ECMA-376 a:lin@ang is clockwise from the positive x-axis in y-down space:
        // 0 = left-to-right, 90 = top-to-bottom
        const { x0, y0, x1, y1 } = computeLinearGradientPoints(centerX, centerY, diagonal, angleDegrees);
        const gradient = ctx.createLinearGradient(x0, y0, x1, y1);
        for (const stop of stops) {
            gradient.addColorStop(stop.position, this.colorResolver.rgbaToCss(stop.color));
        }
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, width, height);
        this.logger.debug('Filled linear gradient background', {
            angle: angleDegrees,
            stopCount: stops.length,
        });
    }
    /**
     * Fills the canvas with a radial gradient.
     */
    fillRadialGradient(ctx, width, height, stops) {
        const centerX = width / 2;
        const centerY = height / 2;
        const radius = Math.max(width, height) / 2;
        const gradient = ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        // Radial gradients in OpenXML go from outside to center, so reverse the stops
        for (const stop of stops) {
            gradient.addColorStop(1 - stop.position, this.colorResolver.rgbaToCss(stop.color));
        }
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, width, height);
        this.logger.debug('Filled radial gradient background', { stopCount: stops.length });
    }
    /**
     * Gets the background color if it's a solid fill (for contrast calculations).
     */
    getBackgroundColor(slideNode, layoutNode, masterNode) {
        const resolver = this.createColorResolverForSlide(slideNode, layoutNode, masterNode);
        const background = this.resolveBackgroundFromChain(slideNode, layoutNode, masterNode, resolver);
        if (background?.type === 'solid' && background.color) {
            return background.color;
        }
        // For gradient, return the first stop color
        if (background?.type === 'gradient' && background.gradientStops?.length) {
            return background.gradientStops[0]?.color;
        }
        return Colors.white;
    }
}
//# sourceMappingURL=BackgroundRenderer.js.map