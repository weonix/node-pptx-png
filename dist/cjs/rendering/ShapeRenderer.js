"use strict";
/**
 * Main shape rendering orchestration.
 * Coordinates geometry calculation, transforms, fills, strokes, and text.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ShapeRenderer = void 0;
exports.createShapeRenderer = createShapeRenderer;
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const PlaceholderResolver_js_1 = require("../core/PlaceholderResolver.js");
const ColorResolver_js_1 = require("../theme/ColorResolver.js");
const TransformCalculator_js_1 = require("../geometry/TransformCalculator.js");
const PresetGeometryCalculator_js_1 = require("../geometry/PresetGeometryCalculator.js");
const GeometryEngine_js_1 = require("../geometry/GeometryEngine.js");
const CustomGeometryParser_js_1 = require("../geometry/CustomGeometryParser.js");
const FillRenderer_js_1 = require("./FillRenderer.js");
const StrokeRenderer_js_1 = require("./StrokeRenderer.js");
const TextRenderer_js_1 = require("./TextRenderer.js");
const TextParser_js_1 = require("../parsers/TextParser.js");
const EffectParser_js_1 = require("../parsers/EffectParser.js");
const EffectRenderer_js_1 = require("./EffectRenderer.js");
const ImageRenderer_js_1 = require("./ImageRenderer.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Renders shapes to canvas.
 */
class ShapeRenderer {
    logger;
    theme;
    scaleX;
    scaleY;
    parser;
    sourcePath;
    layoutNode;
    masterNode;
    /** p:defaultTextStyle extracted from the presentation node, if provided */
    defaultTextStyle;
    colorMap;
    colorResolver;
    transformCalculator;
    geometryCalculator;
    placeholderResolver;
    /** Fill renderer for external use */
    fillRenderer;
    /** Stroke renderer for external use */
    strokeRenderer;
    /** Text renderer for external use */
    textRenderer;
    /** Text parser for external use */
    textParser;
    /** Effect parser for external use */
    effectParser;
    /** Silhouette-based effect renderer (perspective shadows, glow, ...) */
    effectRenderer;
    imageRenderer = null;
    warnings;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'ShapeRenderer');
        this.theme = config.theme;
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
        this.parser = config.parser;
        this.sourcePath = config.sourcePath;
        this.layoutNode = config.layoutNode;
        this.masterNode = config.masterNode;
        this.defaultTextStyle = config.presentationNode
            ? (0, PptxParser_js_1.getXmlChild)(config.presentationNode, 'p:defaultTextStyle')
            : undefined;
        this.colorMap = (0, ColorResolver_js_1.resolveEffectiveColorMap)(config.masterNode, config.layoutNode, config.slideNode);
        this.colorResolver = new ColorResolver_js_1.ColorResolver(config.theme.colors, this.colorMap);
        this.transformCalculator = new TransformCalculator_js_1.TransformCalculator();
        this.geometryCalculator = new PresetGeometryCalculator_js_1.PresetGeometryCalculator();
        this.placeholderResolver = new PlaceholderResolver_js_1.PlaceholderResolver(this.logger.child?.('Placeholder'));
        this.fillRenderer = new FillRenderer_js_1.FillRenderer({
            theme: config.theme,
            colorMap: this.colorMap,
            logger: this.logger.child?.('FillRenderer'),
        });
        this.strokeRenderer = new StrokeRenderer_js_1.StrokeRenderer({
            theme: config.theme,
            colorMap: this.colorMap,
            logger: this.logger.child?.('StrokeRenderer'),
        });
        this.warnings = config.warnings;
        this.textRenderer = new TextRenderer_js_1.TextRenderer({
            theme: config.theme,
            scaleX: config.scaleX,
            scaleY: config.scaleY,
            logger: this.logger.child?.('TextRenderer'),
            warnings: config.warnings,
        });
        this.textParser = new TextParser_js_1.TextParser({
            theme: config.theme,
            colorMap: this.colorMap,
            slideNumber: config.slideNumber,
            logger: this.logger.child?.('TextParser'),
        });
        this.effectParser = new EffectParser_js_1.EffectParser({
            theme: config.theme,
            colorMap: this.colorMap,
            logger: this.logger.child?.('EffectParser'),
        });
        this.effectRenderer = new EffectRenderer_js_1.EffectRenderer({
            scaleX: config.scaleX,
            scaleY: config.scaleY,
            logger: this.logger.child?.('EffectRenderer'),
        });
        // Initialize image renderer if parser and source path are provided
        if (config.parser && config.sourcePath) {
            this.imageRenderer = new ImageRenderer_js_1.ImageRenderer({
                parser: config.parser,
                sourcePath: config.sourcePath,
                scaleX: config.scaleX,
                scaleY: config.scaleY,
                logger: this.logger.child?.('ImageRenderer'),
                warnings: config.warnings,
            });
        }
    }
    /**
     * Renders a shape element to the canvas.
     * @param ctx Canvas 2D context
     * @param spNode Shape XML node (p:sp)
     */
    async renderShape(ctx, spNode) {
        // Parse the shape
        const shape = this.parseShape(spNode);
        if (!shape || shape.hidden) {
            return;
        }
        this.logger.debug('Rendering shape', {
            id: shape.id,
            name: shape.name,
            geometry: shape.geometryType,
            hasText: !!shape.textBody,
        });
        // Create the geometry path
        const bounds = {
            x: 0,
            y: 0,
            width: shape.transform.width,
            height: shape.transform.height,
        };
        const paths = this.createShapePaths(shape.geometryType, bounds, shape.adjustValues, shape.customGeometry);
        if (paths.length === 0) {
            return;
        }
        // Apply transform and render. The try/finally guarantees the context
        // stack is rebalanced even when a fill/stroke/text pass throws (callers
        // like DiagramRenderer contain per-shape errors and keep rendering).
        ctx.save();
        try {
            this.transformCalculator.applyTransform(ctx, shape.transform);
            // Get fill color for text contrast calculation
            let fillColor;
            if (shape.fill?.type === 'solid' && paths.some((p) => p.fill !== false)) {
                fillColor = shape.fill.color;
            }
            await this.renderEffectsAndBody(ctx, paths, bounds, shape.fill, shape.stroke, shape.effects);
            // Finally render text if present
            if (shape.textBody && shape.textBody.paragraphs.length > 0) {
                // Default text color: the style fontRef color when the shape carries a
                // p:style, otherwise fall back to the fill-contrast heuristic
                const defaultTextColor = shape.styleFontColor ??
                    (fillColor ? this.textRenderer.getContrastingColor(fillColor) : undefined);
                // Get text bounds (may be different from shape bounds for non-rectangular shapes)
                const textBounds = this.getShapeTextBounds(shape.geometryType, bounds, shape.adjustValues, shape.customGeometry);
                this.textRenderer.renderText(ctx, shape.textBody, textBounds, defaultTextColor);
            }
        }
        finally {
            ctx.restore();
        }
    }
    /**
     * Renders a shape's effects and its fill/stroke passes in z-order:
     * reflection, outer shadow, glow beneath the shape; then the shape body
     * (through the soft-edge mask when present); then the inner shadow on
     * top of the fill. Plain offset shadows keep the fast ctx.shadow* path;
     * perspective (sx/sy/kx/ky) and picture-fill shadows go through the
     * silhouette-based EffectRenderer.
     *
     * Completes synchronously unless an async picture fill runs
     * (allowAsyncFills with an initialized ImageRenderer), so the
     * synchronous group-rendering path keeps its draw ordering.
     *
     * @param ctx Canvas context, already in the shape's local space
     * @param paths Geometry paths in local pixel space
     * @param bounds Local-space bounds of the shape
     * @param fill Effective fill
     * @param stroke Effective stroke
     * @param effects Effective effects
     * @param allowAsyncFills Whether picture fills may render asynchronously
     * @returns A promise only when an async picture fill is rendering
     */
    renderEffectsAndBody(ctx, paths, bounds, fill, stroke, effects, allowAsyncFills = true) {
        const source = this.createSilhouetteSource(paths, bounds, stroke, fill);
        const outerShadow = effects?.outerShadow;
        const silhouetteShadow = !!outerShadow && this.effectRenderer.needsSilhouetteShadow(outerShadow, fill);
        const fastShadow = silhouetteShadow ? undefined : outerShadow;
        const drawContent = this.makeFillStrokeCallback(paths, bounds, fill, stroke, allowAsyncFills);
        // Beneath the shape: reflection, then outer shadow, then glow.
        // Reflection completes synchronously here: it is skipped for picture
        // fills, the only asynchronous content source (noted approximation).
        if (effects?.reflection && fill?.type !== 'picture') {
            void this.effectRenderer.renderReflection(ctx, source, effects.reflection, drawContent);
        }
        if (outerShadow && silhouetteShadow) {
            this.effectRenderer.renderOuterShadow(ctx, source, outerShadow);
        }
        if (effects?.glow) {
            this.effectRenderer.renderGlow(ctx, source, effects.glow);
        }
        // The shape body: through the soft-edge mask when present (a plain
        // shadow then wraps the composited drawImage), else the direct passes
        let bodyDone;
        if (effects?.softEdge) {
            if (fastShadow) {
                this.setOuterShadow(ctx, fastShadow);
            }
            bodyDone = this.effectRenderer.renderWithSoftEdge(ctx, source, effects.softEdge, drawContent);
        }
        else {
            bodyDone = this.renderBodyPasses(ctx, paths, bounds, fill, stroke, fastShadow, allowAsyncFills);
        }
        const finish = () => {
            if (fastShadow) {
                this.clearShadow(ctx);
            }
            if (effects?.innerShadow) {
                this.effectRenderer.renderInnerShadow(ctx, source, effects.innerShadow);
            }
        };
        if (bodyDone && typeof bodyDone.then === 'function') {
            return bodyDone.then(finish);
        }
        finish();
    }
    /**
     * Renders the fill/stroke passes for each path. A plain outer shadow is
     * applied to exactly one pass — the first fill-enabled path, or the
     * first stroked path when nothing fills (stroke-only shapes still cast
     * shadows). The first path of a geometry carries the silhouette;
     * shadowing every subpath would double-darken overlaps.
     * @returns A promise only when an async picture fill is rendering
     */
    renderBodyPasses(ctx, paths, bounds, fill, stroke, outerShadow, allowAsyncFills) {
        if (allowAsyncFills && fill?.type === 'picture' && this.imageRenderer) {
            return this.renderBodyPassesAsync(ctx, paths, bounds, fill, stroke, outerShadow);
        }
        const hasFillPass = !!fill && fill.type !== 'none' && paths.some((p) => p.fill !== false);
        let pendingShadow = outerShadow;
        // Render each path: fill first, then stroke (per-path flags apply)
        for (const path of paths) {
            if (fill && fill.type !== 'none' && path.fill !== false) {
                const shadow = pendingShadow;
                if (shadow) {
                    this.setOuterShadow(ctx, shadow);
                }
                this.fillRenderer.renderFill(ctx, path, fill, bounds);
                if (shadow) {
                    this.clearShadow(ctx);
                    pendingShadow = undefined;
                }
            }
            if (stroke && path.stroke !== false) {
                const shadow = hasFillPass ? undefined : pendingShadow;
                if (shadow) {
                    this.setOuterShadow(ctx, shadow);
                }
                this.strokeRenderer.renderStroke(ctx, path, stroke, this.scaleX, this.scaleY);
                if (shadow) {
                    this.clearShadow(ctx);
                    pendingShadow = undefined;
                }
            }
        }
    }
    /**
     * Async variant of the body passes for picture fills. Picture-fill outer
     * shadows always take the silhouette path (ctx.shadow* output would be
     * clipped away with the fill), so no shadow state is set here; the
     * shadow parameter guards the stroke-only case for completeness.
     */
    async renderBodyPassesAsync(ctx, paths, bounds, fill, stroke, outerShadow) {
        let pendingStrokeShadow = outerShadow;
        for (const path of paths) {
            if (path.fill !== false) {
                await this.fillRenderer.renderFillAsync(ctx, path, fill, bounds, this.imageRenderer ?? undefined);
            }
            if (stroke && path.stroke !== false) {
                const shadow = pendingStrokeShadow;
                if (shadow) {
                    this.setOuterShadow(ctx, shadow);
                }
                this.strokeRenderer.renderStroke(ctx, path, stroke, this.scaleX, this.scaleY);
                if (shadow) {
                    this.clearShadow(ctx);
                    pendingStrokeShadow = undefined;
                }
            }
        }
    }
    /**
     * Builds the silhouette source EffectRenderer works from: the shape's
     * paths, bounds, and stroke width (used for stroke-only silhouettes and
     * layer padding).
     */
    createSilhouetteSource(paths, bounds, stroke, fill) {
        const silhouetteSource = {
            paths,
            bounds,
            hasFill: fill !== undefined && fill.type !== 'none',
        };
        if (stroke) {
            silhouetteSource.strokeWidthPx = Math.max((0, UnitConverter_js_1.emuToPixels)(stroke.width) * Math.min(this.scaleX, this.scaleY), 0.5);
        }
        return silhouetteSource;
    }
    /**
     * Builds the content callback effect layers redraw the shape through
     * (soft edge and reflection): the plain fill/stroke passes without any
     * shadow state. Asynchronous only for picture fills when allowed.
     */
    makeFillStrokeCallback(paths, bounds, fill, stroke, allowAsyncFills) {
        const imageRenderer = this.imageRenderer;
        if (allowAsyncFills && fill?.type === 'picture' && imageRenderer) {
            return async (target) => {
                for (const path of paths) {
                    if (path.fill !== false) {
                        await this.fillRenderer.renderFillAsync(target, path, fill, bounds, imageRenderer);
                    }
                    if (stroke && path.stroke !== false) {
                        this.strokeRenderer.renderStroke(target, path, stroke, this.scaleX, this.scaleY);
                    }
                }
            };
        }
        return (target) => {
            for (const path of paths) {
                if (fill && fill.type !== 'none' && path.fill !== false) {
                    this.fillRenderer.renderFill(target, path, fill, bounds);
                }
                if (stroke && path.stroke !== false) {
                    this.strokeRenderer.renderStroke(target, path, stroke, this.scaleX, this.scaleY);
                }
            }
        };
    }
    /**
     * Creates the renderable paths for a shape's geometry.
     * Unknown preset geometries and unparsable custom geometries fall back to
     * a rectangle (with a logged warning) so fill, stroke, and text still
     * render instead of the shape vanishing.
     *
     * @returns Paths in pixel space; empty only for degenerate bounds
     */
    createShapePaths(geometryType, bounds, adjustValues, customGeometry) {
        if (customGeometry) {
            try {
                const paths = (0, GeometryEngine_js_1.buildGeometryPaths)(customGeometry.def, bounds, undefined, customGeometry.space);
                if (paths.length > 0) {
                    return paths;
                }
            }
            catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                this.logger.debug('Custom geometry evaluation failed', { error: message });
            }
            this.logger.warn('Unsupported geometry type, falling back to rectangle', {
                type: 'custGeom',
            });
            this.warnings?.push({
                code: 'unsupported-geometry',
                message: 'Custom geometry could not be evaluated; shape rendered as rectangle',
                detail: { geometryType: 'custGeom' },
            });
        }
        else {
            const paths = this.geometryCalculator.createPaths(geometryType, bounds, adjustValues);
            if (paths && paths.length > 0) {
                return paths;
            }
            this.logger.warn('Unsupported geometry type, falling back to rectangle', {
                type: geometryType,
            });
            this.warnings?.push({
                code: 'unsupported-geometry',
                message: `Unsupported preset geometry "${geometryType}"; shape rendered as rectangle`,
                detail: { geometryType },
            });
        }
        const fallback = this.geometryCalculator.createPath('rect', bounds);
        return fallback ? [fallback] : [];
    }
    /**
     * Computes the text layout bounds for a shape, honoring the geometry's
     * text rectangle (custom geometry a:rect included).
     */
    getShapeTextBounds(geometryType, bounds, adjustValues, customGeometry) {
        if (customGeometry) {
            return (0, GeometryEngine_js_1.computeTextRect)(customGeometry.def, bounds, undefined, customGeometry.space) ?? bounds;
        }
        return this.geometryCalculator.getTextBounds(geometryType, bounds, adjustValues);
    }
    /**
     * Parses a shape XML node into renderable data.
     * @param spNode Shape XML node (p:sp)
     */
    parseShape(spNode) {
        // Get non-visual properties
        const nvSpPr = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:nvSpPr');
        const cNvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:cNvPr') : undefined;
        const nvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        // Check if this is a placeholder shape (a bare <p:ph/> is valid)
        const phNode = this.getPresentChild(nvPr, 'p:ph');
        const placeholderType = phNode ? (0, PptxParser_js_1.getXmlAttr)(phNode, 'type') : undefined;
        const placeholderIdx = phNode ? (0, PptxParser_js_1.getXmlAttr)(phNode, 'idx') : undefined;
        // Get shape properties
        const spPr = this.getSpPrNode(spNode);
        if (!spPr) {
            this.logger.debug('Shape has no spPr, skipping', { id });
            return undefined;
        }
        // Parse transform - first try from shape, then from layout placeholder
        let transform = this.transformCalculator.parseTransform(spPr);
        // If no transform and this is a placeholder, look it up from layout
        if (!transform && phNode) {
            transform = this.getPlaceholderTransform(placeholderType, placeholderIdx);
            if (transform) {
                this.logger.debug('Using placeholder transform from layout', {
                    id,
                    type: placeholderType,
                    idx: placeholderIdx,
                });
            }
        }
        if (!transform) {
            this.logger.debug('Shape has no transform, skipping', { id });
            return undefined;
        }
        const pixelTransform = this.transformCalculator.toPixelTransform(transform, this.scaleX, this.scaleY);
        // Parse geometry - for placeholders, may need to get from layout too
        const parsedGeometry = this.parseGeometry(spPr);
        const { adjustValues, customGeometry } = parsedGeometry;
        let { geometryType } = parsedGeometry;
        // Default to rect for placeholders without explicit geometry
        if (!geometryType && phNode) {
            geometryType = 'rect';
        }
        // Resolve the shape style references (p:style)
        const style = this.parseShapeStyle(spNode);
        // Parse fill - explicit spPr fill (including a:noFill) beats the style ref
        const fill = this.fillRenderer.parseFill(spPr) ?? style?.fill;
        // Parse stroke - explicit a:ln properties beat the style ref
        const stroke = this.strokeRenderer.parseStroke(spPr, style?.stroke);
        // Parse effects - an explicit spPr a:effectLst beats the style ref
        const effects = this.resolveShapeEffects(spPr, style);
        // Parse text body, merging the placeholder text-style inheritance chain
        // (master p:txStyles -> layout placeholder lstStyle) for placeholders.
        // Non-placeholder shapes (text boxes) still inherit the presentation's
        // p:defaultTextStyle (ECMA-376 §19.2.1.12), the document-wide base.
        const txBody = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:txBody');
        const inheritedListStyles = phNode
            ? this.getInheritedListStyles(placeholderType, placeholderIdx)
            : this.getTextBoxListStyles();
        const textBody = this.textParser.parseTextBody(txBody, inheritedListStyles);
        return {
            id,
            name,
            transform: pixelTransform,
            geometryType: geometryType || 'rect',
            customGeometry,
            adjustValues,
            fill,
            stroke,
            textBody,
            styleFontColor: style?.fontColor,
            effects,
            hidden,
        };
    }
    /**
     * Gets a child element, preserving the presence of empty elements. The XML
     * parser yields self-closed childless elements (e.g. the `<p:spPr/>` and
     * `<p:ph/>` placeholder shapes routinely carry) as empty strings, which
     * must not be confused with a missing element.
     */
    getPresentChild(node, name) {
        if (!node || typeof node !== 'object')
            return undefined;
        const child = node[name];
        if (child === undefined)
            return undefined;
        return typeof child === 'object' && child !== null ? child : {};
    }
    /**
     * Gets a shape's p:spPr node, treating an empty element as an empty node
     * (placeholder shapes get their transform from the layout).
     */
    getSpPrNode(spNode) {
        return this.getPresentChild(spNode, 'p:spPr');
    }
    /**
     * Resolves the inherited list-style chain for a placeholder shape, ordered
     * lowest precedence first (presentation defaultTextStyle, master txStyles
     * bucket, master placeholder lstStyle, then the matching layout
     * placeholder's lstStyle). Returns undefined when nothing is inherited.
     */
    getInheritedListStyles(placeholderType, placeholderIdx) {
        const chain = this.placeholderResolver.resolveTextStyleChain({
            type: placeholderType,
            idx: placeholderIdx !== undefined ? parseInt(placeholderIdx, 10) : undefined,
        }, this.layoutNode, this.masterNode, this.defaultTextStyle);
        return chain.length > 0 ? chain : undefined;
    }
    /**
     * Inherited list styles for NON-placeholder shapes (plain text boxes):
     * only the presentation part's p:defaultTextStyle applies (ECMA-376
     * §19.2.1.12) — master txStyles buckets and layout lstStyles are
     * placeholder-inheritance layers and never reach a text box.
     */
    getTextBoxListStyles() {
        return this.defaultTextStyle ? [this.defaultTextStyle] : undefined;
    }
    /**
     * Parses and resolves a shape's p:style element (a:fillRef, a:lnRef,
     * a:fontRef) against the theme's style matrix. Per ECMA-376 the style
     * supplies defaults that explicit spPr formatting overrides.
     * @param shapeNode The shape node (p:sp / p:cxnSp / p:pic)
     * @returns The resolved style, or undefined when the shape has no p:style
     */
    parseShapeStyle(shapeNode) {
        const styleNode = (0, PptxParser_js_1.getXmlChild)(shapeNode, 'p:style');
        if (!styleNode)
            return undefined;
        const style = {};
        const fillRef = (0, PptxParser_js_1.getXmlChild)(styleNode, 'a:fillRef');
        if (fillRef) {
            style.fill = this.resolveStyleFill(fillRef);
        }
        const lnRef = (0, PptxParser_js_1.getXmlChild)(styleNode, 'a:lnRef');
        if (lnRef) {
            style.stroke = this.resolveStyleStroke(lnRef);
        }
        const fontRef = (0, PptxParser_js_1.getXmlChild)(styleNode, 'a:fontRef');
        if (fontRef) {
            style.fontRefIdx = (0, PptxParser_js_1.getXmlAttr)(fontRef, 'idx');
            style.fontColor = this.colorResolver.resolveColorElement(fontRef);
        }
        const effectRef = (0, PptxParser_js_1.getXmlChild)(styleNode, 'a:effectRef');
        if (effectRef) {
            style.outerShadow = this.resolveStyleEffect(effectRef);
        }
        return style;
    }
    /**
     * Resolves an a:fillRef into a fill using the theme style matrix.
     * Per ECMA-376 §20.1.4.2.10: idx 0 or 1000 means no fill, 1-999 indexes
     * fillStyleLst (1-based), and 1001+ indexes bgFillStyleLst (1001-based).
     * The ref's child color substitutes phClr placeholders.
     */
    resolveStyleFill(fillRef) {
        const idx = parseInt((0, PptxParser_js_1.getXmlAttr)(fillRef, 'idx') ?? '0', 10);
        const phClr = this.colorResolver.resolveColorElement(fillRef);
        if (!Number.isFinite(idx) || idx <= 0 || idx === 1000) {
            return undefined;
        }
        const formatScheme = this.theme.formatScheme;
        const themeFill = idx < 1000
            ? formatScheme?.fillStyles[idx - 1]
            : formatScheme?.backgroundFillStyles[idx - 1001];
        if (!themeFill) {
            // No parsed style matrix (or out-of-range index): degrade to a solid
            // fill of the reference color so the shape still renders
            this.logger.debug('Style fill index not found in theme, using ref color', { idx });
            return phClr ? { type: 'solid', color: phClr } : undefined;
        }
        return this.fillRenderer.resolveThemeFill(themeFill, phClr);
    }
    /**
     * Resolves an a:lnRef into a stroke using the theme lnStyleLst
     * (ECMA-376 §20.1.4.2.19; 1-based index, 0 means no line).
     * The ref's child color substitutes phClr placeholders.
     */
    resolveStyleStroke(lnRef) {
        const idx = parseInt((0, PptxParser_js_1.getXmlAttr)(lnRef, 'idx') ?? '0', 10);
        const phClr = this.colorResolver.resolveColorElement(lnRef);
        if (!Number.isFinite(idx) || idx <= 0) {
            return undefined;
        }
        const lineStyle = this.theme.formatScheme?.lineStyles[idx - 1];
        if (!lineStyle) {
            // No parsed style matrix (or out-of-range index): degrade to a default
            // width line in the reference color
            this.logger.debug('Style line index not found in theme, using ref color', { idx });
            return phClr ? { width: StrokeRenderer_js_1.DEFAULT_STROKE_WIDTH_EMU, color: phClr } : undefined;
        }
        return this.strokeRenderer.resolveThemeLineStyle(lineStyle, phClr);
    }
    /**
     * Resolves an a:effectRef into an outer shadow using the theme
     * effectStyleLst (ECMA-376 §20.1.4.2.7; 1-based index, 0 means no
     * effect). The ref's child color substitutes phClr placeholders in the
     * shadow color. Only outer shadows are resolved; other themed effects
     * are not rendered.
     */
    resolveStyleEffect(effectRef) {
        const idx = parseInt((0, PptxParser_js_1.getXmlAttr)(effectRef, 'idx') ?? '0', 10);
        if (!Number.isFinite(idx) || idx <= 0) {
            return undefined;
        }
        const effectStyle = this.theme.formatScheme?.effectStyles[idx - 1];
        if (!effectStyle?.outerShadow) {
            this.logger.debug('Style effect index has no outer shadow', { idx });
            return undefined;
        }
        const phClr = this.colorResolver.resolveColorElement(effectRef);
        return this.effectParser.resolveThemeOuterShadow(effectStyle.outerShadow, phClr);
    }
    /**
     * Resolves the effective effects for a shape. Per ECMA-376 an explicit
     * spPr a:effectLst (even an empty one) overrides the style's a:effectRef.
     */
    resolveShapeEffects(spPr, style) {
        const explicit = this.effectParser.parseShapeEffects(spPr);
        if (explicit) {
            return explicit;
        }
        return style?.outerShadow ? { outerShadow: style.outerShadow } : undefined;
    }
    /**
     * Enables canvas shadow state for a plain outer shadow (no scale/skew;
     * perspective shadows go through EffectRenderer instead). Offsets and
     * blur are converted EMU -> pixels at the renderer's scale. Canvas
     * shadow offsets live in device space (unaffected by the CTM), which
     * matches rotWithShape="0"; rotated shapes with rotWithShape="1" keep a
     * device-space shadow — an accepted approximation.
     */
    setOuterShadow(ctx, shadow) {
        const { dx, dy } = (0, EffectParser_js_1.computeShadowOffset)(shadow.distance, shadow.direction);
        ctx.shadowColor = this.colorResolver.rgbaToCss(shadow.color);
        ctx.shadowBlur = (0, UnitConverter_js_1.emuToPixels)(shadow.blurRadius) * Math.min(this.scaleX, this.scaleY);
        ctx.shadowOffsetX = (0, UnitConverter_js_1.emuToPixels)(dx) * this.scaleX;
        ctx.shadowOffsetY = (0, UnitConverter_js_1.emuToPixels)(dy) * this.scaleY;
    }
    /**
     * Resets canvas shadow state so later passes (strokes, text) do not
     * inherit the shadow.
     */
    clearShadow(ctx) {
        ctx.shadowColor = 'rgba(0, 0, 0, 0)';
        ctx.shadowBlur = 0;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 0;
    }
    /**
     * Parses geometry information from shape properties.
     */
    parseGeometry(spPr) {
        // Check for preset geometry
        const prstGeom = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:prstGeom');
        if (prstGeom) {
            const prst = (0, PptxParser_js_1.getXmlAttr)(prstGeom, 'prst') ?? 'rect';
            const adjustValues = this.parseAdjustValues(prstGeom);
            return { geometryType: prst, adjustValues };
        }
        // Check for custom geometry
        const custGeom = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:custGeom');
        if (custGeom) {
            const def = (0, CustomGeometryParser_js_1.parseCustomGeometryDefinition)(custGeom);
            if (def) {
                return {
                    geometryType: 'custGeom',
                    customGeometry: { def, space: this.getGuideSpace(spPr) },
                };
            }
            this.logger.warn('Custom geometry could not be parsed, using rectangle');
            this.warnings?.push({
                code: 'unsupported-geometry',
                message: 'Custom geometry could not be parsed; shape rendered as rectangle',
                detail: { geometryType: 'custGeom' },
            });
            return { geometryType: 'rect' };
        }
        // Default to rectangle
        return { geometryType: 'rect' };
    }
    /**
     * Reads the shape's EMU extents from a:xfrm to use as the custom-geometry
     * guide coordinate space (custGeom coordinates are absolute EMU values).
     */
    getGuideSpace(spPr) {
        const xfrm = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:xfrm');
        const ext = xfrm ? (0, PptxParser_js_1.getXmlChild)(xfrm, 'a:ext') : undefined;
        if (!ext)
            return undefined;
        const w = parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cx') ?? '0', 10);
        const h = parseInt((0, PptxParser_js_1.getXmlAttr)(ext, 'cy') ?? '0', 10);
        if (!isFinite(w) || !isFinite(h) || (w <= 0 && h <= 0))
            return undefined;
        return { w, h };
    }
    /**
     * Parses adjustment values from preset geometry.
     */
    parseAdjustValues(prstGeom) {
        const avLst = (0, PptxParser_js_1.getXmlChild)(prstGeom, 'a:avLst');
        if (!avLst)
            return undefined;
        const gdNodes = (0, PptxParser_js_1.getXmlChildren)(avLst, 'a:gd');
        if (gdNodes.length === 0)
            return undefined;
        const adjustValues = new Map();
        for (const gd of gdNodes) {
            const name = (0, PptxParser_js_1.getXmlAttr)(gd, 'name');
            const fmla = (0, PptxParser_js_1.getXmlAttr)(gd, 'fmla');
            if (name && fmla) {
                // Parse "val X" formula. Guide values may be negative (e.g. arc/pie
                // start angles, callout offsets) and may carry a decimal fraction.
                const match = fmla.match(/^val\s+([-+]?\d+(?:\.\d+)?)$/);
                if (match?.[1]) {
                    adjustValues.set(name, parseFloat(match[1]));
                }
            }
        }
        return adjustValues.size > 0 ? adjustValues : undefined;
    }
    /**
     * Looks up a placeholder's transform from the slide layout or master.
     * @param placeholderType The placeholder type (title, body, etc.)
     * @param placeholderIdx The placeholder index
     * @returns The parsed transform or undefined if not found
     */
    getPlaceholderTransform(placeholderType, placeholderIdx) {
        // Try layout first, then master
        const sources = [this.layoutNode, this.masterNode].filter(Boolean);
        for (const sourceNode of sources) {
            if (!sourceNode)
                continue;
            // Navigate to shape tree in layout/master
            // Layout/master structure: p:sldLayout/p:sldMaster > p:cSld > p:spTree
            const cSld = (0, PptxParser_js_1.getXmlChild)(sourceNode, 'p:cSld');
            if (!cSld)
                continue;
            const spTree = (0, PptxParser_js_1.getXmlChild)(cSld, 'p:spTree');
            if (!spTree)
                continue;
            // Find matching placeholder shape
            const shapes = (0, PptxParser_js_1.getXmlChildren)(spTree, 'p:sp');
            for (const sp of shapes) {
                const nvSpPr = (0, PptxParser_js_1.getXmlChild)(sp, 'p:nvSpPr');
                if (!nvSpPr)
                    continue;
                const nvPr = (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr');
                if (!nvPr)
                    continue;
                const ph = (0, PptxParser_js_1.getXmlChild)(nvPr, 'p:ph');
                if (!ph)
                    continue;
                // Check if this placeholder matches
                const layoutPhType = (0, PptxParser_js_1.getXmlAttr)(ph, 'type');
                const layoutPhIdx = (0, PptxParser_js_1.getXmlAttr)(ph, 'idx');
                // Match by type (for title, body, etc.) or by idx
                const typeMatch = placeholderType && layoutPhType === placeholderType;
                const idxMatch = placeholderIdx !== undefined && layoutPhIdx === placeholderIdx;
                if (typeMatch || idxMatch) {
                    // Found matching placeholder - get its transform
                    const spPr = (0, PptxParser_js_1.getXmlChild)(sp, 'p:spPr');
                    if (spPr) {
                        const transform = this.transformCalculator.parseTransform(spPr);
                        if (transform) {
                            return transform;
                        }
                    }
                }
            }
        }
        return undefined;
    }
    /**
     * Creates the paths for a connector's geometry (straightConnector1,
     * bentConnector2..5, curvedConnector2..5, ...), falling back to a
     * straight line for unknown geometry.
     */
    createConnectorPaths(geometryType, bounds, adjustValues) {
        if (geometryType !== 'custGeom') {
            const paths = this.geometryCalculator.createPaths(geometryType, bounds, adjustValues);
            if (paths && paths.length > 0) {
                return paths;
            }
        }
        const line = this.geometryCalculator.createPath('line', bounds);
        return line ? [line] : [];
    }
    /**
     * Renders a connection shape (connector line).
     * @param ctx Canvas 2D context
     * @param cxnSpNode Connection shape XML node (p:cxnSp)
     */
    renderConnectionShape(ctx, cxnSpNode) {
        // Connection shapes are simpler - they're typically lines
        const nvCxnSpPr = (0, PptxParser_js_1.getXmlChild)(cxnSpNode, 'p:nvCxnSpPr');
        const cNvPr = nvCxnSpPr ? (0, PptxParser_js_1.getXmlChild)(nvCxnSpPr, 'p:cNvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden)
            return;
        const spPr = (0, PptxParser_js_1.getXmlChild)(cxnSpNode, 'p:spPr');
        if (!spPr)
            return;
        // Parse transform
        const transform = this.transformCalculator.parseTransform(spPr);
        if (!transform)
            return;
        const pixelTransform = this.transformCalculator.toPixelTransform(transform, this.scaleX, this.scaleY);
        // Parse geometry (straightConnector1, bentConnector3, ...)
        const { geometryType, adjustValues } = this.parseGeometry(spPr);
        // Create bounds
        const bounds = {
            x: 0,
            y: 0,
            width: pixelTransform.width,
            height: pixelTransform.height,
        };
        // Create the connector path (falls back to a straight line)
        const paths = this.createConnectorPaths(geometryType, bounds, adjustValues);
        if (paths.length === 0)
            return;
        // Parse stroke - explicit a:ln properties beat the style ref (a:lnRef)
        const style = this.parseShapeStyle(cxnSpNode);
        const stroke = this.strokeRenderer.parseStroke(spPr, style?.stroke);
        if (!stroke)
            return;
        // Parse effects - explicit spPr a:effectLst beats the style ref
        const effects = this.resolveShapeEffects(spPr, style);
        this.logger.debug('Rendering connection shape', { id, geometry: geometryType });
        // Apply transform and render stroke only (no fill for connectors)
        ctx.save();
        this.transformCalculator.applyTransform(ctx, pixelTransform);
        this.renderConnectorPasses(ctx, paths, bounds, stroke, effects);
        ctx.restore();
    }
    /**
     * Renders a connector's stroke passes with its effects: silhouette
     * outer shadow for perspective parameters (else the fast ctx.shadow*
     * path on the first stroked path only), and glow beneath the strokes.
     * Inner shadow, soft edge, and reflection are not applied to connectors.
     */
    renderConnectorPasses(ctx, paths, bounds, stroke, effects) {
        const outerShadow = effects?.outerShadow;
        // Connectors are stroke-only: the silhouette must come from the outline
        const source = this.createSilhouetteSource(paths, bounds, stroke, undefined);
        const silhouetteShadow = !!outerShadow && this.effectRenderer.needsSilhouetteShadow(outerShadow);
        if (outerShadow && silhouetteShadow) {
            this.effectRenderer.renderOuterShadow(ctx, source, outerShadow);
        }
        if (effects?.glow) {
            this.effectRenderer.renderGlow(ctx, source, effects.glow);
        }
        let pendingShadow = silhouetteShadow ? undefined : outerShadow;
        for (const path of paths) {
            const shadow = pendingShadow;
            if (shadow) {
                this.setOuterShadow(ctx, shadow);
            }
            this.strokeRenderer.renderStroke(ctx, path, stroke, this.scaleX, this.scaleY);
            if (shadow) {
                this.clearShadow(ctx);
                pendingShadow = undefined;
            }
        }
    }
    /**
     * Gets the image renderer for external use.
     * Note: This is a getter because imageRenderer can be null and may be recreated.
     */
    getImageRenderer() {
        return this.imageRenderer;
    }
    /**
     * Sets a new source path for image rendering.
     * Creates a new ImageRenderer for the given path.
     */
    setSourcePath(sourcePath) {
        if (this.parser) {
            this.imageRenderer = new ImageRenderer_js_1.ImageRenderer({
                parser: this.parser,
                sourcePath,
                scaleX: this.scaleX,
                scaleY: this.scaleY,
                logger: this.logger.child?.('ImageRenderer'),
                warnings: this.warnings,
            });
        }
    }
    /**
     * Renders a picture element (p:pic) to the canvas.
     * @param ctx Canvas 2D context
     * @param picNode Picture XML node (p:pic)
     */
    async renderPicture(ctx, picNode) {
        if (!this.imageRenderer) {
            this.logger.warn('Cannot render picture: ImageRenderer not initialized (missing parser or sourcePath)');
            return;
        }
        // Get non-visual properties
        const nvPicPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:nvPicPr');
        const cNvPr = nvPicPr ? (0, PptxParser_js_1.getXmlChild)(nvPicPr, 'p:cNvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden) {
            this.logger.debug('Picture is hidden, skipping', { id });
            return;
        }
        // Get shape properties for transform
        const spPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:spPr');
        if (!spPr) {
            this.logger.debug('Picture has no spPr, skipping', { id });
            return;
        }
        // Parse transform
        const transform = this.transformCalculator.parseTransform(spPr);
        if (!transform) {
            this.logger.debug('Picture has no transform, skipping', { id });
            return;
        }
        const pixelTransform = this.transformCalculator.toPixelTransform(transform, this.scaleX, this.scaleY);
        this.logger.debug('Rendering picture', {
            id,
            name,
            width: pixelTransform.width,
            height: pixelTransform.height,
        });
        const outerShadow = this.getPictureOuterShadow(picNode, spPr);
        try {
            if (outerShadow) {
                this.setOuterShadow(ctx, outerShadow);
            }
            // Render the picture using ImageRenderer
            await this.imageRenderer.renderPictureElement(ctx, picNode, pixelTransform);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to render picture', { id, error: message });
        }
        finally {
            if (outerShadow) {
                this.clearShadow(ctx);
            }
        }
    }
    /**
     * Resolves the outer shadow to apply around a picture draw. Tiled blip
     * fills draw many tiles inside a clip — a per-tile shadow would darken
     * seams — so tiled pictures render shadowless until ImageRenderer
     * supports silhouette shadows.
     */
    getPictureOuterShadow(picNode, spPr) {
        const blipFill = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:blipFill');
        if (blipFill && typeof blipFill === 'object' && (0, PptxParser_js_1.hasXmlChild)(blipFill, 'a:tile')) {
            return undefined;
        }
        const style = this.parseShapeStyle(picNode);
        return this.resolveShapeEffects(spPr, style)?.outerShadow;
    }
    /**
     * Parses a picture element to get its picture data.
     * Used for extracting picture information without rendering.
     */
    parsePicture(picNode) {
        // Get non-visual properties
        const nvPicPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:nvPicPr');
        const cNvPr = nvPicPr ? (0, PptxParser_js_1.getXmlChild)(nvPicPr, 'p:cNvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        // Get shape properties for transform
        const spPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:spPr');
        let transform;
        if (spPr) {
            const parsedTransform = this.transformCalculator.parseTransform(spPr);
            if (parsedTransform) {
                transform = this.transformCalculator.toPixelTransform(parsedTransform, this.scaleX, this.scaleY);
            }
        }
        // Get blip relationship ID
        const blipFill = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:blipFill');
        const blip = blipFill ? (0, PptxParser_js_1.getXmlChild)(blipFill, 'a:blip') : undefined;
        const blipRelId = blip ? (0, PptxParser_js_1.getXmlAttr)(blip, 'r:embed') : undefined;
        // Parse effects - an explicit spPr a:effectLst beats the style ref
        const effects = this.resolveShapeEffects(spPr, this.parseShapeStyle(picNode));
        return { id, name, transform, blipRelId, effects };
    }
    /**
     * Renders a shape element with a pre-computed pixel transform.
     * Used for rendering shapes within groups where the transform has already been calculated.
     * @param ctx Canvas 2D context
     * @param spNode Shape XML node (p:sp)
     * @param pixelTransform Pre-computed pixel transform
     */
    renderShapeWithTransform(ctx, spNode, pixelTransform) {
        // Get non-visual properties for visibility check
        const nvSpPr = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:nvSpPr');
        const cNvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:cNvPr') : undefined;
        const nvPr = nvSpPr ? (0, PptxParser_js_1.getXmlChild)(nvSpPr, 'p:nvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden)
            return;
        // Check if this is a placeholder shape (a bare <p:ph/> is valid)
        const phNode = this.getPresentChild(nvPr, 'p:ph');
        const placeholderType = phNode ? (0, PptxParser_js_1.getXmlAttr)(phNode, 'type') : undefined;
        const placeholderIdx = phNode ? (0, PptxParser_js_1.getXmlAttr)(phNode, 'idx') : undefined;
        // Get shape properties for geometry, fill, stroke
        const spPr = this.getSpPrNode(spNode);
        if (!spPr) {
            this.logger.debug('Shape has no spPr, skipping', { id });
            return;
        }
        // Parse geometry
        const { geometryType, adjustValues, customGeometry } = this.parseGeometry(spPr);
        // Resolve the shape style references (p:style)
        const style = this.parseShapeStyle(spNode);
        // Parse fill - explicit spPr fill (including a:noFill) beats the style ref
        const fill = this.fillRenderer.parseFill(spPr) ?? style?.fill;
        // Parse stroke - explicit a:ln properties beat the style ref
        const stroke = this.strokeRenderer.parseStroke(spPr, style?.stroke);
        // Parse effects - an explicit spPr a:effectLst beats the style ref
        const effects = this.resolveShapeEffects(spPr, style);
        // Parse text body, merging the placeholder text-style inheritance chain
        // (text boxes still inherit the presentation p:defaultTextStyle)
        const txBody = (0, PptxParser_js_1.getXmlChild)(spNode, 'p:txBody');
        const inheritedListStyles = phNode
            ? this.getInheritedListStyles(placeholderType, placeholderIdx)
            : this.getTextBoxListStyles();
        const textBody = this.textParser.parseTextBody(txBody, inheritedListStyles);
        this.logger.debug('Rendering shape with transform', {
            id,
            name,
            geometry: geometryType,
            hasText: !!textBody,
        });
        // Create the geometry path
        const bounds = {
            x: 0,
            y: 0,
            width: pixelTransform.width,
            height: pixelTransform.height,
        };
        const paths = this.createShapePaths(geometryType, bounds, adjustValues, customGeometry);
        if (paths.length === 0) {
            return;
        }
        // Apply transform and render
        ctx.save();
        this.transformCalculator.applyTransform(ctx, pixelTransform);
        // Get fill color for text contrast calculation
        let fillColor;
        if (fill?.type === 'solid' && paths.some((p) => p.fill !== false)) {
            fillColor = fill.color;
        }
        // Group rendering is synchronous: picture fills degrade as before, and
        // the sync content callbacks keep effect compositing in draw order
        void this.renderEffectsAndBody(ctx, paths, bounds, fill, stroke, effects, false);
        // Finally render text if present
        if (textBody && textBody.paragraphs.length > 0) {
            // Default text color: the style fontRef color when the shape carries a
            // p:style, otherwise fall back to the fill-contrast heuristic
            const defaultTextColor = style?.fontColor ??
                (fillColor ? this.textRenderer.getContrastingColor(fillColor) : undefined);
            // Get text bounds (may be different from shape bounds for non-rectangular shapes)
            const textBounds = this.getShapeTextBounds(geometryType, bounds, adjustValues, customGeometry);
            this.textRenderer.renderText(ctx, textBody, textBounds, defaultTextColor);
        }
        ctx.restore();
    }
    /**
     * Renders a connection shape with a pre-computed pixel transform.
     * Used for rendering connection shapes within groups.
     * @param ctx Canvas 2D context
     * @param cxnSpNode Connection shape XML node (p:cxnSp)
     * @param pixelTransform Pre-computed pixel transform
     */
    renderConnectionShapeWithTransform(ctx, cxnSpNode, pixelTransform) {
        // Get non-visual properties for visibility check
        const nvCxnSpPr = (0, PptxParser_js_1.getXmlChild)(cxnSpNode, 'p:nvCxnSpPr');
        const cNvPr = nvCxnSpPr ? (0, PptxParser_js_1.getXmlChild)(nvCxnSpPr, 'p:cNvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden)
            return;
        const spPr = (0, PptxParser_js_1.getXmlChild)(cxnSpNode, 'p:spPr');
        if (!spPr)
            return;
        // Parse geometry
        const { geometryType, adjustValues } = this.parseGeometry(spPr);
        // Create bounds
        const bounds = {
            x: 0,
            y: 0,
            width: pixelTransform.width,
            height: pixelTransform.height,
        };
        // Create the connector path (falls back to a straight line)
        const paths = this.createConnectorPaths(geometryType, bounds, adjustValues);
        if (paths.length === 0)
            return;
        // Parse stroke - explicit a:ln properties beat the style ref (a:lnRef)
        const style = this.parseShapeStyle(cxnSpNode);
        const stroke = this.strokeRenderer.parseStroke(spPr, style?.stroke);
        if (!stroke)
            return;
        // Parse effects - explicit spPr a:effectLst beats the style ref
        const effects = this.resolveShapeEffects(spPr, style);
        this.logger.debug('Rendering connection shape with transform', { id, geometry: geometryType });
        // Apply transform and render stroke only
        ctx.save();
        this.transformCalculator.applyTransform(ctx, pixelTransform);
        this.renderConnectorPasses(ctx, paths, bounds, stroke, effects);
        ctx.restore();
    }
    /**
     * Renders a picture element with a pre-computed pixel transform.
     * Used for rendering pictures within groups.
     * @param ctx Canvas 2D context
     * @param picNode Picture XML node (p:pic)
     * @param pixelTransform Pre-computed pixel transform
     */
    async renderPictureWithTransform(ctx, picNode, pixelTransform) {
        if (!this.imageRenderer) {
            this.logger.warn('Cannot render picture: ImageRenderer not initialized');
            return;
        }
        // Get non-visual properties
        const nvPicPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:nvPicPr');
        const cNvPr = nvPicPr ? (0, PptxParser_js_1.getXmlChild)(nvPicPr, 'p:cNvPr') : undefined;
        const id = cNvPr ? ((0, PptxParser_js_1.getXmlAttr)(cNvPr, 'id') ?? '0') : '0';
        const name = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'name') : undefined;
        const hidden = cNvPr ? (0, PptxParser_js_1.getXmlAttr)(cNvPr, 'hidden') === '1' : false;
        if (hidden) {
            this.logger.debug('Picture is hidden, skipping', { id });
            return;
        }
        this.logger.debug('Rendering picture with transform', {
            id,
            name,
            width: pixelTransform.width,
            height: pixelTransform.height,
        });
        const spPr = (0, PptxParser_js_1.getXmlChild)(picNode, 'p:spPr');
        const outerShadow = this.getPictureOuterShadow(picNode, spPr);
        try {
            if (outerShadow) {
                this.setOuterShadow(ctx, outerShadow);
            }
            // Render the picture using ImageRenderer with the pre-computed transform
            await this.imageRenderer.renderPictureElement(ctx, picNode, pixelTransform);
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.error('Failed to render picture', { id, error: message });
        }
        finally {
            if (outerShadow) {
                this.clearShadow(ctx);
            }
        }
    }
}
exports.ShapeRenderer = ShapeRenderer;
/**
 * Default shape renderer factory.
 */
function createShapeRenderer(theme, scaleX, scaleY, logger, parser, sourcePath) {
    return new ShapeRenderer({ theme, scaleX, scaleY, logger, parser, sourcePath });
}
