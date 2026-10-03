"use strict";
/**
 * Main text rendering to canvas.
 * Uses TextLayoutEngine for positioning and renders text with proper styling.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.TextRenderer = void 0;
exports.createTextRenderer = createTextRenderer;
const TextLayoutEngine_js_1 = require("../text/TextLayoutEngine.js");
const ColorResolver_js_1 = require("../theme/ColorResolver.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Renders text to canvas.
 */
class TextRenderer {
    logger;
    theme;
    scaleX;
    scaleY;
    colorResolver;
    layoutEngine;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'TextRenderer');
        this.theme = config.theme;
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
        this.colorResolver = new ColorResolver_js_1.ColorResolver(config.theme.colors);
        this.layoutEngine = new TextLayoutEngine_js_1.TextLayoutEngine({
            fontScheme: config.theme.fonts,
            logger: this.logger.child?.('Layout'),
            warnings: config.warnings,
        });
    }
    /**
     * Renders text body within shape bounds.
     *
     * @param ctx Canvas 2D context
     * @param textBody Text body to render
     * @param shapeBounds Shape bounds in pixels (after transform applied)
     * @param defaultColor Default text color (from shape fill contrast)
     */
    renderText(ctx, textBody, shapeBounds, defaultColor) {
        // Layout the text
        const layout = this.layoutEngine.layoutText(ctx, textBody, shapeBounds, this.scaleX, this.scaleY);
        // Check if we have text rotation
        const rotation = textBody.bodyProperties?.rotation;
        if (rotation) {
            this.renderRotatedText(ctx, layout, shapeBounds, defaultColor, rotation);
        }
        else {
            this.renderLayoutLines(ctx, layout.lines, defaultColor);
        }
        this.logger.debug('Rendered text', {
            lineCount: layout.lines.length,
            totalHeight: layout.totalHeight,
        });
    }
    /**
     * Renders text with rotation.
     */
    renderRotatedText(ctx, layout, shapeBounds, defaultColor, rotationDegrees) {
        ctx.save();
        // Rotate around center of text bounds
        const centerX = layout.bounds.x + layout.bounds.width / 2;
        const centerY = layout.bounds.y + layout.bounds.height / 2;
        ctx.translate(centerX, centerY);
        ctx.rotate((rotationDegrees * Math.PI) / 180);
        ctx.translate(-centerX, -centerY);
        this.renderLayoutLines(ctx, layout.lines, defaultColor);
        ctx.restore();
    }
    /**
     * Renders all layout lines.
     */
    renderLayoutLines(ctx, lines, defaultColor) {
        for (const line of lines) {
            // Render bullet if present
            if (line.bullet) {
                this.renderBullet(ctx, line.bullet, defaultColor);
            }
            // Render text runs
            for (const run of line.runs) {
                this.renderTextRun(ctx, run, defaultColor);
            }
        }
    }
    /**
     * Renders a single text run.
     */
    renderTextRun(ctx, run, defaultColor) {
        if (!run.text)
            return;
        ctx.save();
        // Set font
        ctx.font = run.font.fontString;
        // Letter tracking (a:rPr spc): skia-canvas applies ctx.letterSpacing
        // natively between glyphs, matching the spacing-aware measurement used
        // during layout. Scoped by the surrounding save/restore. Unkerned
        // (noKern) runs place glyphs individually instead, so tracking is part
        // of each manual advance rather than ctx state.
        if (run.font.letterSpacingPx && !run.font.noKern) {
            ctx.letterSpacing = `${run.font.letterSpacingPx}px`;
        }
        // Set fill color
        const color = run.color ?? defaultColor ?? { r: 0, g: 0, b: 0, a: 255 };
        ctx.fillStyle = this.colorResolver.rgbaToCss(color);
        // Handle baseline offset (super/subscript)
        let y = run.y;
        let baselineScale = 1;
        if (run.baselineOffset) {
            // Baseline offset is in 1000ths of percentage
            // Positive = superscript (move up), Negative = subscript (move down)
            const offsetPercent = run.baselineOffset / 1000;
            // sizePx already includes the render scale
            const offsetPixels = (run.font.sizePx * offsetPercent) / 100;
            y -= offsetPixels;
            // Also scale font for super/subscript
            if (Math.abs(offsetPercent) > 20) {
                const scaledSize = run.font.sizePx * 0.6;
                const scaledFont = run.font.fontString.replace(`${run.font.sizePx}px`, `${scaledSize}px`);
                ctx.font = scaledFont;
                baselineScale = 0.6;
            }
        }
        // Draw text. Unkerned runs (a:rPr kern threshold not met) place each
        // glyph at its bare cumulative advance — skia's shaped fillText always
        // kerns, which condenses small print PowerPoint lays out unkerned.
        // Advances scale with any super/subscript font scaling so glyph
        // placement matches the drawn size.
        if (run.font.noKern) {
            const advances = this.layoutEngine.getFontResolver().charAdvances(ctx, run.text, run.font);
            let x = run.x;
            let i = 0;
            for (const char of run.text) {
                ctx.fillText(char, x, y);
                x += advances[i++] * baselineScale;
            }
        }
        else {
            ctx.fillText(run.text, run.x, y);
        }
        // Draw underline
        if (run.underline) {
            this.drawUnderline(ctx, run, color);
        }
        // Draw strikethrough
        if (run.strikethrough) {
            this.drawStrikethrough(ctx, run, color);
        }
        ctx.restore();
    }
    /**
     * Renders a bullet.
     */
    renderBullet(ctx, bullet, defaultColor) {
        if (!bullet.text)
            return;
        ctx.save();
        // Set font
        ctx.font = bullet.font.fontString;
        // Set fill color (bullet color or default)
        const color = bullet.color ?? defaultColor ?? { r: 0, g: 0, b: 0, a: 255 };
        ctx.fillStyle = this.colorResolver.rgbaToCss(color);
        // Draw bullet
        ctx.fillText(bullet.text, bullet.x, bullet.y);
        ctx.restore();
    }
    /**
     * Draws underline decoration.
     */
    drawUnderline(ctx, run, color) {
        // Use pre-computed width from layout to avoid re-measuring.
        // sizePx already includes the render scale.
        const lineWidth = Math.max(1, run.font.sizePx * 0.05);
        const offset = run.font.sizePx * 0.1;
        ctx.strokeStyle = this.colorResolver.rgbaToCss(color);
        ctx.lineWidth = lineWidth;
        ctx.beginPath();
        ctx.moveTo(run.x, run.y + offset);
        ctx.lineTo(run.x + run.width, run.y + offset);
        ctx.stroke();
    }
    /**
     * Draws strikethrough decoration.
     */
    drawStrikethrough(ctx, run, color) {
        // Use pre-computed width from layout to avoid re-measuring.
        // sizePx already includes the render scale.
        const lineWidth = Math.max(1, run.font.sizePx * 0.05);
        const offset = -run.font.sizePx * 0.3;
        ctx.strokeStyle = this.colorResolver.rgbaToCss(color);
        ctx.lineWidth = lineWidth;
        ctx.beginPath();
        ctx.moveTo(run.x, run.y + offset);
        ctx.lineTo(run.x + run.width, run.y + offset);
        ctx.stroke();
    }
    /**
     * Gets the layout engine for external use.
     */
    getLayoutEngine() {
        return this.layoutEngine;
    }
    /**
     * Gets the color resolver for external use.
     */
    getColorResolver() {
        return this.colorResolver;
    }
    /**
     * Calculates a contrasting text color based on background.
     *
     * @param backgroundColor Background color to contrast against
     * @returns Black or white, whichever provides better contrast
     */
    getContrastingColor(backgroundColor) {
        const isDark = this.colorResolver.isDarkColor(backgroundColor);
        return isDark ? { r: 255, g: 255, b: 255, a: 255 } : { r: 0, g: 0, b: 0, a: 255 };
    }
}
exports.TextRenderer = TextRenderer;
/**
 * Creates a TextRenderer instance.
 */
function createTextRenderer(theme, scaleX, scaleY, logger) {
    return new TextRenderer({ theme, scaleX, scaleY, logger });
}
