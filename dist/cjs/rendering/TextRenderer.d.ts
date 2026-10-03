/**
 * Main text rendering to canvas.
 * Uses TextLayoutEngine for positioning and renders text with proper styling.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rect, Rgba } from '../types/geometry.js';
import type { TextBody } from '../types/elements.js';
import type { ResolvedTheme } from '../types/theme.js';
import { TextLayoutEngine } from '../text/TextLayoutEngine.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Configuration for TextRenderer.
 */
export interface TextRendererConfig {
    /** Resolved theme */
    theme: ResolvedTheme;
    /** Horizontal scale factor */
    scaleX: number;
    /** Vertical scale factor */
    scaleY: number;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector, forwarded to the text layout pipeline */
    warnings?: WarningCollector;
}
/**
 * Renders text to canvas.
 */
export declare class TextRenderer {
    private readonly logger;
    private readonly theme;
    private readonly scaleX;
    private readonly scaleY;
    private readonly colorResolver;
    private readonly layoutEngine;
    constructor(config: TextRendererConfig);
    /**
     * Renders text body within shape bounds.
     *
     * @param ctx Canvas 2D context
     * @param textBody Text body to render
     * @param shapeBounds Shape bounds in pixels (after transform applied)
     * @param defaultColor Default text color (from shape fill contrast)
     */
    renderText(ctx: CanvasRenderingContext2D, textBody: TextBody, shapeBounds: Rect, defaultColor?: Rgba): void;
    /**
     * Renders text with rotation.
     */
    private renderRotatedText;
    /**
     * Renders all layout lines.
     */
    private renderLayoutLines;
    /**
     * Renders a single text run.
     */
    private renderTextRun;
    /**
     * Renders a bullet.
     */
    private renderBullet;
    /**
     * Draws underline decoration.
     */
    private drawUnderline;
    /**
     * Draws strikethrough decoration.
     */
    private drawStrikethrough;
    /**
     * Gets the layout engine for external use.
     */
    getLayoutEngine(): TextLayoutEngine;
    /**
     * Gets the color resolver for external use.
     */
    getColorResolver(): ColorResolver;
    /**
     * Calculates a contrasting text color based on background.
     *
     * @param backgroundColor Background color to contrast against
     * @returns Black or white, whichever provides better contrast
     */
    getContrastingColor(backgroundColor: Rgba): Rgba;
}
/**
 * Creates a TextRenderer instance.
 */
export declare function createTextRenderer(theme: ResolvedTheme, scaleX: number, scaleY: number, logger?: ILogger): TextRenderer;
