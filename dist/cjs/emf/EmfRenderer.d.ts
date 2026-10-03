/**
 * Replays a parsed EMF/WMF record stream onto a skia-canvas Canvas.
 *
 * Coordinate pipeline (GDI semantics): world transform -> page space
 * (map mode + window/viewport org/ext) -> device units. The canvas is
 * scaled once so the metafile frame (rclFrame) maps to the output size,
 * then all drawing happens in device units.
 */
import { Canvas } from 'skia-canvas';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
import type { ParsedMetafile } from './types.js';
/**
 * Configuration for EmfRenderer.
 */
export interface EmfRendererConfig {
    /** Logger instance */
    logger?: ILogger;
    /** Maximum output dimension in pixels (default 4096) */
    maxDimensionPx?: number;
    /** Additional resolution multiplier applied to the natural size (default 1) */
    scale?: number;
    /** Structured warning collector for partial-render fidelity events */
    warnings?: WarningCollector;
}
/**
 * Renders parsed metafiles to a skia-canvas Canvas.
 */
export declare class EmfRenderer {
    private readonly logger;
    private readonly maxDimensionPx;
    private readonly scale;
    private readonly warnings;
    /**
     * Resolves metafile face names to canvas font family chains with the same
     * policy as regular text: registered/system availability checks, session
     * fonts.fallbacks, metric-compatible substitutes, and missing-font
     * warnings. Metafiles carry concrete face names only, so the theme scheme
     * here is just the resolver's required default and is never consulted.
     */
    private readonly fontResolver;
    constructor(config?: EmfRendererConfig);
    /**
     * Renders a parsed metafile to a canvas at its natural device resolution
     * (the frame rectangle mapped through the recording device's pixel pitch),
     * capped at maxDimensionPx.
     *
     * @param parsed The parsed metafile
     * @returns The rendered canvas (transparent background)
     */
    render(parsed: ParsedMetafile): Promise<Canvas>;
    /**
     * Computes the source rectangle in device units that maps to the output
     * canvas: the frame (.01mm) scaled by the recording device's px/mm pitch,
     * falling back to the header bounds when the frame is degenerate.
     */
    private computeSourceRect;
    /** Device pixel pitch of the recording device (px per millimeter). */
    private deviceMetrics;
    /** GDI default DC state. */
    private initialState;
    /** Deep-copies a DC state for SaveDC. */
    private cloneState;
    /**
     * Applies a single record. Returns the (possibly replaced) DC state, or
     * null when EMR_EOF is reached.
     */
    private applyRecord;
    /** Applies ModifyWorldTransform per its mode. */
    private modifyWorld;
    /** Handles SelectObject for both created and stock objects. */
    private selectObject;
    /** Applies a stock object selection to the DC. */
    private selectStockObject;
    /** Page-space scale factors from the current map mode. */
    private pageScale;
    /** Transforms a logical point to device units (world -> page -> device). */
    private toDevice;
    /** Magnitudes of the logical->device scale along each axis. */
    private scaleMagnitudes;
    /** Runs a draw callback with the DC clip applied around it. */
    private withClip;
    /** Sets stroke style from the DC pen; returns false for PS_NULL. */
    private applyPen;
    /** Sets fill style from the DC brush; returns false for BS_NULL. */
    private applyBrush;
    /** Builds a path via callback, then fills (brush) and strokes (pen). */
    private fillStrokePath;
    /** Builds a path via callback, then strokes it with the DC pen. */
    private strokePath;
    /** Appends a rectangle (4 transformed corners) to the current path. */
    private rectPath;
    /** Appends a poly(gon/line) to the current path. */
    private polyPath;
    /**
     * Appends cubic bezier segments to the current path.
     * @param fromCurrent When true (PolyBezierTo), the start point is the DC
     * current position and all points are control/end points.
     */
    private bezierPath;
    /** Renders an ExtTextOut record. */
    private drawText;
    /** Renders a dibDraw record (BitBlt/StretchBlt/StretchDIBits). */
    private drawDib;
    /**
     * Decodes an embedded DIB. Uncompressed 24/32bpp DIBs are converted
     * directly; other formats are wrapped in a BMP file header and decoded
     * via skia-canvas.
     */
    private decodeDib;
    /** Converts an uncompressed 24/32bpp DIB to a canvas via ImageData. */
    private decodeUncompressedDib;
}
