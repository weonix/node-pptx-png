/**
 * Replays a parsed EMF/WMF record stream onto a skia-canvas Canvas.
 *
 * Coordinate pipeline (GDI semantics): world transform -> page space
 * (map mode + window/viewport org/ext) -> device units. The canvas is
 * scaled once so the metafile frame (rclFrame) maps to the output size,
 * then all drawing happens in device units.
 */
import { Canvas, loadImage } from 'skia-canvas';
import { createLogger } from '../utils/Logger.js';
import { FontResolver } from '../text/FontResolver.js';
import { BK_MODE, BRUSH_STYLE, colorRefToCss, ETO, MAP_MODE, MWT, PEN_STYLE, POLY_FILL_MODE, RGN_MODE, ROP, STOCK_OBJECT, STOCK_OBJECT_FLAG, TEXT_ALIGN, } from './types.js';
/** Default font used when the metafile selects a stock font. */
const DEFAULT_FONT = {
    height: -12,
    weight: 400,
    italic: false,
    underline: false,
    strikeOut: false,
    faceName: 'Arial',
};
/**
 * Renders parsed metafiles to a skia-canvas Canvas.
 */
export class EmfRenderer {
    logger;
    maxDimensionPx;
    scale;
    warnings;
    /**
     * Resolves metafile face names to canvas font family chains with the same
     * policy as regular text: registered/system availability checks, session
     * fonts.fallbacks, metric-compatible substitutes, and missing-font
     * warnings. Metafiles carry concrete face names only, so the theme scheme
     * here is just the resolver's required default and is never consulted.
     */
    fontResolver;
    constructor(config = {}) {
        this.logger = config.logger ?? createLogger('warn', 'EmfRenderer');
        this.maxDimensionPx = config.maxDimensionPx ?? 4096;
        this.scale = config.scale ?? 1;
        this.warnings = config.warnings;
        this.fontResolver = new FontResolver({
            fontScheme: { majorFont: 'Arial', minorFont: 'Arial' },
            logger: this.logger.child?.('FontResolver'),
            warnings: config.warnings,
        });
    }
    /**
     * Renders a parsed metafile to a canvas at its natural device resolution
     * (the frame rectangle mapped through the recording device's pixel pitch),
     * capped at maxDimensionPx.
     *
     * @param parsed The parsed metafile
     * @returns The rendered canvas (transparent background)
     */
    async render(parsed) {
        if (parsed.emfPlusDetected && !parsed.emfPlusDual) {
            this.logger.warn('EMF+ stream without dual GDI records; rendering the partial GDI subset only');
            this.warnings?.push({
                code: 'metafile-partial',
                message: 'EMF+ metafile without dual GDI records; only the partial GDI subset was rendered',
                detail: { source: parsed.source },
            });
        }
        if (parsed.truncated) {
            this.logger.warn('Metafile was truncated or malformed; rendering the parsed subset only');
            this.warnings?.push({
                code: 'metafile-partial',
                message: 'Metafile was truncated or malformed; only the parsed records were rendered',
                detail: { source: parsed.source, records: parsed.records.length },
            });
        }
        const src = this.computeSourceRect(parsed);
        let scale = this.scale > 0 ? this.scale : 1;
        const maxSrc = Math.max(src.width, src.height) * scale;
        if (maxSrc > this.maxDimensionPx) {
            scale *= this.maxDimensionPx / maxSrc;
        }
        const outWidth = Math.max(1, Math.round(src.width * scale));
        const outHeight = Math.max(1, Math.round(src.height * scale));
        const canvas = new Canvas(outWidth, outHeight);
        const ctx = canvas.getContext('2d');
        ctx.scale(outWidth / src.width, outHeight / src.height);
        ctx.translate(-src.left, -src.top);
        // Pre-decode embedded DIBs (async) so replay itself is synchronous
        const dibImages = new Map();
        for (const record of parsed.records) {
            if (record.type === 'dibDraw' && record.dib) {
                const decoded = await this.decodeDib(record.dib);
                if (decoded) {
                    dibImages.set(record, decoded);
                }
            }
        }
        const device = this.deviceMetrics(parsed);
        const state = this.initialState();
        const stack = [];
        const objects = new Map();
        let current = state;
        let replayFailureCollected = false;
        for (const record of parsed.records) {
            try {
                const next = this.applyRecord(ctx, current, stack, objects, dibImages, device, record);
                if (next === null) {
                    break; // eof
                }
                current = next;
            }
            catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                this.logger.warn('Failed to replay metafile record; skipping', {
                    type: record.type,
                    error: message,
                });
                // One structured warning per render is enough: a broken metafile
                // can fail on hundreds of records.
                if (!replayFailureCollected) {
                    replayFailureCollected = true;
                    this.warnings?.push({
                        code: 'metafile-partial',
                        message: `Some metafile records could not be replayed (first failure: ${record.type}: ${message})`,
                        detail: { source: parsed.source, recordType: record.type },
                    });
                }
            }
        }
        return canvas;
    }
    /**
     * Computes the source rectangle in device units that maps to the output
     * canvas: the frame (.01mm) scaled by the recording device's px/mm pitch,
     * falling back to the header bounds when the frame is degenerate.
     */
    computeSourceRect(parsed) {
        const { frame, bounds } = parsed.header;
        const { pxPerMmX, pxPerMmY } = this.deviceMetrics(parsed);
        const left = (frame.left / 100) * pxPerMmX;
        const top = (frame.top / 100) * pxPerMmY;
        const width = ((frame.right - frame.left) / 100) * pxPerMmX;
        const height = ((frame.bottom - frame.top) / 100) * pxPerMmY;
        if (width >= 1 && height >= 1) {
            return { left, top, width, height };
        }
        const boundsWidth = bounds.right - bounds.left + 1;
        const boundsHeight = bounds.bottom - bounds.top + 1;
        if (boundsWidth >= 1 && boundsHeight >= 1) {
            return { left: bounds.left, top: bounds.top, width: boundsWidth, height: boundsHeight };
        }
        return { left: 0, top: 0, width: 1, height: 1 };
    }
    /** Device pixel pitch of the recording device (px per millimeter). */
    deviceMetrics(parsed) {
        const { devicePx, deviceMm } = parsed.header;
        const fallback = 96 / 25.4;
        return {
            pxPerMmX: deviceMm.cx > 0 && devicePx.cx > 0 ? devicePx.cx / deviceMm.cx : fallback,
            pxPerMmY: deviceMm.cy > 0 && devicePx.cy > 0 ? devicePx.cy / deviceMm.cy : fallback,
        };
    }
    /** GDI default DC state. */
    initialState() {
        return {
            mapMode: MAP_MODE.TEXT,
            windowOrg: { x: 0, y: 0 },
            windowExt: { cx: 1, cy: 1 },
            viewportOrg: { x: 0, y: 0 },
            viewportExt: { cx: 1, cy: 1 },
            bkMode: BK_MODE.OPAQUE,
            bkColor: 0x00ffffff,
            textColor: 0x00000000,
            textAlign: 0,
            polyFillMode: POLY_FILL_MODE.ALTERNATE,
            pen: { style: PEN_STYLE.SOLID, width: 0, color: 0x00000000 },
            brush: { style: BRUSH_STYLE.SOLID, color: 0x00ffffff },
            font: { ...DEFAULT_FONT },
            clip: [],
            cur: { x: 0, y: 0 },
        };
    }
    /** Deep-copies a DC state for SaveDC. */
    cloneState(state) {
        return {
            ...state,
            windowOrg: { ...state.windowOrg },
            windowExt: { ...state.windowExt },
            viewportOrg: { ...state.viewportOrg },
            viewportExt: { ...state.viewportExt },
            world: state.world ? { ...state.world } : undefined,
            pen: { ...state.pen },
            brush: { ...state.brush },
            font: { ...state.font },
            clip: state.clip.map((step) => step.map((rect) => ({ ...rect }))),
            cur: { ...state.cur },
        };
    }
    /**
     * Applies a single record. Returns the (possibly replaced) DC state, or
     * null when EMR_EOF is reached.
     */
    applyRecord(ctx, state, stack, objects, dibImages, device, record) {
        switch (record.type) {
            case 'eof':
                return null;
            case 'setMapMode':
                state.mapMode = record.mode;
                break;
            case 'setWindowExt':
                state.windowExt = { cx: record.cx, cy: record.cy };
                break;
            case 'setWindowOrg':
                state.windowOrg = { x: record.x, y: record.y };
                break;
            case 'setViewportExt':
                state.viewportExt = { cx: record.cx, cy: record.cy };
                break;
            case 'setViewportOrg':
                state.viewportOrg = { x: record.x, y: record.y };
                break;
            case 'setWorldTransform':
                state.world = { ...record.xform };
                break;
            case 'modifyWorldTransform':
                this.modifyWorld(state, record.mode, record.xform);
                break;
            case 'saveDc':
                stack.push(this.cloneState(state));
                break;
            case 'restoreDc': {
                const target = record.saved < 0 ? stack.length + record.saved : record.saved - 1;
                if (target >= 0 && target < stack.length) {
                    const restored = stack[target];
                    stack.length = target;
                    if (restored) {
                        return restored;
                    }
                }
                break;
            }
            case 'setBkMode':
                state.bkMode = record.mode;
                break;
            case 'setBkColor':
                state.bkColor = record.color;
                break;
            case 'setTextColor':
                state.textColor = record.color;
                break;
            case 'setTextAlign':
                state.textAlign = record.align;
                break;
            case 'setRop2':
            case 'setStretchBltMode':
                // Stored implicitly; only default modes are emulated
                break;
            case 'setPolyFillMode':
                state.polyFillMode = record.mode;
                break;
            case 'intersectClipRect': {
                // Logical units: transform to device
                const p1 = this.toDevice(state, device, record.rect.left, record.rect.top);
                const p2 = this.toDevice(state, device, record.rect.right, record.rect.bottom);
                state.clip.push([normalizeRect(p1, p2)]);
                break;
            }
            case 'selectClipRgn':
                // Region data is already in device units per [MS-EMF]
                if (record.mode === RGN_MODE.COPY) {
                    state.clip = record.rects.length > 0 ? [record.rects.map((r) => ({ ...r }))] : [];
                }
                else if (record.mode === RGN_MODE.AND && record.rects.length > 0) {
                    state.clip.push(record.rects.map((r) => ({ ...r })));
                }
                else if (record.rects.length > 0) {
                    this.logger.debug('Unsupported clip region mode; ignoring', { mode: record.mode });
                }
                break;
            case 'createPen':
                objects.set(record.index, {
                    kind: 'pen',
                    pen: { style: record.style, width: record.width, color: record.color },
                });
                break;
            case 'createBrush':
                objects.set(record.index, {
                    kind: 'brush',
                    brush: { style: record.style, color: record.color },
                });
                break;
            case 'createFont':
                objects.set(record.index, { kind: 'font', font: { ...record.font } });
                break;
            case 'selectObject':
                this.selectObject(state, objects, record.index);
                break;
            case 'deleteObject':
                objects.delete(record.index);
                break;
            case 'moveTo':
                state.cur = { x: record.x, y: record.y };
                break;
            case 'lineTo':
                this.withClip(ctx, state, () => {
                    if (this.applyPen(ctx, state, device)) {
                        ctx.beginPath();
                        const from = this.toDevice(state, device, state.cur.x, state.cur.y);
                        const to = this.toDevice(state, device, record.x, record.y);
                        ctx.moveTo(from.x, from.y);
                        ctx.lineTo(to.x, to.y);
                        ctx.stroke();
                    }
                });
                state.cur = { x: record.x, y: record.y };
                break;
            case 'rectangle':
                this.fillStrokePath(ctx, state, device, () => {
                    this.rectPath(ctx, state, device, record.rect);
                });
                break;
            case 'ellipse':
                this.fillStrokePath(ctx, state, device, () => {
                    const p1 = this.toDevice(state, device, record.rect.left, record.rect.top);
                    const p2 = this.toDevice(state, device, record.rect.right, record.rect.bottom);
                    const rect = normalizeRect(p1, p2);
                    ctx.ellipse((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2, (rect.right - rect.left) / 2, (rect.bottom - rect.top) / 2, 0, 0, Math.PI * 2);
                });
                break;
            case 'roundRect':
                this.fillStrokePath(ctx, state, device, () => {
                    const p1 = this.toDevice(state, device, record.rect.left, record.rect.top);
                    const p2 = this.toDevice(state, device, record.rect.right, record.rect.bottom);
                    const rect = normalizeRect(p1, p2);
                    const scales = this.scaleMagnitudes(state, device);
                    const radius = Math.min(Math.max(Math.abs((record.cornerWidth / 2) * scales.sx), Math.abs((record.cornerHeight / 2) * scales.sy)), (rect.right - rect.left) / 2, (rect.bottom - rect.top) / 2);
                    ctx.moveTo(rect.left + radius, rect.top);
                    ctx.arcTo(rect.right, rect.top, rect.right, rect.bottom, radius);
                    ctx.arcTo(rect.right, rect.bottom, rect.left, rect.bottom, radius);
                    ctx.arcTo(rect.left, rect.bottom, rect.left, rect.top, radius);
                    ctx.arcTo(rect.left, rect.top, rect.right, rect.top, radius);
                    ctx.closePath();
                });
                break;
            case 'polygon':
                this.fillStrokePath(ctx, state, device, () => {
                    this.polyPath(ctx, state, device, record.points, true);
                });
                break;
            case 'polyPolygon':
                this.fillStrokePath(ctx, state, device, () => {
                    for (const poly of record.polygons) {
                        this.polyPath(ctx, state, device, poly, true);
                    }
                });
                break;
            case 'polyline':
                this.strokePath(ctx, state, device, () => {
                    this.polyPath(ctx, state, device, record.points, false);
                });
                break;
            case 'polyPolyline':
                this.strokePath(ctx, state, device, () => {
                    for (const poly of record.polylines) {
                        this.polyPath(ctx, state, device, poly, false);
                    }
                });
                break;
            case 'polylineTo': {
                this.strokePath(ctx, state, device, () => {
                    this.polyPath(ctx, state, device, [state.cur, ...record.points], false);
                });
                const last = record.points[record.points.length - 1];
                if (last) {
                    state.cur = { ...last };
                }
                break;
            }
            case 'polyBezier':
                this.strokePath(ctx, state, device, () => {
                    this.bezierPath(ctx, state, device, record.points, false);
                });
                break;
            case 'polyBezierTo': {
                this.strokePath(ctx, state, device, () => {
                    this.bezierPath(ctx, state, device, record.points, true);
                });
                const last = record.points[record.points.length - 1];
                if (last) {
                    state.cur = { ...last };
                }
                break;
            }
            case 'textOut':
                this.drawText(ctx, state, device, record);
                break;
            case 'dibDraw':
                this.drawDib(ctx, state, device, record, dibImages.get(record));
                break;
            default:
                break;
        }
        return state;
    }
    /** Applies ModifyWorldTransform per its mode. */
    modifyWorld(state, mode, xform) {
        if (mode === MWT.IDENTITY) {
            state.world = undefined;
            return;
        }
        const current = state.world ?? { m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 };
        if (mode === MWT.LEFTMULTIPLY) {
            state.world = multiplyXform(xform, current);
        }
        else if (mode === MWT.RIGHTMULTIPLY) {
            state.world = multiplyXform(current, xform);
        }
        else {
            this.logger.debug('Unknown ModifyWorldTransform mode; ignoring', { mode });
        }
    }
    /** Handles SelectObject for both created and stock objects. */
    selectObject(state, objects, index) {
        if ((index & STOCK_OBJECT_FLAG) !== 0) {
            this.selectStockObject(state, index & ~STOCK_OBJECT_FLAG);
            return;
        }
        const object = objects.get(index);
        if (!object) {
            this.logger.debug('SelectObject for unknown handle; ignoring', { index });
            return;
        }
        if (object.kind === 'pen') {
            state.pen = { ...object.pen };
        }
        else if (object.kind === 'brush') {
            state.brush = { ...object.brush };
        }
        else {
            state.font = { ...object.font };
        }
    }
    /** Applies a stock object selection to the DC. */
    selectStockObject(state, id) {
        switch (id) {
            case STOCK_OBJECT.WHITE_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00ffffff };
                break;
            case STOCK_OBJECT.LTGRAY_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00c0c0c0 };
                break;
            case STOCK_OBJECT.GRAY_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00808080 };
                break;
            case STOCK_OBJECT.DKGRAY_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00404040 };
                break;
            case STOCK_OBJECT.BLACK_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00000000 };
                break;
            case STOCK_OBJECT.NULL_BRUSH:
                state.brush = { style: BRUSH_STYLE.NULL, color: 0 };
                break;
            case STOCK_OBJECT.WHITE_PEN:
                state.pen = { style: PEN_STYLE.SOLID, width: 0, color: 0x00ffffff };
                break;
            case STOCK_OBJECT.BLACK_PEN:
            case STOCK_OBJECT.DC_PEN:
                state.pen = { style: PEN_STYLE.SOLID, width: 0, color: 0x00000000 };
                break;
            case STOCK_OBJECT.NULL_PEN:
                state.pen = { style: PEN_STYLE.NULL, width: 0, color: 0 };
                break;
            case STOCK_OBJECT.DC_BRUSH:
                state.brush = { style: BRUSH_STYLE.SOLID, color: 0x00ffffff };
                break;
            case STOCK_OBJECT.OEM_FIXED_FONT:
            case STOCK_OBJECT.ANSI_FIXED_FONT:
            case STOCK_OBJECT.ANSI_VAR_FONT:
            case STOCK_OBJECT.SYSTEM_FONT:
            case STOCK_OBJECT.DEVICE_DEFAULT_FONT:
            case STOCK_OBJECT.SYSTEM_FIXED_FONT:
            case STOCK_OBJECT.DEFAULT_GUI_FONT:
                state.font = { ...DEFAULT_FONT };
                break;
            default:
                this.logger.debug('Unknown stock object; ignoring', { id });
                break;
        }
    }
    /** Page-space scale factors from the current map mode. */
    pageScale(state, device) {
        switch (state.mapMode) {
            case MAP_MODE.LOMETRIC:
                return { sx: device.pxPerMmX * 0.1, sy: -device.pxPerMmY * 0.1 };
            case MAP_MODE.HIMETRIC:
                return { sx: device.pxPerMmX * 0.01, sy: -device.pxPerMmY * 0.01 };
            case MAP_MODE.LOENGLISH:
                return { sx: device.pxPerMmX * 0.254, sy: -device.pxPerMmY * 0.254 };
            case MAP_MODE.HIENGLISH:
                return { sx: device.pxPerMmX * 0.0254, sy: -device.pxPerMmY * 0.0254 };
            case MAP_MODE.TWIPS:
                return { sx: (device.pxPerMmX * 25.4) / 1440, sy: -(device.pxPerMmY * 25.4) / 1440 };
            case MAP_MODE.ISOTROPIC:
            case MAP_MODE.ANISOTROPIC: {
                const sx = state.windowExt.cx !== 0 ? state.viewportExt.cx / state.windowExt.cx : 1;
                const sy = state.windowExt.cy !== 0 ? state.viewportExt.cy / state.windowExt.cy : 1;
                if (state.mapMode === MAP_MODE.ISOTROPIC) {
                    const magnitude = Math.min(Math.abs(sx), Math.abs(sy));
                    return { sx: Math.sign(sx || 1) * magnitude, sy: Math.sign(sy || 1) * magnitude };
                }
                return { sx, sy };
            }
            case MAP_MODE.TEXT:
            default:
                return { sx: 1, sy: 1 };
        }
    }
    /** Transforms a logical point to device units (world -> page -> device). */
    toDevice(state, device, x, y) {
        let px = x;
        let py = y;
        if (state.world) {
            const w = state.world;
            const tx = w.m11 * px + w.m21 * py + w.dx;
            const ty = w.m12 * px + w.m22 * py + w.dy;
            px = tx;
            py = ty;
        }
        const { sx, sy } = this.pageScale(state, device);
        return {
            x: (px - state.windowOrg.x) * sx + state.viewportOrg.x,
            y: (py - state.windowOrg.y) * sy + state.viewportOrg.y,
        };
    }
    /** Magnitudes of the logical->device scale along each axis. */
    scaleMagnitudes(state, device) {
        const origin = this.toDevice(state, device, 0, 0);
        const unitX = this.toDevice(state, device, 1, 0);
        const unitY = this.toDevice(state, device, 0, 1);
        return {
            sx: Math.hypot(unitX.x - origin.x, unitX.y - origin.y) || 1,
            sy: Math.hypot(unitY.x - origin.x, unitY.y - origin.y) || 1,
        };
    }
    /** Runs a draw callback with the DC clip applied around it. */
    withClip(ctx, state, draw) {
        ctx.save();
        try {
            for (const step of state.clip) {
                if (step.length === 0) {
                    continue;
                }
                ctx.beginPath();
                for (const rect of step) {
                    ctx.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
                }
                ctx.clip();
            }
            draw();
        }
        finally {
            ctx.restore();
        }
    }
    /** Sets stroke style from the DC pen; returns false for PS_NULL. */
    applyPen(ctx, state, device) {
        const style = state.pen.style & PEN_STYLE.STYLE_MASK;
        if (style === PEN_STYLE.NULL) {
            return false;
        }
        const scales = this.scaleMagnitudes(state, device);
        const width = state.pen.width > 0 ? state.pen.width * ((scales.sx + scales.sy) / 2) : 1;
        ctx.strokeStyle = colorRefToCss(state.pen.color);
        ctx.lineWidth = Math.max(width, 0.5);
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        switch (style) {
            case PEN_STYLE.DASH:
                ctx.setLineDash([width * 4, width * 2]);
                break;
            case PEN_STYLE.DOT:
                ctx.setLineDash([width, width * 2]);
                break;
            case PEN_STYLE.DASHDOT:
                ctx.setLineDash([width * 4, width * 2, width, width * 2]);
                break;
            case PEN_STYLE.DASHDOTDOT:
                ctx.setLineDash([width * 4, width * 2, width, width * 2, width, width * 2]);
                break;
            default:
                ctx.setLineDash([]);
                break;
        }
        return true;
    }
    /** Sets fill style from the DC brush; returns false for BS_NULL. */
    applyBrush(ctx, state) {
        if (state.brush.style === BRUSH_STYLE.NULL) {
            return false;
        }
        // BS_HATCHED and BS_PATTERN are approximated as solid foreground fills
        ctx.fillStyle = colorRefToCss(state.brush.color);
        return true;
    }
    /** Builds a path via callback, then fills (brush) and strokes (pen). */
    fillStrokePath(ctx, state, device, buildPath) {
        this.withClip(ctx, state, () => {
            ctx.beginPath();
            buildPath();
            if (this.applyBrush(ctx, state)) {
                ctx.fill(state.polyFillMode === POLY_FILL_MODE.WINDING ? 'nonzero' : 'evenodd');
            }
            if (this.applyPen(ctx, state, device)) {
                ctx.stroke();
            }
        });
    }
    /** Builds a path via callback, then strokes it with the DC pen. */
    strokePath(ctx, state, device, buildPath) {
        this.withClip(ctx, state, () => {
            if (this.applyPen(ctx, state, device)) {
                ctx.beginPath();
                buildPath();
                ctx.stroke();
            }
        });
    }
    /** Appends a rectangle (4 transformed corners) to the current path. */
    rectPath(ctx, state, device, rect) {
        const corners = [
            this.toDevice(state, device, rect.left, rect.top),
            this.toDevice(state, device, rect.right, rect.top),
            this.toDevice(state, device, rect.right, rect.bottom),
            this.toDevice(state, device, rect.left, rect.bottom),
        ];
        const first = corners[0];
        if (!first) {
            return;
        }
        ctx.moveTo(first.x, first.y);
        for (let i = 1; i < corners.length; i++) {
            const corner = corners[i];
            if (corner) {
                ctx.lineTo(corner.x, corner.y);
            }
        }
        ctx.closePath();
    }
    /** Appends a poly(gon/line) to the current path. */
    polyPath(ctx, state, device, points, close) {
        if (points.length === 0) {
            return;
        }
        for (let i = 0; i < points.length; i++) {
            const point = points[i];
            if (!point) {
                continue;
            }
            const devicePoint = this.toDevice(state, device, point.x, point.y);
            if (i === 0) {
                ctx.moveTo(devicePoint.x, devicePoint.y);
            }
            else {
                ctx.lineTo(devicePoint.x, devicePoint.y);
            }
        }
        if (close) {
            ctx.closePath();
        }
    }
    /**
     * Appends cubic bezier segments to the current path.
     * @param fromCurrent When true (PolyBezierTo), the start point is the DC
     * current position and all points are control/end points.
     */
    bezierPath(ctx, state, device, points, fromCurrent) {
        let index = 0;
        let start;
        if (fromCurrent) {
            start = state.cur;
        }
        else {
            const first = points[0];
            if (!first) {
                return;
            }
            start = first;
            index = 1;
        }
        const startDevice = this.toDevice(state, device, start.x, start.y);
        ctx.moveTo(startDevice.x, startDevice.y);
        while (index + 3 <= points.length) {
            const c1 = points[index];
            const c2 = points[index + 1];
            const end = points[index + 2];
            if (!c1 || !c2 || !end) {
                break;
            }
            const d1 = this.toDevice(state, device, c1.x, c1.y);
            const d2 = this.toDevice(state, device, c2.x, c2.y);
            const de = this.toDevice(state, device, end.x, end.y);
            ctx.bezierCurveTo(d1.x, d1.y, d2.x, d2.y, de.x, de.y);
            index += 3;
        }
    }
    /** Renders an ExtTextOut record. */
    drawText(ctx, state, device, record) {
        if (record.text.length === 0) {
            return;
        }
        const scales = this.scaleMagnitudes(state, device);
        const font = state.font;
        // Negative lfHeight = em size; positive = cell height (em + leading)
        const emScale = font.height < 0 ? 1 : 0.85;
        const sizePx = Math.max(1, Math.abs(font.height) * scales.sy * emScale);
        const family = font.faceName.replace(/["\\]/g, '');
        const bold = font.weight >= 600 ? 'bold ' : '';
        const italic = font.italic ? 'italic ' : '';
        // Resolve through the shared font policy so registered fonts, user
        // fallback chains, and metric substitutes apply to metafile text too
        // (and a missing-font warning is emitted when the face is unavailable).
        const familyChain = this.fontResolver.getFontFamilyWithFallbacks(family);
        ctx.font = `${italic}${bold}${sizePx.toFixed(2)}px ${familyChain}`;
        ctx.textAlign = 'left';
        const useCurrentPos = (state.textAlign & TEXT_ALIGN.UPDATECP) !== 0;
        const refLogical = useCurrentPos ? state.cur : { x: record.x, y: record.y };
        const ref = this.toDevice(state, device, refLogical.x, refLogical.y);
        const advances = record.dx ? record.dx.map((d) => d * scales.sx) : undefined;
        const totalWidth = advances
            ? advances.reduce((sum, a) => sum + a, 0)
            : ctx.measureText(record.text).width;
        const horizontal = state.textAlign & TEXT_ALIGN.HORIZONTAL_MASK;
        let startX = ref.x;
        if (horizontal === TEXT_ALIGN.RIGHT) {
            startX -= totalWidth;
        }
        else if (horizontal === TEXT_ALIGN.CENTER) {
            startX -= totalWidth / 2;
        }
        const vertical = state.textAlign & TEXT_ALIGN.VERTICAL_MASK;
        let baseline = 'top';
        let baselineY = ref.y + sizePx * 0.8;
        if (vertical === TEXT_ALIGN.BASELINE) {
            baseline = 'alphabetic';
            baselineY = ref.y;
        }
        else if (vertical === TEXT_ALIGN.BOTTOM) {
            baseline = 'bottom';
            baselineY = ref.y - sizePx * 0.2;
        }
        ctx.textBaseline = baseline;
        this.withClip(ctx, state, () => {
            ctx.save();
            try {
                if ((record.options & ETO.CLIPPED) !== 0 && record.rect) {
                    const p1 = this.toDevice(state, device, record.rect.left, record.rect.top);
                    const p2 = this.toDevice(state, device, record.rect.right, record.rect.bottom);
                    const clipRect = normalizeRect(p1, p2);
                    ctx.beginPath();
                    ctx.rect(clipRect.left, clipRect.top, clipRect.right - clipRect.left, clipRect.bottom - clipRect.top);
                    ctx.clip();
                }
                // Opaque background: ETO_OPAQUE fills the record rect with bkColor;
                // OPAQUE bkMode fills the text cell
                if ((record.options & ETO.OPAQUE) !== 0 && record.rect) {
                    const p1 = this.toDevice(state, device, record.rect.left, record.rect.top);
                    const p2 = this.toDevice(state, device, record.rect.right, record.rect.bottom);
                    const bg = normalizeRect(p1, p2);
                    ctx.fillStyle = colorRefToCss(state.bkColor);
                    ctx.fillRect(bg.left, bg.top, bg.right - bg.left, bg.bottom - bg.top);
                }
                else if (state.bkMode === BK_MODE.OPAQUE) {
                    const cellTop = baselineY - sizePx * 0.8;
                    ctx.fillStyle = colorRefToCss(state.bkColor);
                    ctx.fillRect(startX, cellTop, totalWidth, sizePx * 1.05);
                }
                ctx.fillStyle = colorRefToCss(state.textColor);
                // GDI rasterizes metafile text with ClearType-style stem darkening,
                // noticeably heavier than skia's grayscale antialiasing: PowerPoint
                // exports show table digits so bold that neighbors touch. A hairline
                // outline stroke in the fill color (capped ~0.8px, ~em/40) restores
                // the reference weight without changing metrics.
                const darkening = Math.min(0.8, sizePx / 40);
                ctx.strokeStyle = ctx.fillStyle;
                ctx.lineWidth = darkening;
                ctx.setLineDash([]);
                const drawGlyph = (glyph, gx) => {
                    ctx.fillText(glyph, gx, ref.y);
                    if (darkening > 0.1) {
                        ctx.strokeText(glyph, gx, ref.y);
                    }
                };
                if (advances) {
                    let x = startX;
                    for (let i = 0; i < record.text.length; i++) {
                        // Keep surrogate pairs together: EMF stores UTF-16 code units
                        // with one dx entry each, so a pair consumes both advances.
                        const code = record.text.charCodeAt(i);
                        const isPair = code >= 0xd800 && code <= 0xdbff && i + 1 < record.text.length;
                        const glyph = isPair ? record.text.slice(i, i + 2) : record.text.charAt(i);
                        drawGlyph(glyph, x);
                        x += advances[i] ?? 0;
                        if (isPair) {
                            i++;
                            x += advances[i] ?? 0;
                        }
                    }
                }
                else {
                    drawGlyph(record.text, startX);
                }
                if (font.underline || font.strikeOut) {
                    ctx.strokeStyle = colorRefToCss(state.textColor);
                    ctx.lineWidth = Math.max(1, sizePx / 14);
                    ctx.setLineDash([]);
                    if (font.underline) {
                        ctx.beginPath();
                        ctx.moveTo(startX, baselineY + sizePx * 0.08);
                        ctx.lineTo(startX + totalWidth, baselineY + sizePx * 0.08);
                        ctx.stroke();
                    }
                    if (font.strikeOut) {
                        ctx.beginPath();
                        ctx.moveTo(startX, baselineY - sizePx * 0.3);
                        ctx.lineTo(startX + totalWidth, baselineY - sizePx * 0.3);
                        ctx.stroke();
                    }
                }
            }
            finally {
                ctx.restore();
            }
        });
        if (useCurrentPos && scales.sx > 0) {
            state.cur = { x: state.cur.x + totalWidth / scales.sx, y: state.cur.y };
        }
    }
    /** Renders a dibDraw record (BitBlt/StretchBlt/StretchDIBits). */
    drawDib(ctx, state, device, record, decoded) {
        const p1 = this.toDevice(state, device, record.destX, record.destY);
        const p2 = this.toDevice(state, device, record.destX + record.destWidth, record.destY + record.destHeight);
        const dest = normalizeRect(p1, p2);
        const destWidth = dest.right - dest.left;
        const destHeight = dest.bottom - dest.top;
        if (destWidth <= 0 || destHeight <= 0) {
            return;
        }
        this.withClip(ctx, state, () => {
            if (decoded) {
                const srcWidth = record.srcWidth > 0 ? record.srcWidth : decoded.width;
                const srcHeight = record.srcHeight > 0 ? record.srcHeight : decoded.height;
                // GDI measures ySrc from the bottom for bottom-up DIBs; the decoded
                // source is stored top-down, so remap the band.
                const srcYTopDown = decoded.bottomUp
                    ? Math.max(0, decoded.height - record.srcY - srcHeight)
                    : record.srcY;
                ctx.drawImage(decoded.source, record.srcX, srcYTopDown, srcWidth, srcHeight, dest.left, dest.top, destWidth, destHeight);
                return;
            }
            // No source bits: pattern/solid fills via raster op
            if (record.rop === ROP.PATCOPY) {
                if (this.applyBrush(ctx, state)) {
                    ctx.fillRect(dest.left, dest.top, destWidth, destHeight);
                }
            }
            else if (record.rop === ROP.BLACKNESS) {
                ctx.fillStyle = 'rgb(0,0,0)';
                ctx.fillRect(dest.left, dest.top, destWidth, destHeight);
            }
            else if (record.rop === ROP.WHITENESS) {
                ctx.fillStyle = 'rgb(255,255,255)';
                ctx.fillRect(dest.left, dest.top, destWidth, destHeight);
            }
            else {
                this.logger.debug('Unsupported raster operation without bits; skipping', {
                    rop: record.rop,
                });
            }
        });
    }
    /**
     * Decodes an embedded DIB. Uncompressed 24/32bpp DIBs are converted
     * directly; other formats are wrapped in a BMP file header and decoded
     * via skia-canvas.
     */
    async decodeDib(dib) {
        try {
            if (dib.bmi.length < 40) {
                return undefined;
            }
            const view = new DataView(dib.bmi.buffer, dib.bmi.byteOffset, dib.bmi.byteLength);
            const width = view.getInt32(4, true);
            const rawHeight = view.getInt32(8, true);
            const bitCount = view.getUint16(14, true);
            const compression = view.getUint32(16, true);
            const height = Math.abs(rawHeight);
            if (width <= 0 ||
                height === 0 ||
                width > 8192 ||
                height > 8192 ||
                width * height > 33_554_432 // 32 MP allocation guard
            ) {
                return undefined;
            }
            if ((bitCount === 24 || bitCount === 32) && compression === 0) {
                return this.decodeUncompressedDib(dib, width, rawHeight, bitCount);
            }
            // Fallback: synthesize a BMP file and let skia decode it
            const fileHeader = Buffer.alloc(14);
            fileHeader.write('BM', 0, 'ascii');
            fileHeader.writeUInt32LE(14 + dib.bmi.length + dib.bits.length, 2);
            fileHeader.writeUInt32LE(14 + dib.bmi.length, 10);
            const bmp = Buffer.concat([fileHeader, Buffer.from(dib.bmi), Buffer.from(dib.bits)]);
            const image = await loadImage(bmp);
            return { source: image, width: image.width, height: image.height, bottomUp: rawHeight > 0 };
        }
        catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            this.logger.warn('Failed to decode embedded DIB; skipping', { error: message });
            return undefined;
        }
    }
    /** Converts an uncompressed 24/32bpp DIB to a canvas via ImageData. */
    decodeUncompressedDib(dib, width, rawHeight, bitCount) {
        const height = Math.abs(rawHeight);
        const bytesPerPixel = bitCount / 8;
        const rowSize = (width * bytesPerPixel + 3) & ~3;
        if (rowSize * height > dib.bits.length) {
            return undefined;
        }
        const canvas = new Canvas(width, height);
        const ctx = canvas.getContext('2d');
        const imageData = ctx.createImageData(width, height);
        const data = imageData.data;
        const topDown = rawHeight < 0;
        for (let y = 0; y < height; y++) {
            const srcRow = topDown ? y : height - 1 - y;
            for (let x = 0; x < width; x++) {
                const src = srcRow * rowSize + x * bytesPerPixel;
                const dst = (y * width + x) * 4;
                data[dst] = dib.bits[src + 2] ?? 0;
                data[dst + 1] = dib.bits[src + 1] ?? 0;
                data[dst + 2] = dib.bits[src] ?? 0;
                // BI_RGB 32bpp reserves the 4th byte (usually 0); treat as opaque
                data[dst + 3] = 255;
            }
        }
        ctx.putImageData(imageData, 0, 0);
        return { source: canvas, width, height, bottomUp: rawHeight > 0 };
    }
}
/** Normalizes two corner points into a left/top/right/bottom rect. */
function normalizeRect(p1, p2) {
    return {
        left: Math.min(p1.x, p2.x),
        top: Math.min(p1.y, p2.y),
        right: Math.max(p1.x, p2.x),
        bottom: Math.max(p1.y, p2.y),
    };
}
/** Multiplies two XFORM matrices (a then b). */
function multiplyXform(a, b) {
    return {
        m11: a.m11 * b.m11 + a.m12 * b.m21,
        m12: a.m11 * b.m12 + a.m12 * b.m22,
        m21: a.m21 * b.m11 + a.m22 * b.m21,
        m22: a.m21 * b.m12 + a.m22 * b.m22,
        dx: a.dx * b.m11 + a.dy * b.m21 + b.dx,
        dy: a.dx * b.m12 + a.dy * b.m22 + b.dy,
    };
}
//# sourceMappingURL=EmfRenderer.js.map