/**
 * EMF/WMF vector metafile rendering.
 *
 * Office paste content (Excel tables/charts, OLE previews) is stored as
 * EMF (or legacy WMF) images inside PPTX media. This module parses the
 * GDI record streams and replays them onto a skia-canvas Canvas so those
 * pictures render instead of vanishing.
 */
export { EmfParser, isEmf } from './EmfParser.js';
export { WmfParser, isWmf } from './WmfParser.js';
export { EmfRenderer } from './EmfRenderer.js';
export { BK_MODE, BRUSH_STYLE, colorRefToCss, EMR, ETO, MAP_MODE, MWT, PEN_STYLE, POLY_FILL_MODE, RGN_MODE, ROP, STOCK_OBJECT, STOCK_OBJECT_FLAG, TEXT_ALIGN, } from './types.js';
//# sourceMappingURL=index.js.map