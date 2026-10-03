"use strict";
/**
 * EMF/WMF vector metafile rendering.
 *
 * Office paste content (Excel tables/charts, OLE previews) is stored as
 * EMF (or legacy WMF) images inside PPTX media. This module parses the
 * GDI record streams and replays them onto a skia-canvas Canvas so those
 * pictures render instead of vanishing.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.TEXT_ALIGN = exports.STOCK_OBJECT_FLAG = exports.STOCK_OBJECT = exports.ROP = exports.RGN_MODE = exports.POLY_FILL_MODE = exports.PEN_STYLE = exports.MWT = exports.MAP_MODE = exports.ETO = exports.EMR = exports.colorRefToCss = exports.BRUSH_STYLE = exports.BK_MODE = exports.EmfRenderer = exports.isWmf = exports.WmfParser = exports.isEmf = exports.EmfParser = void 0;
var EmfParser_js_1 = require("./EmfParser.js");
Object.defineProperty(exports, "EmfParser", { enumerable: true, get: function () { return EmfParser_js_1.EmfParser; } });
Object.defineProperty(exports, "isEmf", { enumerable: true, get: function () { return EmfParser_js_1.isEmf; } });
var WmfParser_js_1 = require("./WmfParser.js");
Object.defineProperty(exports, "WmfParser", { enumerable: true, get: function () { return WmfParser_js_1.WmfParser; } });
Object.defineProperty(exports, "isWmf", { enumerable: true, get: function () { return WmfParser_js_1.isWmf; } });
var EmfRenderer_js_1 = require("./EmfRenderer.js");
Object.defineProperty(exports, "EmfRenderer", { enumerable: true, get: function () { return EmfRenderer_js_1.EmfRenderer; } });
var types_js_1 = require("./types.js");
Object.defineProperty(exports, "BK_MODE", { enumerable: true, get: function () { return types_js_1.BK_MODE; } });
Object.defineProperty(exports, "BRUSH_STYLE", { enumerable: true, get: function () { return types_js_1.BRUSH_STYLE; } });
Object.defineProperty(exports, "colorRefToCss", { enumerable: true, get: function () { return types_js_1.colorRefToCss; } });
Object.defineProperty(exports, "EMR", { enumerable: true, get: function () { return types_js_1.EMR; } });
Object.defineProperty(exports, "ETO", { enumerable: true, get: function () { return types_js_1.ETO; } });
Object.defineProperty(exports, "MAP_MODE", { enumerable: true, get: function () { return types_js_1.MAP_MODE; } });
Object.defineProperty(exports, "MWT", { enumerable: true, get: function () { return types_js_1.MWT; } });
Object.defineProperty(exports, "PEN_STYLE", { enumerable: true, get: function () { return types_js_1.PEN_STYLE; } });
Object.defineProperty(exports, "POLY_FILL_MODE", { enumerable: true, get: function () { return types_js_1.POLY_FILL_MODE; } });
Object.defineProperty(exports, "RGN_MODE", { enumerable: true, get: function () { return types_js_1.RGN_MODE; } });
Object.defineProperty(exports, "ROP", { enumerable: true, get: function () { return types_js_1.ROP; } });
Object.defineProperty(exports, "STOCK_OBJECT", { enumerable: true, get: function () { return types_js_1.STOCK_OBJECT; } });
Object.defineProperty(exports, "STOCK_OBJECT_FLAG", { enumerable: true, get: function () { return types_js_1.STOCK_OBJECT_FLAG; } });
Object.defineProperty(exports, "TEXT_ALIGN", { enumerable: true, get: function () { return types_js_1.TEXT_ALIGN; } });
