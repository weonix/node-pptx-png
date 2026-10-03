"use strict";
/**
 * pptimg - PPTX to Image Converter
 *
 * High-fidelity PowerPoint presentation to image conversion for Node.js.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.EmfRenderer = exports.WmfParser = exports.EmfParser = exports.Logger = exports.createLogger = exports.TextLayoutEngine = exports.WordWrapper = exports.BulletFormatter = exports.FontResolver = exports.TextParser = exports.ShapeParser = exports.PresetGeometryCalculator = exports.TransformCalculator = exports.PathBuilder = exports.TextRenderer = exports.StrokeRenderer = exports.FillRenderer = exports.ShapeRenderer = exports.BackgroundRenderer = exports.SlideRenderer = exports.ColorResolver = exports.ThemeResolver = exports.fontSizeToPoints = exports.emuToPoints = exports.emuToPixels = exports.UnitConverter = exports.PptxParser = exports.Colors = exports.DEFAULT_FONT_SCHEME = exports.DEFAULT_OFFICE_COLORS = exports.DEFAULT_THEME = exports.DEFAULT_WARNING_LIMIT = exports.dedupeWarnings = exports.WarningCollector = exports.SIZE_PRESET_WIDTHS = exports.DEFAULT_RENDER_OPTIONS = exports.shardRoundRobin = exports.createRenderPool = exports.RenderPool = exports.openPresentation = exports.PptxDocument = exports.getSlideDimensions = exports.getSlideCount = exports.renderSlide = exports.renderPresentation = exports.createRenderer = exports.PptxImageRenderer = void 0;
// Main entry point
var PptxImageRenderer_js_1 = require("./core/PptxImageRenderer.js");
Object.defineProperty(exports, "PptxImageRenderer", { enumerable: true, get: function () { return PptxImageRenderer_js_1.PptxImageRenderer; } });
Object.defineProperty(exports, "createRenderer", { enumerable: true, get: function () { return PptxImageRenderer_js_1.createRenderer; } });
Object.defineProperty(exports, "renderPresentation", { enumerable: true, get: function () { return PptxImageRenderer_js_1.renderPresentation; } });
Object.defineProperty(exports, "renderSlide", { enumerable: true, get: function () { return PptxImageRenderer_js_1.renderSlide; } });
Object.defineProperty(exports, "getSlideCount", { enumerable: true, get: function () { return PptxImageRenderer_js_1.getSlideCount; } });
Object.defineProperty(exports, "getSlideDimensions", { enumerable: true, get: function () { return PptxImageRenderer_js_1.getSlideDimensions; } });
// Document handle API (parse once, render many)
var PptxDocument_js_1 = require("./core/PptxDocument.js");
Object.defineProperty(exports, "PptxDocument", { enumerable: true, get: function () { return PptxDocument_js_1.PptxDocument; } });
Object.defineProperty(exports, "openPresentation", { enumerable: true, get: function () { return PptxDocument_js_1.openPresentation; } });
// Worker-thread render pool (parallel rendering across CPU cores)
var RenderPool_js_1 = require("./pool/RenderPool.js");
Object.defineProperty(exports, "RenderPool", { enumerable: true, get: function () { return RenderPool_js_1.RenderPool; } });
Object.defineProperty(exports, "createRenderPool", { enumerable: true, get: function () { return RenderPool_js_1.createRenderPool; } });
Object.defineProperty(exports, "shardRoundRobin", { enumerable: true, get: function () { return RenderPool_js_1.shardRoundRobin; } });
var index_js_1 = require("./types/index.js");
Object.defineProperty(exports, "DEFAULT_RENDER_OPTIONS", { enumerable: true, get: function () { return index_js_1.DEFAULT_RENDER_OPTIONS; } });
Object.defineProperty(exports, "SIZE_PRESET_WIDTHS", { enumerable: true, get: function () { return index_js_1.SIZE_PRESET_WIDTHS; } });
// Structured warnings channel
var WarningCollector_js_1 = require("./utils/WarningCollector.js");
Object.defineProperty(exports, "WarningCollector", { enumerable: true, get: function () { return WarningCollector_js_1.WarningCollector; } });
Object.defineProperty(exports, "dedupeWarnings", { enumerable: true, get: function () { return WarningCollector_js_1.dedupeWarnings; } });
Object.defineProperty(exports, "DEFAULT_WARNING_LIMIT", { enumerable: true, get: function () { return WarningCollector_js_1.DEFAULT_WARNING_LIMIT; } });
var index_js_2 = require("./types/index.js");
Object.defineProperty(exports, "DEFAULT_THEME", { enumerable: true, get: function () { return index_js_2.DEFAULT_THEME; } });
Object.defineProperty(exports, "DEFAULT_OFFICE_COLORS", { enumerable: true, get: function () { return index_js_2.DEFAULT_OFFICE_COLORS; } });
Object.defineProperty(exports, "DEFAULT_FONT_SCHEME", { enumerable: true, get: function () { return index_js_2.DEFAULT_FONT_SCHEME; } });
var index_js_3 = require("./types/index.js");
Object.defineProperty(exports, "Colors", { enumerable: true, get: function () { return index_js_3.Colors; } });
// Core components (for advanced usage)
var PptxParser_js_1 = require("./core/PptxParser.js");
Object.defineProperty(exports, "PptxParser", { enumerable: true, get: function () { return PptxParser_js_1.PptxParser; } });
var UnitConverter_js_1 = require("./core/UnitConverter.js");
Object.defineProperty(exports, "UnitConverter", { enumerable: true, get: function () { return UnitConverter_js_1.UnitConverter; } });
Object.defineProperty(exports, "emuToPixels", { enumerable: true, get: function () { return UnitConverter_js_1.emuToPixels; } });
Object.defineProperty(exports, "emuToPoints", { enumerable: true, get: function () { return UnitConverter_js_1.emuToPoints; } });
Object.defineProperty(exports, "fontSizeToPoints", { enumerable: true, get: function () { return UnitConverter_js_1.fontSizeToPoints; } });
// Theme components (for advanced usage)
var ThemeResolver_js_1 = require("./theme/ThemeResolver.js");
Object.defineProperty(exports, "ThemeResolver", { enumerable: true, get: function () { return ThemeResolver_js_1.ThemeResolver; } });
var ColorResolver_js_1 = require("./theme/ColorResolver.js");
Object.defineProperty(exports, "ColorResolver", { enumerable: true, get: function () { return ColorResolver_js_1.ColorResolver; } });
// Rendering components (for advanced usage)
var SlideRenderer_js_1 = require("./rendering/SlideRenderer.js");
Object.defineProperty(exports, "SlideRenderer", { enumerable: true, get: function () { return SlideRenderer_js_1.SlideRenderer; } });
var BackgroundRenderer_js_1 = require("./rendering/BackgroundRenderer.js");
Object.defineProperty(exports, "BackgroundRenderer", { enumerable: true, get: function () { return BackgroundRenderer_js_1.BackgroundRenderer; } });
var index_js_4 = require("./rendering/index.js");
Object.defineProperty(exports, "ShapeRenderer", { enumerable: true, get: function () { return index_js_4.ShapeRenderer; } });
Object.defineProperty(exports, "FillRenderer", { enumerable: true, get: function () { return index_js_4.FillRenderer; } });
Object.defineProperty(exports, "StrokeRenderer", { enumerable: true, get: function () { return index_js_4.StrokeRenderer; } });
Object.defineProperty(exports, "TextRenderer", { enumerable: true, get: function () { return index_js_4.TextRenderer; } });
// Geometry components (for advanced usage)
var index_js_5 = require("./geometry/index.js");
Object.defineProperty(exports, "PathBuilder", { enumerable: true, get: function () { return index_js_5.PathBuilder; } });
Object.defineProperty(exports, "TransformCalculator", { enumerable: true, get: function () { return index_js_5.TransformCalculator; } });
Object.defineProperty(exports, "PresetGeometryCalculator", { enumerable: true, get: function () { return index_js_5.PresetGeometryCalculator; } });
// Parser components (for advanced usage)
var index_js_6 = require("./parsers/index.js");
Object.defineProperty(exports, "ShapeParser", { enumerable: true, get: function () { return index_js_6.ShapeParser; } });
Object.defineProperty(exports, "TextParser", { enumerable: true, get: function () { return index_js_6.TextParser; } });
// Text components (for advanced usage)
var index_js_7 = require("./text/index.js");
Object.defineProperty(exports, "FontResolver", { enumerable: true, get: function () { return index_js_7.FontResolver; } });
Object.defineProperty(exports, "BulletFormatter", { enumerable: true, get: function () { return index_js_7.BulletFormatter; } });
Object.defineProperty(exports, "WordWrapper", { enumerable: true, get: function () { return index_js_7.WordWrapper; } });
Object.defineProperty(exports, "TextLayoutEngine", { enumerable: true, get: function () { return index_js_7.TextLayoutEngine; } });
// Logger
var Logger_js_1 = require("./utils/Logger.js");
Object.defineProperty(exports, "createLogger", { enumerable: true, get: function () { return Logger_js_1.createLogger; } });
Object.defineProperty(exports, "Logger", { enumerable: true, get: function () { return Logger_js_1.Logger; } });
// Metafile (EMF/WMF) rendering
var index_js_8 = require("./emf/index.js");
Object.defineProperty(exports, "EmfParser", { enumerable: true, get: function () { return index_js_8.EmfParser; } });
Object.defineProperty(exports, "WmfParser", { enumerable: true, get: function () { return index_js_8.WmfParser; } });
Object.defineProperty(exports, "EmfRenderer", { enumerable: true, get: function () { return index_js_8.EmfRenderer; } });
