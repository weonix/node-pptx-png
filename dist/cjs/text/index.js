"use strict";
/**
 * Text module for text layout and rendering.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.solveAutofitScale = exports.FIT_EPSILON_PX = exports.AUTOFIT_FLOOR_INDEX = exports.AUTOFIT_CANDIDATES = exports.createTextLayoutEngine = exports.TextLayoutEngine = exports.createWordWrapper = exports.WordWrapper = exports.createBulletFormatter = exports.BulletFormatter = exports.getRegisteredFontFamilies = exports.sniffFontFormat = exports.deobfuscateOdttf = exports.decodeFntdata = exports.createFontManager = exports.FontManager = exports.getCustomFontFallbacks = exports.setCustomFontFallbacks = exports.createFontResolver = exports.FontResolver = void 0;
var FontResolver_js_1 = require("./FontResolver.js");
Object.defineProperty(exports, "FontResolver", { enumerable: true, get: function () { return FontResolver_js_1.FontResolver; } });
Object.defineProperty(exports, "createFontResolver", { enumerable: true, get: function () { return FontResolver_js_1.createFontResolver; } });
Object.defineProperty(exports, "setCustomFontFallbacks", { enumerable: true, get: function () { return FontResolver_js_1.setCustomFontFallbacks; } });
Object.defineProperty(exports, "getCustomFontFallbacks", { enumerable: true, get: function () { return FontResolver_js_1.getCustomFontFallbacks; } });
var FontManager_js_1 = require("./FontManager.js");
Object.defineProperty(exports, "FontManager", { enumerable: true, get: function () { return FontManager_js_1.FontManager; } });
Object.defineProperty(exports, "createFontManager", { enumerable: true, get: function () { return FontManager_js_1.createFontManager; } });
Object.defineProperty(exports, "decodeFntdata", { enumerable: true, get: function () { return FontManager_js_1.decodeFntdata; } });
Object.defineProperty(exports, "deobfuscateOdttf", { enumerable: true, get: function () { return FontManager_js_1.deobfuscateOdttf; } });
Object.defineProperty(exports, "sniffFontFormat", { enumerable: true, get: function () { return FontManager_js_1.sniffFontFormat; } });
Object.defineProperty(exports, "getRegisteredFontFamilies", { enumerable: true, get: function () { return FontManager_js_1.getRegisteredFontFamilies; } });
var BulletFormatter_js_1 = require("./BulletFormatter.js");
Object.defineProperty(exports, "BulletFormatter", { enumerable: true, get: function () { return BulletFormatter_js_1.BulletFormatter; } });
Object.defineProperty(exports, "createBulletFormatter", { enumerable: true, get: function () { return BulletFormatter_js_1.createBulletFormatter; } });
var WordWrapper_js_1 = require("./WordWrapper.js");
Object.defineProperty(exports, "WordWrapper", { enumerable: true, get: function () { return WordWrapper_js_1.WordWrapper; } });
Object.defineProperty(exports, "createWordWrapper", { enumerable: true, get: function () { return WordWrapper_js_1.createWordWrapper; } });
var TextLayoutEngine_js_1 = require("./TextLayoutEngine.js");
Object.defineProperty(exports, "TextLayoutEngine", { enumerable: true, get: function () { return TextLayoutEngine_js_1.TextLayoutEngine; } });
Object.defineProperty(exports, "createTextLayoutEngine", { enumerable: true, get: function () { return TextLayoutEngine_js_1.createTextLayoutEngine; } });
var AutofitSolver_js_1 = require("./AutofitSolver.js");
Object.defineProperty(exports, "AUTOFIT_CANDIDATES", { enumerable: true, get: function () { return AutofitSolver_js_1.AUTOFIT_CANDIDATES; } });
Object.defineProperty(exports, "AUTOFIT_FLOOR_INDEX", { enumerable: true, get: function () { return AutofitSolver_js_1.AUTOFIT_FLOOR_INDEX; } });
Object.defineProperty(exports, "FIT_EPSILON_PX", { enumerable: true, get: function () { return AutofitSolver_js_1.FIT_EPSILON_PX; } });
Object.defineProperty(exports, "solveAutofitScale", { enumerable: true, get: function () { return AutofitSolver_js_1.solveAutofitScale; } });
