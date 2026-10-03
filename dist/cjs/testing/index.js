"use strict";
/**
 * Visual Testing Infrastructure
 *
 * Provides tools for comparing rendered images against baseline images
 * and calculating fidelity scores to ensure rendering quality.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.BaselineGenerator = exports.isLibreOfficeAvailable = exports.findLibreOfficePath = exports.generateBaselinesViaPdf = exports.generateBaselines = exports.FidelityTester = exports.formatReport = exports.runFidelityTest = exports.VisualComparator = exports.compareImages = void 0;
// Visual Comparator
var VisualComparator_js_1 = require("./VisualComparator.js");
Object.defineProperty(exports, "compareImages", { enumerable: true, get: function () { return VisualComparator_js_1.compareImages; } });
Object.defineProperty(exports, "VisualComparator", { enumerable: true, get: function () { return VisualComparator_js_1.VisualComparator; } });
// Fidelity Tester
var FidelityTester_js_1 = require("./FidelityTester.js");
Object.defineProperty(exports, "runFidelityTest", { enumerable: true, get: function () { return FidelityTester_js_1.runFidelityTest; } });
Object.defineProperty(exports, "formatReport", { enumerable: true, get: function () { return FidelityTester_js_1.formatReport; } });
Object.defineProperty(exports, "FidelityTester", { enumerable: true, get: function () { return FidelityTester_js_1.FidelityTester; } });
// Baseline Generator
var BaselineGenerator_js_1 = require("./BaselineGenerator.js");
Object.defineProperty(exports, "generateBaselines", { enumerable: true, get: function () { return BaselineGenerator_js_1.generateBaselines; } });
Object.defineProperty(exports, "generateBaselinesViaPdf", { enumerable: true, get: function () { return BaselineGenerator_js_1.generateBaselinesViaPdf; } });
Object.defineProperty(exports, "findLibreOfficePath", { enumerable: true, get: function () { return BaselineGenerator_js_1.findLibreOfficePath; } });
Object.defineProperty(exports, "isLibreOfficeAvailable", { enumerable: true, get: function () { return BaselineGenerator_js_1.isLibreOfficeAvailable; } });
Object.defineProperty(exports, "BaselineGenerator", { enumerable: true, get: function () { return BaselineGenerator_js_1.BaselineGenerator; } });
