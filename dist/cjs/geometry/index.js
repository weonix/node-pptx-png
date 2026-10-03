"use strict";
/**
 * Geometry module for path building and calculations.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.PRESET_SHAPE_DEFINITIONS = exports.parseCustomGeometryDefinition = exports.ooxmlArcToBeziers = exports.computeTextRect = exports.buildGeometryPaths = exports.resolveGuideValue = exports.radiansToOoxmlAngle = exports.ooxmlAngleToRadians = exports.evaluateGuideList = exports.evaluateFormula = exports.createGuideContext = exports.ANGLE_UNITS_PER_CIRCLE = exports.ANGLE_UNITS_PER_DEGREE = exports.presetGeometryCalculator = exports.PresetGeometryCalculator = exports.defaultTransformCalculator = exports.TransformCalculator = exports.pathToPath2D = exports.applyPathToContext = exports.calculatePathBounds = exports.PathBuilder = void 0;
var PathBuilder_js_1 = require("./PathBuilder.js");
Object.defineProperty(exports, "PathBuilder", { enumerable: true, get: function () { return PathBuilder_js_1.PathBuilder; } });
Object.defineProperty(exports, "calculatePathBounds", { enumerable: true, get: function () { return PathBuilder_js_1.calculatePathBounds; } });
Object.defineProperty(exports, "applyPathToContext", { enumerable: true, get: function () { return PathBuilder_js_1.applyPathToContext; } });
Object.defineProperty(exports, "pathToPath2D", { enumerable: true, get: function () { return PathBuilder_js_1.pathToPath2D; } });
var TransformCalculator_js_1 = require("./TransformCalculator.js");
Object.defineProperty(exports, "TransformCalculator", { enumerable: true, get: function () { return TransformCalculator_js_1.TransformCalculator; } });
Object.defineProperty(exports, "defaultTransformCalculator", { enumerable: true, get: function () { return TransformCalculator_js_1.defaultTransformCalculator; } });
var PresetGeometryCalculator_js_1 = require("./PresetGeometryCalculator.js");
Object.defineProperty(exports, "PresetGeometryCalculator", { enumerable: true, get: function () { return PresetGeometryCalculator_js_1.PresetGeometryCalculator; } });
Object.defineProperty(exports, "presetGeometryCalculator", { enumerable: true, get: function () { return PresetGeometryCalculator_js_1.presetGeometryCalculator; } });
var GuideEvaluator_js_1 = require("./GuideEvaluator.js");
Object.defineProperty(exports, "ANGLE_UNITS_PER_DEGREE", { enumerable: true, get: function () { return GuideEvaluator_js_1.ANGLE_UNITS_PER_DEGREE; } });
Object.defineProperty(exports, "ANGLE_UNITS_PER_CIRCLE", { enumerable: true, get: function () { return GuideEvaluator_js_1.ANGLE_UNITS_PER_CIRCLE; } });
Object.defineProperty(exports, "createGuideContext", { enumerable: true, get: function () { return GuideEvaluator_js_1.createGuideContext; } });
Object.defineProperty(exports, "evaluateFormula", { enumerable: true, get: function () { return GuideEvaluator_js_1.evaluateFormula; } });
Object.defineProperty(exports, "evaluateGuideList", { enumerable: true, get: function () { return GuideEvaluator_js_1.evaluateGuideList; } });
Object.defineProperty(exports, "ooxmlAngleToRadians", { enumerable: true, get: function () { return GuideEvaluator_js_1.ooxmlAngleToRadians; } });
Object.defineProperty(exports, "radiansToOoxmlAngle", { enumerable: true, get: function () { return GuideEvaluator_js_1.radiansToOoxmlAngle; } });
Object.defineProperty(exports, "resolveGuideValue", { enumerable: true, get: function () { return GuideEvaluator_js_1.resolveGuideValue; } });
var GeometryEngine_js_1 = require("./GeometryEngine.js");
Object.defineProperty(exports, "buildGeometryPaths", { enumerable: true, get: function () { return GeometryEngine_js_1.buildGeometryPaths; } });
Object.defineProperty(exports, "computeTextRect", { enumerable: true, get: function () { return GeometryEngine_js_1.computeTextRect; } });
Object.defineProperty(exports, "ooxmlArcToBeziers", { enumerable: true, get: function () { return GeometryEngine_js_1.ooxmlArcToBeziers; } });
var CustomGeometryParser_js_1 = require("./CustomGeometryParser.js");
Object.defineProperty(exports, "parseCustomGeometryDefinition", { enumerable: true, get: function () { return CustomGeometryParser_js_1.parseCustomGeometryDefinition; } });
var presetShapeDefinitions_js_1 = require("./presetShapeDefinitions.js");
Object.defineProperty(exports, "PRESET_SHAPE_DEFINITIONS", { enumerable: true, get: function () { return presetShapeDefinitions_js_1.PRESET_SHAPE_DEFINITIONS; } });
