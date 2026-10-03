/**
 * Geometry module for path building and calculations.
 */
export { PathBuilder, calculatePathBounds, applyPathToContext, pathToPath2D, } from './PathBuilder.js';
export { TransformCalculator, defaultTransformCalculator, type ParsedTransform, type PixelTransform, } from './TransformCalculator.js';
export { PresetGeometryCalculator, presetGeometryCalculator } from './PresetGeometryCalculator.js';
export { ANGLE_UNITS_PER_DEGREE, ANGLE_UNITS_PER_CIRCLE, createGuideContext, evaluateFormula, evaluateGuideList, ooxmlAngleToRadians, radiansToOoxmlAngle, resolveGuideValue, type GuideContext, } from './GuideEvaluator.js';
export { buildGeometryPaths, computeTextRect, ooxmlArcToBeziers, type GeomValue, type GeometryDefinition, type GeometryPathCommand, type GeometryPathDef, type GeometryPathFillMode, type GeometrySpace, } from './GeometryEngine.js';
export { parseCustomGeometryDefinition } from './CustomGeometryParser.js';
export { PRESET_SHAPE_DEFINITIONS } from './presetShapeDefinitions.js';
//# sourceMappingURL=index.d.ts.map