/**
 * Geometry module for path building and calculations.
 */
export { PathBuilder, calculatePathBounds, applyPathToContext, pathToPath2D, } from './PathBuilder.js';
export { TransformCalculator, defaultTransformCalculator, } from './TransformCalculator.js';
export { PresetGeometryCalculator, presetGeometryCalculator } from './PresetGeometryCalculator.js';
export { ANGLE_UNITS_PER_DEGREE, ANGLE_UNITS_PER_CIRCLE, createGuideContext, evaluateFormula, evaluateGuideList, ooxmlAngleToRadians, radiansToOoxmlAngle, resolveGuideValue, } from './GuideEvaluator.js';
export { buildGeometryPaths, computeTextRect, ooxmlArcToBeziers, } from './GeometryEngine.js';
export { parseCustomGeometryDefinition } from './CustomGeometryParser.js';
export { PRESET_SHAPE_DEFINITIONS } from './presetShapeDefinitions.js';
//# sourceMappingURL=index.js.map