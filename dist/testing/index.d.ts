/**
 * Visual Testing Infrastructure
 *
 * Provides tools for comparing rendered images against baseline images
 * and calculating fidelity scores to ensure rendering quality.
 */
export { compareImages, VisualComparator, type ComparisonResult, type ComparisonOptions, } from './VisualComparator.js';
export { runFidelityTest, formatReport, FidelityTester, type FidelityReport, type SlideReport, type FidelityTestOptions, } from './FidelityTester.js';
export { generateBaselines, generateBaselinesViaPdf, findLibreOfficePath, isLibreOfficeAvailable, BaselineGenerator, type BaselineGenerationResult, type BaselineGeneratorOptions, } from './BaselineGenerator.js';
//# sourceMappingURL=index.d.ts.map