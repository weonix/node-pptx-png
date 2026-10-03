/**
 * Visual Testing Infrastructure
 *
 * Provides tools for comparing rendered images against baseline images
 * and calculating fidelity scores to ensure rendering quality.
 */
// Visual Comparator
export { compareImages, VisualComparator, } from './VisualComparator.js';
// Fidelity Tester
export { runFidelityTest, formatReport, FidelityTester, } from './FidelityTester.js';
// Baseline Generator
export { generateBaselines, generateBaselinesViaPdf, findLibreOfficePath, isLibreOfficeAvailable, BaselineGenerator, } from './BaselineGenerator.js';
//# sourceMappingURL=index.js.map