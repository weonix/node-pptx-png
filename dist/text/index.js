/**
 * Text module for text layout and rendering.
 */
export { FontResolver, createFontResolver, setCustomFontFallbacks, getCustomFontFallbacks, } from './FontResolver.js';
export { FontManager, createFontManager, decodeFntdata, deobfuscateOdttf, sniffFontFormat, getRegisteredFontFamilies, } from './FontManager.js';
export { BulletFormatter, createBulletFormatter, } from './BulletFormatter.js';
export { WordWrapper, createWordWrapper, } from './WordWrapper.js';
export { TextLayoutEngine, createTextLayoutEngine, } from './TextLayoutEngine.js';
export { AUTOFIT_CANDIDATES, AUTOFIT_FLOOR_INDEX, FIT_EPSILON_PX, solveAutofitScale, } from './AutofitSolver.js';
//# sourceMappingURL=index.js.map