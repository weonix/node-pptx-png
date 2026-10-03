/**
 * Text module for text layout and rendering.
 */
export { FontResolver, createFontResolver, setCustomFontFallbacks, getCustomFontFallbacks, type FontResolverConfig, type FontMetrics, type ResolvedFont, } from './FontResolver.js';
export { FontManager, createFontManager, decodeFntdata, deobfuscateOdttf, sniffFontFormat, getRegisteredFontFamilies, type FontManagerConfig, type FontLibraryLike, type FontFormat, type FntdataDecodeResult, type EmbeddedFontRegistrationResult, } from './FontManager.js';
export { BulletFormatter, createBulletFormatter, type BulletFormatterConfig, type BulletType, type AutoNumType, type BulletProps, type FormattedBullet, } from './BulletFormatter.js';
export { WordWrapper, createWordWrapper, type WordWrapperConfig, type WrapMode, type TextFragment, type WrappedLine, type WrapResult, } from './WordWrapper.js';
export { TextLayoutEngine, createTextLayoutEngine, type TextLayoutEngineConfig, type PositionedTextRun, type PositionedBullet, type LayoutLine, type TextLayout, } from './TextLayoutEngine.js';
export { AUTOFIT_CANDIDATES, AUTOFIT_FLOOR_INDEX, FIT_EPSILON_PX, solveAutofitScale, type AutofitCandidate, } from './AutofitSolver.js';
//# sourceMappingURL=index.d.ts.map