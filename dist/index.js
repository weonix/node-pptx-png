/**
 * pptimg - PPTX to Image Converter
 *
 * High-fidelity PowerPoint presentation to image conversion for Node.js.
 */
// Main entry point
export { PptxImageRenderer, createRenderer, renderPresentation, renderSlide, getSlideCount, getSlideDimensions, } from './core/PptxImageRenderer.js';
// Document handle API (parse once, render many)
export { PptxDocument, openPresentation } from './core/PptxDocument.js';
// Worker-thread render pool (parallel rendering across CPU cores)
export { RenderPool, createRenderPool, shardRoundRobin } from './pool/RenderPool.js';
export { DEFAULT_RENDER_OPTIONS, SIZE_PRESET_WIDTHS } from './types/index.js';
// Structured warnings channel
export { WarningCollector, dedupeWarnings, DEFAULT_WARNING_LIMIT, } from './utils/WarningCollector.js';
export { DEFAULT_THEME, DEFAULT_OFFICE_COLORS, DEFAULT_FONT_SCHEME } from './types/index.js';
export { Colors } from './types/index.js';
// Core components (for advanced usage)
export { PptxParser } from './core/PptxParser.js';
export { UnitConverter, emuToPixels, emuToPoints, fontSizeToPoints } from './core/UnitConverter.js';
// Theme components (for advanced usage)
export { ThemeResolver } from './theme/ThemeResolver.js';
export { ColorResolver } from './theme/ColorResolver.js';
// Rendering components (for advanced usage)
export { SlideRenderer } from './rendering/SlideRenderer.js';
export { BackgroundRenderer } from './rendering/BackgroundRenderer.js';
export { ShapeRenderer, FillRenderer, StrokeRenderer, TextRenderer } from './rendering/index.js';
// Geometry components (for advanced usage)
export { PathBuilder, TransformCalculator, PresetGeometryCalculator } from './geometry/index.js';
// Parser components (for advanced usage)
export { ShapeParser, TextParser } from './parsers/index.js';
// Text components (for advanced usage)
export { FontResolver, BulletFormatter, WordWrapper, TextLayoutEngine } from './text/index.js';
// Logger
export { createLogger, Logger } from './utils/Logger.js';
// Metafile (EMF/WMF) rendering
export { EmfParser, WmfParser, EmfRenderer } from './emf/index.js';
//# sourceMappingURL=index.js.map