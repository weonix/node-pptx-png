/**
 * pptimg - PPTX to Image Converter
 *
 * High-fidelity PowerPoint presentation to image conversion for Node.js.
 */
export { PptxImageRenderer, createRenderer, renderPresentation, renderSlide, getSlideCount, getSlideDimensions, } from './core/PptxImageRenderer.js';
export type { IPptxImageRenderer } from './core/PptxImageRenderer.js';
export { PptxDocument, openPresentation } from './core/PptxDocument.js';
export { RenderPool, createRenderPool, shardRoundRobin } from './pool/RenderPool.js';
export type { RenderPoolOptions, PoolRenderOptions } from './pool/RenderPool.js';
export type { PptxRenderOptions, ImageFormat, LogLevel, SizePreset, OpenOptions, SlidesOptions, SlideSelection, SlideProgressEvent, ExportPdfOptions, SlideRenderResult, PresentationRenderResult, RenderError, RenderErrorLevel, RenderWarning, WarningCode, } from './types/index.js';
export { DEFAULT_RENDER_OPTIONS, SIZE_PRESET_WIDTHS } from './types/index.js';
export { WarningCollector, dedupeWarnings, DEFAULT_WARNING_LIMIT, } from './utils/WarningCollector.js';
export type { WarningCollectorOptions } from './utils/WarningCollector.js';
export type { ResolvedTheme, ResolvedColorScheme, ResolvedFontScheme, SchemeColorType, ColorTransform, } from './types/index.js';
export { DEFAULT_THEME, DEFAULT_OFFICE_COLORS, DEFAULT_FONT_SCHEME } from './types/index.js';
export type { Rgba, Point, Size, Rect, Transform2D, ShapeTransform, Path, PathSegment, } from './types/index.js';
export { Colors } from './types/index.js';
export { PptxParser } from './core/PptxParser.js';
export type { PresentationData, SlideData } from './core/PptxParser.js';
export { UnitConverter, emuToPixels, emuToPoints, fontSizeToPoints } from './core/UnitConverter.js';
export { ThemeResolver } from './theme/ThemeResolver.js';
export { ColorResolver } from './theme/ColorResolver.js';
export { SlideRenderer } from './rendering/SlideRenderer.js';
export type { SlideRenderContext } from './rendering/SlideRenderer.js';
export { BackgroundRenderer } from './rendering/BackgroundRenderer.js';
export { ShapeRenderer, FillRenderer, StrokeRenderer, TextRenderer } from './rendering/index.js';
export { PathBuilder, TransformCalculator, PresetGeometryCalculator } from './geometry/index.js';
export { ShapeParser, TextParser } from './parsers/index.js';
export { FontResolver, BulletFormatter, WordWrapper, TextLayoutEngine } from './text/index.js';
export type { FontMetrics, ResolvedFont, BulletProps, FormattedBullet, WrapMode, TextFragment, WrappedLine, PositionedTextRun, PositionedBullet, LayoutLine, TextLayout, } from './text/index.js';
export { createLogger, Logger } from './utils/Logger.js';
export type { ILogger } from './utils/Logger.js';
export { EmfParser, WmfParser, EmfRenderer } from './emf/index.js';
//# sourceMappingURL=index.d.ts.map