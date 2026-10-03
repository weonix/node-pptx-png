export { Logger, createLogger } from './Logger.js';
export type { ILogger, LogEntry } from './Logger.js';
export { ImageDecoder, createImageDecoder, type ImageFormat, type DecodedImage, type ImageDecoderConfig, } from './ImageDecoder.js';
export { PngOptimizer, createPngOptimizer, PNG_PRESETS } from './PngOptimizer.js';
export { WarningCollector, dedupeWarnings, DEFAULT_WARNING_LIMIT, type WarningCollectorOptions, } from './WarningCollector.js';
export type { PngOptimizationPreset, PngOptimizationOptions } from '../types/options.js';
