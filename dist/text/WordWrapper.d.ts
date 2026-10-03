/**
 * Word wrapping algorithm for text layout.
 * Breaks text into lines that fit within a given width.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { FontResolver, ResolvedFont } from './FontResolver.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Text wrapping modes.
 */
export type WrapMode = 'none' | 'word' | 'char';
/**
 * A fragment of text with consistent styling.
 */
export interface TextFragment {
    /** Text content */
    text: string;
    /** Resolved font for this fragment */
    font: ResolvedFont;
    /** Width of fragment in pixels (calculated during wrapping) */
    width?: number;
    /** Additional run properties (color, underline, etc.) - preserved during wrapping */
    [key: string]: unknown;
}
/**
 * A wrapped line containing one or more text fragments.
 */
export interface WrappedLine {
    /** Fragments making up this line */
    fragments: TextFragment[];
    /** Total width of the line in pixels */
    width: number;
    /** Whether this line ends a paragraph */
    endsWithNewline: boolean;
}
/**
 * Result of word wrapping operation.
 */
export interface WrapResult {
    /** Wrapped lines */
    lines: WrappedLine[];
    /** Total height of all lines in pixels */
    totalHeight: number;
    /** Maximum line width in pixels */
    maxWidth: number;
}
/**
 * Configuration for WordWrapper.
 */
export interface WordWrapperConfig {
    /** Font resolver for text measurement */
    fontResolver: FontResolver;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Checks whether text from a source fragment can be merged into an existing
 * fragment. Fragments are only merged when BOTH the font string and all extra
 * run properties (color, underline, strikethrough, baseline, etc.) match —
 * merging on font alone would discard the second run's styling.
 */
export declare function canMergeFragments(existing: TextFragment | undefined, source: TextFragment): boolean;
/**
 * Checks whether a character is CJK for line-breaking purposes. CJK text is
 * breakable after every character rather than only at whitespace.
 *
 * @param char Single character
 * @returns True when the character is in a CJK range
 */
export declare function isCjkChar(char: string): boolean;
/**
 * Cheap scan reporting whether text contains any CJK characters. Used to
 * keep pure-Latin text on the unchanged whitespace-wrapping fast path.
 *
 * @param text Text to scan
 * @returns True when at least one CJK character is present
 */
export declare function containsCjk(text: string): boolean;
/**
 * Kinsoku: reports whether a wrap unit must not begin a line (closing
 * punctuation such as 、。，．）」』]!?：；).
 *
 * @param unit Wrap unit to test (only single characters can match)
 * @returns True when the unit is forbidden at line start
 */
export declare function isKinsokuNoStartChar(unit: string): boolean;
/**
 * Kinsoku: reports whether a character must not end a line (opening
 * brackets such as （「『[).
 *
 * @param char Character to test
 * @returns True when the character is forbidden at line end
 */
export declare function isKinsokuNoEndChar(char: string): boolean;
/**
 * Splits text into wrap units for mixed CJK/Latin wrapping: each CJK
 * character is its own unit (breakable after every character), while
 * non-CJK text is split on whitespace exactly like word wrapping
 * (whitespace runs are preserved as their own units).
 *
 * @param text Text to segment (a single line, without newlines)
 * @returns Ordered wrap units whose concatenation equals the input
 */
export declare function segmentWrapUnits(text: string): string[];
/**
 * Word wrapping algorithm for text layout.
 */
export declare class WordWrapper {
    private readonly logger;
    private readonly fontResolver;
    /** Cache for word widths to avoid O(n^2) re-measurement */
    private wordWidthCache;
    /** Cached space width for current font */
    private spaceWidthCache;
    constructor(config: WordWrapperConfig);
    /**
     * Gets cached word width or measures and caches it.
     */
    private getCachedWidth;
    /**
     * Gets cached space width for a font.
     */
    private getSpaceWidth;
    /**
     * Clears the word width cache. Call between different text bodies.
     */
    clearCache(): void;
    /**
     * Wraps text fragments to fit within the specified width.
     *
     * @param ctx Canvas 2D context for text measurement
     * @param fragments Text fragments to wrap
     * @param maxWidth Maximum width for each line in pixels
     * @param mode Wrapping mode ('none', 'word', 'char')
     * @param lineHeight Line height in pixels
     * @returns Wrapped lines and metrics
     */
    wrapText(ctx: CanvasRenderingContext2D, fragments: TextFragment[], maxWidth: number, mode: WrapMode | undefined, lineHeight: number): WrapResult;
    /**
     * No wrapping - puts all fragments on a single line.
     */
    private noWrap;
    /**
     * Wraps text at word boundaries.
     * Uses cached word widths and running totals to avoid O(n^2) re-measurement.
     */
    private wrapWords;
    /**
     * Wraps mixed CJK/Latin text. Each CJK character is an individually
     * breakable wrap unit while non-CJK text breaks at whitespace, with
     * kinsoku rules applied at every break: closing punctuation never begins
     * a line (a single closer hangs past maxWidth, burasage; longer closer
     * runs move to the next line with their preceding character, oidashi)
     * and opening brackets never end one (they are carried to the next line).
     * Both rules cross fragment boundaries within the working line, so
     * characters committed by earlier fragments of the paragraph (e.g. a
     * Latin '[' before an ea fragment) fold to the correct line.
     */
    private wrapMixed;
    /**
     * Wraps text at character boundaries (for CJK or when words don't fit).
     * Uses cached character widths and running totals to avoid O(n^2) re-measurement.
     */
    private wrapChars;
    /**
     * Measures text width using the font resolver.
     */
    private measureText;
    /**
     * Finds word boundaries in text.
     * Returns array of indices where words start.
     */
    findWordBoundaries(text: string): number[];
    /**
     * Checks if a character is a CJK (Chinese, Japanese, Korean) character.
     * Delegates to the module-level classifier shared with run splitting.
     */
    private isCjkChar;
    /**
     * Finds potential hyphenation points in a word.
     * Returns array of indices where hyphenation is allowed.
     * Note: This is a simple syllable-based approach; full hyphenation
     * would require a dictionary or hyphenation algorithm like Knuth-Liang.
     */
    findHyphenationPoints(word: string): number[];
}
/**
 * Creates a WordWrapper instance.
 */
export declare function createWordWrapper(fontResolver: FontResolver, logger?: ILogger): WordWrapper;
//# sourceMappingURL=WordWrapper.d.ts.map