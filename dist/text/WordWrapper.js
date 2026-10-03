/**
 * Word wrapping algorithm for text layout.
 * Breaks text into lines that fit within a given width.
 */
import { createLogger } from '../utils/Logger.js';
/**
 * Deep value equality for plain parsed data (primitives, arrays, plain objects).
 * Used to compare fragment run properties when deciding whether to merge.
 */
function deepEqual(a, b) {
    if (a === b)
        return true;
    if (typeof a !== typeof b)
        return false;
    if (a === null || b === null)
        return false;
    if (Array.isArray(a) || Array.isArray(b)) {
        if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length)
            return false;
        return a.every((item, i) => deepEqual(item, b[i]));
    }
    if (typeof a === 'object' && typeof b === 'object') {
        const aObj = a;
        const bObj = b;
        const keys = new Set([...Object.keys(aObj), ...Object.keys(bObj)]);
        for (const key of keys) {
            if (!deepEqual(aObj[key], bObj[key]))
                return false;
        }
        return true;
    }
    return false;
}
/**
 * Checks whether text from a source fragment can be merged into an existing
 * fragment. Fragments are only merged when BOTH the font string and all extra
 * run properties (color, underline, strikethrough, baseline, etc.) match —
 * merging on font alone would discard the second run's styling.
 */
export function canMergeFragments(existing, source) {
    if (!existing)
        return false;
    if (existing.font.fontString !== source.font.fontString)
        return false;
    if ((existing.font.letterSpacingPx ?? 0) !== (source.font.letterSpacingPx ?? 0))
        return false;
    // Compare all extra properties (e.g. runProps) by value
    const keys = new Set([...Object.keys(existing), ...Object.keys(source)]);
    keys.delete('text');
    keys.delete('font');
    keys.delete('width');
    for (const key of keys) {
        if (!deepEqual(existing[key], source[key]))
            return false;
    }
    return true;
}
/**
 * Matches any character treated as CJK for line-breaking and font selection:
 * Hangul Jamo (U+1100-11FF), CJK punctuation through CJK Compatibility
 * incl. Hiragana/Katakana, compat jamo, and unit signs (U+3000-33FF), CJK
 * Extension A (U+3400-4DBF), CJK Unified Ideographs (U+4E00-9FFF), Hangul
 * syllables (U+AC00-D7AF), CJK Compatibility Ideographs (U+F900-FAFF),
 * fullwidth/halfwidth forms (U+FF00-FFEF), and astral CJK Extensions B+
 * (U+20000-3FFFF).
 */
const CJK_CHAR_REGEX = /[\u1100-\u11FF\u3000-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF\u{20000}-\u{3FFFF}]/u;
/**
 * Kinsoku shori: closing punctuation that must not begin a line. Includes
 * the fullwidth analogues of the ASCII terminators.
 */
const KINSOKU_NO_LINE_START = new Set([
    ...'、。，．）」』]!?：；！？｝〕】〉》］・ー)},.',
    ...'ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ',
]);
/**
 * Kinsoku shori: opening brackets that must not end a line.
 */
const KINSOKU_NO_LINE_END = new Set([...'（「『[｛〔【〈《［(']);
/**
 * Checks whether a character is CJK for line-breaking purposes. CJK text is
 * breakable after every character rather than only at whitespace.
 *
 * @param char Single character
 * @returns True when the character is in a CJK range
 */
export function isCjkChar(char) {
    return CJK_CHAR_REGEX.test(char);
}
/**
 * Cheap scan reporting whether text contains any CJK characters. Used to
 * keep pure-Latin text on the unchanged whitespace-wrapping fast path.
 *
 * @param text Text to scan
 * @returns True when at least one CJK character is present
 */
export function containsCjk(text) {
    return CJK_CHAR_REGEX.test(text);
}
/**
 * Kinsoku: reports whether a wrap unit must not begin a line (closing
 * punctuation such as 、。，．）」』]!?：；).
 *
 * @param unit Wrap unit to test (only single characters can match)
 * @returns True when the unit is forbidden at line start
 */
export function isKinsokuNoStartChar(unit) {
    return KINSOKU_NO_LINE_START.has(unit);
}
/**
 * Kinsoku: reports whether a character must not end a line (opening
 * brackets such as （「『[).
 *
 * @param char Character to test
 * @returns True when the character is forbidden at line end
 */
export function isKinsokuNoEndChar(char) {
    return KINSOKU_NO_LINE_END.has(char);
}
/**
 * Reports whether text begins with a kinsoku no-start (closing punctuation)
 * character. Used to engage CJK-aware wrapping for the fragment directly
 * after a CJK fragment so the closer cannot strand at a line start.
 */
function startsWithKinsokuNoStart(text) {
    const codePoint = text.codePointAt(0);
    return codePoint !== undefined && KINSOKU_NO_LINE_START.has(String.fromCodePoint(codePoint));
}
/**
 * Splits text into wrap units for mixed CJK/Latin wrapping: each CJK
 * character is its own unit (breakable after every character), while
 * non-CJK text is split on whitespace exactly like word wrapping
 * (whitespace runs are preserved as their own units).
 *
 * @param text Text to segment (a single line, without newlines)
 * @returns Ordered wrap units whose concatenation equals the input
 */
export function segmentWrapUnits(text) {
    const units = [];
    let nonCjk = '';
    const flushNonCjk = () => {
        if (!nonCjk)
            return;
        for (const token of nonCjk.split(/(\s+)/)) {
            if (token !== '')
                units.push(token);
        }
        nonCjk = '';
    };
    for (const char of text) {
        if (isCjkChar(char)) {
            flushNonCjk();
            units.push(char);
        }
        else {
            nonCjk += char;
        }
    }
    flushNonCjk();
    return units;
}
/**
 * Word wrapping algorithm for text layout.
 */
export class WordWrapper {
    logger;
    fontResolver;
    /** Cache for word widths to avoid O(n^2) re-measurement */
    wordWidthCache = new Map();
    /** Cached space width for current font */
    spaceWidthCache = new Map();
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'WordWrapper');
        this.fontResolver = config.fontResolver;
    }
    /**
     * Gets cached word width or measures and caches it.
     */
    getCachedWidth(ctx, text, font) {
        // Letter tracking affects measured widths, so it is part of the key
        const cacheKey = `${font.fontString}:${font.letterSpacingPx ?? 0}:${text}`;
        const cached = this.wordWidthCache.get(cacheKey);
        if (cached !== undefined) {
            return cached;
        }
        const width = this.fontResolver.measureText(ctx, text, font);
        this.wordWidthCache.set(cacheKey, width);
        return width;
    }
    /**
     * Gets cached space width for a font.
     */
    getSpaceWidth(ctx, font) {
        const cacheKey = `${font.fontString}:${font.letterSpacingPx ?? 0}`;
        const cached = this.spaceWidthCache.get(cacheKey);
        if (cached !== undefined) {
            return cached;
        }
        const width = this.fontResolver.measureText(ctx, ' ', font);
        this.spaceWidthCache.set(cacheKey, width);
        return width;
    }
    /**
     * Clears the word width cache. Call between different text bodies.
     */
    clearCache() {
        this.wordWidthCache.clear();
        this.spaceWidthCache.clear();
    }
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
    wrapText(ctx, fragments, maxWidth, mode = 'word', lineHeight) {
        if (mode === 'none') {
            return this.noWrap(ctx, fragments, lineHeight);
        }
        const lines = [];
        let currentLine = [];
        let currentLineWidth = 0;
        // Whether the previously wrapped part (same paragraph, no intervening
        // newline) contained CJK text. Lets kinsoku act across the fragment
        // boundaries created by script splitting: a fragment beginning with
        // closing punctuation directly after an ea fragment (e.g. ']' after
        // '世界') must go through the kinsoku-aware wrapper.
        let prevPartCjk = false;
        for (const fragment of fragments) {
            // Handle newlines in fragment text
            const parts = fragment.text.split(/(\r?\n)/);
            for (let i = 0; i < parts.length; i++) {
                const part = parts[i];
                if (part === undefined)
                    continue;
                // Check for newline
                if (part === '\n' || part === '\r\n') {
                    // End current line
                    lines.push({
                        fragments: currentLine.length > 0 ? currentLine : [{ text: '', font: fragment.font, width: 0 }],
                        width: currentLineWidth,
                        endsWithNewline: true,
                    });
                    currentLine = [];
                    currentLineWidth = 0;
                    prevPartCjk = false;
                    continue;
                }
                if (part === '')
                    continue;
                // CJK-aware wrapping is engaged for parts containing CJK text and
                // for the part directly after a CJK part when it begins with a
                // no-start closer (cross-fragment kinsoku); other pure-Latin
                // content stays on the unchanged whitespace-wrapping fast path.
                const partCjk = containsCjk(part);
                const useCjkWrap = mode === 'word' && (partCjk || (prevPartCjk && startsWithKinsokuNoStart(part)));
                prevPartCjk = partCjk;
                // Wrap this part and get updated state
                let wrapState;
                if (mode === 'word') {
                    wrapState = useCjkWrap
                        ? this.wrapMixed(ctx, part, fragment, maxWidth, currentLine, currentLineWidth, lines)
                        : this.wrapWords(ctx, part, fragment, maxWidth, currentLine, currentLineWidth, lines);
                }
                else {
                    wrapState = this.wrapChars(ctx, part, fragment, maxWidth, currentLine, currentLineWidth, lines);
                }
                // Update current line state from the returned WrapState
                if (lines.length > 0) {
                    const lastLine = lines[lines.length - 1];
                    if (lastLine && !lastLine.endsWithNewline) {
                        // The last line is the current working line - use state from wrapState
                        currentLine = wrapState.workingLine;
                        currentLineWidth = wrapState.lineWidth;
                        lines.pop();
                    }
                    else {
                        currentLine = [];
                        currentLineWidth = 0;
                    }
                }
            }
        }
        // Add remaining text as final line
        if (currentLine.length > 0) {
            lines.push({
                fragments: currentLine,
                width: currentLineWidth,
                endsWithNewline: false,
            });
        }
        // Calculate total metrics
        const totalHeight = lines.length * lineHeight;
        const maxLineWidth = Math.max(0, ...lines.map((l) => l.width));
        this.logger.debug('Wrapped text', {
            fragmentCount: fragments.length,
            lineCount: lines.length,
            maxWidth: maxLineWidth,
        });
        return {
            lines,
            totalHeight,
            maxWidth: maxLineWidth,
        };
    }
    /**
     * No wrapping - puts all fragments on a single line.
     */
    noWrap(ctx, fragments, lineHeight) {
        const measuredFragments = [];
        let totalWidth = 0;
        for (const fragment of fragments) {
            const text = fragment.text.replace(/\r?\n/g, ' ');
            const width = this.measureText(ctx, text, fragment.font);
            measuredFragments.push({ ...fragment, text, width });
            totalWidth += width;
        }
        return {
            lines: [
                {
                    fragments: measuredFragments,
                    width: totalWidth,
                    endsWithNewline: false,
                },
            ],
            totalHeight: lineHeight,
            maxWidth: totalWidth,
        };
    }
    /**
     * Wraps text at word boundaries.
     * Uses cached word widths and running totals to avoid O(n^2) re-measurement.
     */
    wrapWords(ctx, text, sourceFragment, maxWidth, currentLine, currentLineWidth, lines) {
        const font = sourceFragment.font;
        // Split text into words, preserving spaces
        const words = text.split(/(\s+)/);
        let lineWidth = currentLineWidth;
        let lineText = '';
        let lineTextWidth = 0;
        let workingLine = [...currentLine]; // Clone to avoid mutating input
        // Helper to create fragment preserving extra properties from source
        const createFragment = (fragText, fragWidth) => {
            const frag = { text: fragText, font, width: fragWidth };
            // Copy extra properties from source fragment (like runProps)
            for (const key of Object.keys(sourceFragment)) {
                if (key !== 'text' && key !== 'font' && key !== 'width') {
                    frag[key] = sourceFragment[key];
                }
            }
            return frag;
        };
        // A fragment marked noBreakBefore starts mid-word (e.g. a small-caps
        // continuation segment): its first word must stay glued to the line
        // rather than becoming a break opportunity.
        let glueNextWord = sourceFragment['noBreakBefore'] === true && currentLineWidth > 0;
        for (const word of words) {
            if (word === '')
                continue;
            // Use cached width measurement
            const wordWidth = this.getCachedWidth(ctx, word, font);
            const glued = glueNextWord;
            glueNextWord = false;
            // Whitespace never wraps: the space at a soft break stays at the end
            // of the current line, hanging past the wrap edge (PowerPoint
            // behavior). Lines therefore never start with a wrap-point space.
            const isWhitespace = /^\s+$/.test(word);
            // Check if word fits on current line
            if (glued || isWhitespace || lineWidth + wordWidth <= maxWidth || lineWidth === 0) {
                // Word fits, add to current line - track width incrementally
                lineText += word;
                lineTextWidth += wordWidth;
                lineWidth += wordWidth;
            }
            else {
                // Word doesn't fit, start new line
                if (lineText) {
                    // Add current line text as fragment with pre-computed width.
                    // Only merge into the previous fragment when its full run
                    // properties match, not just the font string.
                    const lastFrag = workingLine[workingLine.length - 1];
                    if (canMergeFragments(lastFrag, sourceFragment) && lastFrag) {
                        lastFrag.text += lineText;
                        // Add to existing width instead of re-measuring
                        lastFrag.width = (lastFrag.width ?? 0) + lineTextWidth;
                    }
                    else {
                        workingLine.push(createFragment(lineText, lineTextWidth));
                    }
                }
                lines.push({
                    fragments: workingLine,
                    width: lineWidth,
                    endsWithNewline: false,
                });
                // Start new line with current word - create new array instead of splice
                workingLine = [];
                lineText = word;
                lineTextWidth = wordWidth;
                lineWidth = wordWidth;
            }
        }
        // Add remaining text to current line. Only merge into the previous
        // fragment when its full run properties match, not just the font string.
        if (lineText) {
            const lastFrag = workingLine[workingLine.length - 1];
            if (canMergeFragments(lastFrag, sourceFragment) && lastFrag) {
                lastFrag.text += lineText;
                // Add to existing width instead of re-measuring
                lastFrag.width = (lastFrag.width ?? 0) + lineTextWidth;
            }
            else {
                workingLine.push(createFragment(lineText, lineTextWidth));
            }
        }
        // Push current line as the "working" line
        lines.push({
            fragments: [...workingLine],
            width: lineWidth,
            endsWithNewline: false,
        });
        // Return the updated state
        return { workingLine, lineWidth };
    }
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
    wrapMixed(ctx, text, sourceFragment, maxWidth, currentLine, currentLineWidth, lines) {
        const font = sourceFragment.font;
        const units = segmentWrapUnits(text);
        let lineWidth = currentLineWidth;
        let lineText = '';
        let lineTextWidth = 0;
        let workingLine = [...currentLine]; // Clone to avoid mutating input
        // Helper to create fragment preserving extra properties from source
        const createFragment = (fragText, fragWidth) => {
            const frag = { text: fragText, font, width: fragWidth };
            // Copy extra properties from source fragment (like runProps)
            for (const key of Object.keys(sourceFragment)) {
                if (key !== 'text' && key !== 'font' && key !== 'width') {
                    frag[key] = sourceFragment[key];
                }
            }
            return frag;
        };
        // Appends already-measured text to the accumulated line state
        const appendText = (unitText, unitWidth) => {
            lineText += unitText;
            lineTextWidth += unitWidth;
            lineWidth += unitWidth;
        };
        // Commits the accumulated text to the working line, pushes it as a
        // completed (non-newline) line, and resets the accumulation state.
        // Merging into the previous fragment follows the same full-run-property
        // rule as wrapWords.
        const flushLine = () => {
            if (lineText) {
                const lastFrag = workingLine[workingLine.length - 1];
                if (canMergeFragments(lastFrag, sourceFragment) && lastFrag) {
                    lastFrag.text += lineText;
                    lastFrag.width = (lastFrag.width ?? 0) + lineTextWidth;
                }
                else {
                    workingLine.push(createFragment(lineText, lineTextWidth));
                }
            }
            lines.push({ fragments: workingLine, width: lineWidth, endsWithNewline: false });
            workingLine = [];
            lineText = '';
            lineTextWidth = 0;
            lineWidth = 0;
        };
        // Counts the characters currently on the line (committed fragments plus
        // the in-progress accumulation); used to keep kinsoku carries from
        // emptying the line entirely.
        const lineCharCount = () => workingLine.reduce((count, frag) => count + [...frag.text].length, 0) + [...lineText].length;
        // Reads the line's last character without removing it.
        const peekLastChar = () => {
            if (lineText)
                return [...lineText].pop();
            for (let f = workingLine.length - 1; f >= 0; f--) {
                const fragText = workingLine[f]?.text ?? '';
                if (fragText)
                    return [...fragText].pop();
            }
            return undefined;
        };
        // Removes the line's last character and returns it with the fragment
        // whose styling it carries. Crosses fragment boundaries so kinsoku can
        // act on characters committed by earlier fragments of the paragraph
        // (e.g. a Latin '[' fragment directly before this ea fragment).
        const pullLastChar = () => {
            if (lineText) {
                const chars = [...lineText];
                const char = chars.pop();
                if (char === undefined)
                    return undefined;
                const width = this.getCachedWidth(ctx, char, font);
                lineText = chars.join('');
                lineTextWidth -= width;
                lineWidth -= width;
                return { text: char, width, source: sourceFragment };
            }
            while (workingLine.length > 0) {
                const frag = workingLine[workingLine.length - 1];
                if (!frag)
                    return undefined;
                const chars = [...frag.text];
                const char = chars.pop();
                if (char === undefined) {
                    workingLine.pop();
                    continue;
                }
                const width = this.getCachedWidth(ctx, char, frag.font);
                frag.text = chars.join('');
                frag.width = Math.max(0, (frag.width ?? 0) - width);
                lineWidth -= width;
                if (frag.text === '')
                    workingLine.pop();
                return { text: char, width, source: frag };
            }
            return undefined;
        };
        // Re-emits carried characters (in reading order) at the start of the
        // fresh working line, preserving each piece's source styling.
        const seedCarried = (pieces) => {
            for (const piece of pieces) {
                const last = workingLine[workingLine.length - 1];
                if (last && canMergeFragments(last, piece.source)) {
                    last.text += piece.text;
                    last.width = (last.width ?? 0) + piece.width;
                }
                else {
                    workingLine.push({ ...piece.source, text: piece.text, width: piece.width });
                }
                lineWidth += piece.width;
            }
        };
        // Kinsoku: opening brackets must not end a line. Pulls trailing openers
        // (across fragment boundaries) so they move to the next line together
        // with the unit that triggered the break. An opener that is the line's
        // only character stays: never carry the entire line.
        const pullTrailingOpeners = () => {
            const pulled = [];
            while (lineCharCount() > 1) {
                const last = peekLastChar();
                if (last === undefined || !isKinsokuNoEndChar(last))
                    break;
                const piece = pullLastChar();
                if (!piece)
                    break;
                pulled.unshift(piece);
            }
            return pulled;
        };
        for (let i = 0; i < units.length; i++) {
            const unit = units[i];
            if (unit === undefined)
                continue;
            const unitWidth = this.getCachedWidth(ctx, unit, font);
            // Whitespace units hang at the end of the current line rather than
            // starting the next one (PowerPoint behavior; mirrors wrapWords) —
            // lines therefore never begin with a wrap-point space.
            const isWhitespaceUnit = /^\s+$/.test(unit);
            // Unit fits (a leading unit always fits, matching wrapWords)
            if (isWhitespaceUnit || lineWidth + unitWidth <= maxWidth || lineWidth === 0) {
                appendText(unit, unitWidth);
                continue;
            }
            // A multi-character unit may begin with closing punctuation (e.g.
            // ']test' produced by script splitting); split the leading closers
            // into their own units so the no-start rule below can act on them.
            if (!isKinsokuNoStartChar(unit)) {
                const chars = [...unit];
                const first = chars[0];
                if (chars.length > 1 && first !== undefined && isKinsokuNoStartChar(first)) {
                    let split = 1;
                    while (split < chars.length) {
                        const char = chars[split];
                        if (char === undefined || !isKinsokuNoStartChar(char))
                            break;
                        split++;
                    }
                    const rest = chars.slice(split).join('');
                    units.splice(i, 1, ...chars.slice(0, split), ...(rest ? [rest] : []));
                    i--;
                    continue;
                }
            }
            // Kinsoku: closing punctuation must not begin a line. A single
            // trailing closer hangs past maxWidth (burasage); a longer closer run
            // moves to the next line with its preceding character (oidashi).
            if (isKinsokuNoStartChar(unit)) {
                let closerEnd = i + 1;
                while (closerEnd < units.length) {
                    const next = units[closerEnd];
                    if (next === undefined || !isKinsokuNoStartChar(next))
                        break;
                    closerEnd++;
                }
                if (closerEnd === i + 1) {
                    // burasage: at most one closer may hang past maxWidth
                    appendText(unit, unitWidth);
                    flushLine();
                    continue;
                }
                // oidashi: pull the preceding character (stepping over closers
                // already on the line so the next line cannot begin with one).
                const pulled = [];
                while (lineCharCount() > 1) {
                    const piece = pullLastChar();
                    if (!piece)
                        break;
                    pulled.unshift(piece);
                    if (!isKinsokuNoStartChar(piece.text))
                        break;
                }
                // The pull must be anchored by a non-closer, or the next line would
                // still begin with closing punctuation — put the pieces back and
                // hang instead.
                if (pulled.length > 0 && isKinsokuNoStartChar(pulled[0].text)) {
                    for (const piece of pulled) {
                        appendText(piece.text, piece.width);
                    }
                    pulled.length = 0;
                }
                if (pulled.length === 0) {
                    // The line cannot give up its only character: fall back to
                    // hanging the whole closer run
                    appendText(unit, unitWidth);
                    while (i + 1 < closerEnd) {
                        i++;
                        const next = units[i];
                        if (next !== undefined)
                            appendText(next, this.getCachedWidth(ctx, next, font));
                    }
                    flushLine();
                    continue;
                }
                // The pull may have exposed trailing openers; move them down too
                pulled.unshift(...pullTrailingOpeners());
                flushLine();
                seedCarried(pulled);
                appendText(unit, unitWidth);
                continue;
            }
            // Ordinary break: carry any trailing openers onto the next line
            // together with the unit that triggered the break.
            const carried = pullTrailingOpeners();
            flushLine();
            seedCarried(carried);
            appendText(unit, unitWidth);
        }
        // Add remaining text to the working line, mirroring wrapWords
        if (lineText) {
            const lastFrag = workingLine[workingLine.length - 1];
            if (canMergeFragments(lastFrag, sourceFragment) && lastFrag) {
                lastFrag.text += lineText;
                lastFrag.width = (lastFrag.width ?? 0) + lineTextWidth;
            }
            else {
                workingLine.push(createFragment(lineText, lineTextWidth));
            }
        }
        // Push current line as the "working" line
        lines.push({
            fragments: [...workingLine],
            width: lineWidth,
            endsWithNewline: false,
        });
        // Return the updated state
        return { workingLine, lineWidth };
    }
    /**
     * Wraps text at character boundaries (for CJK or when words don't fit).
     * Uses cached character widths and running totals to avoid O(n^2) re-measurement.
     */
    wrapChars(ctx, text, sourceFragment, maxWidth, currentLine, currentLineWidth, lines) {
        const font = sourceFragment.font;
        let lineWidth = currentLineWidth;
        let lineText = '';
        let lineTextWidth = 0;
        let workingLine = [...currentLine]; // Clone to avoid mutating input
        // Helper to create fragment preserving extra properties from source
        const createFragment = (fragText, fragWidth) => {
            const frag = { text: fragText, font, width: fragWidth };
            // Copy extra properties from source fragment (like runProps)
            for (const key of Object.keys(sourceFragment)) {
                if (key !== 'text' && key !== 'font' && key !== 'width') {
                    frag[key] = sourceFragment[key];
                }
            }
            return frag;
        };
        for (const char of text) {
            // Use cached width measurement
            const charWidth = this.getCachedWidth(ctx, char, font);
            if (lineWidth + charWidth <= maxWidth || lineWidth === 0) {
                lineText += char;
                lineTextWidth += charWidth;
                lineWidth += charWidth;
            }
            else {
                // Line is full, push it
                if (lineText) {
                    workingLine.push(createFragment(lineText, lineTextWidth));
                }
                lines.push({
                    fragments: [...workingLine],
                    width: lineWidth,
                    endsWithNewline: false,
                });
                // Start new line - create new array instead of splice
                workingLine = [];
                lineText = char;
                lineTextWidth = charWidth;
                lineWidth = charWidth;
            }
        }
        // Add remaining characters
        if (lineText) {
            workingLine.push(createFragment(lineText, lineTextWidth));
        }
        lines.push({
            fragments: [...workingLine],
            width: lineWidth,
            endsWithNewline: false,
        });
        // Return the updated state
        return { workingLine, lineWidth };
    }
    /**
     * Measures text width using the font resolver.
     */
    measureText(ctx, text, font) {
        return this.fontResolver.measureText(ctx, text, font);
    }
    /**
     * Finds word boundaries in text.
     * Returns array of indices where words start.
     */
    findWordBoundaries(text) {
        const boundaries = [0];
        for (let i = 1; i < text.length; i++) {
            const prevChar = text[i - 1];
            const currChar = text[i];
            if (prevChar === undefined || currChar === undefined)
                continue;
            // Word boundary conditions:
            // - After whitespace
            // - Before/after punctuation
            // - CJK character boundaries
            if ((/\s/.test(prevChar) && !/\s/.test(currChar)) ||
                this.isCjkChar(currChar) ||
                (this.isCjkChar(prevChar) && !this.isCjkChar(currChar))) {
                boundaries.push(i);
            }
        }
        return boundaries;
    }
    /**
     * Checks if a character is a CJK (Chinese, Japanese, Korean) character.
     * Delegates to the module-level classifier shared with run splitting.
     */
    isCjkChar(char) {
        return isCjkChar(char);
    }
    /**
     * Finds potential hyphenation points in a word.
     * Returns array of indices where hyphenation is allowed.
     * Note: This is a simple syllable-based approach; full hyphenation
     * would require a dictionary or hyphenation algorithm like Knuth-Liang.
     */
    findHyphenationPoints(word) {
        const points = [];
        const minLength = 4;
        if (word.length < minLength) {
            return points;
        }
        // Simple vowel-based syllable detection
        const vowels = 'aeiouAEIOU';
        for (let i = 2; i < word.length - 2; i++) {
            const char = word[i];
            const prevChar = word[i - 1];
            if (char && prevChar && vowels.includes(prevChar) && !vowels.includes(char)) {
                points.push(i);
            }
        }
        return points;
    }
}
/**
 * Creates a WordWrapper instance.
 */
export function createWordWrapper(fontResolver, logger) {
    return new WordWrapper({ fontResolver, logger });
}
//# sourceMappingURL=WordWrapper.js.map