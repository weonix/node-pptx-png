/**
 * Text layout engine for measuring and positioning text within bounds.
 * Handles paragraph properties, alignment, spacing, and indentation.
 */
import { FontResolver, pointsToPixels, } from './FontResolver.js';
import { WordWrapper } from './WordWrapper.js';
import { BulletFormatter } from './BulletFormatter.js';
import { AUTOFIT_CANDIDATES, solveAutofitScale } from './AutofitSolver.js';
import { emuToPixels, fontSizeToPoints } from '../core/UnitConverter.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Resolves the line advance in pixels from a parsed line-spacing value.
 *
 * TextParser encodes a:lnSpc as a single number: positive values are
 * percentages in 1000ths of a percent (a:spcPct, 100000 = 100%), negative
 * values are fixed spacing in hundredths of a point (a:spcPts).
 *
 * @param lineSpacing Encoded line spacing (undefined = single spacing)
 * @param singleLineHeightPx Single-spaced line height in pixels
 * @param scaleY Vertical render scale factor
 * @returns Line height in pixels
 */
export function resolveLineHeightPx(lineSpacing, singleLineHeightPx, scaleY) {
    if (lineSpacing === undefined) {
        return singleLineHeightPx;
    }
    if (lineSpacing < 0) {
        // Fixed spacing: -lineSpacing is in hundredths of a point
        return pointsToPixels(-lineSpacing / 100) * scaleY;
    }
    // Percentage of the single-spaced line height (100000 = 100%)
    return singleLineHeightPx * (lineSpacing / 100000);
}
/**
 * Resolves paragraph spacing (space before/after) in pixels.
 *
 * TextParser encodes a:spcBef / a:spcAft as a single number: positive values
 * are absolute spacing in EMU (a:spcPts), negative values are percentages of
 * the line height in 1000ths of a percent (a:spcPct, -50000 = 50%).
 *
 * @param spacing Encoded spacing value (undefined = no spacing)
 * @param singleLineHeightPx Single-spaced line height in pixels
 * @param scaleY Vertical render scale factor
 * @returns Spacing in pixels
 */
export function resolveParagraphSpacingPx(spacing, singleLineHeightPx, scaleY) {
    if (spacing === undefined) {
        return 0;
    }
    if (spacing < 0) {
        // Percentage of the line height (-100000 = 100%)
        return (-spacing / 100000) * singleLineHeightPx;
    }
    // Absolute spacing in EMU
    return emuToPixels(spacing) * scaleY;
}
/**
 * Default font size in hundredths of a point. Last resort only: placeholder
 * text normally receives its size from the master/layout inheritance chain.
 */
const DEFAULT_FONT_SIZE = 1800; // 18pt
/**
 * Default insets in EMU.
 */
/** Default left/right text inset (ECMA-376 §21.1.2.1.1 lIns/rIns): 0.1" */
const DEFAULT_INSET_EMU = 91440;
/** Default top/bottom text inset (ECMA-376 §21.1.2.1.1 tIns/bIns): 0.05" */
const DEFAULT_VERTICAL_INSET_EMU = 45720;
/**
 * Default left margin for bulleted paragraphs that declare no marL of their
 * own, in native (unscaled) pixels: 0.25 inch at 96 DPI.
 */
const DEFAULT_BULLET_MARGIN_PX = 24;
/**
 * Characters requiring cross-character shaping (contextual forms, combining
 * marks, conjuncts, ZWJ sequences): Arabic/Syriac/Thaana through Devanagari
 * and the Indic block range, Thai/Lao/Tibetan/Myanmar/Khmer, combining
 * diacritics, presentation forms, and ZWJ/ZWNJ.
 */
// ZWJ/ZWNJ are deliberately matched as standalone controls (shaping
// triggers), not as part of a grapheme.
const COMPLEX_SHAPING_REGEX = 
// eslint-disable-next-line no-misleading-character-class
/[\u0300-\u036F\u0590-\u08FF\u0900-\u0DFF\u0E00-\u0EFF\u0F00-\u0FFF\u1000-\u109F\u1780-\u17FF\u200C\u200D\uFB1D-\uFDFF\uFE70-\uFEFF]/;
/**
 * Font-size multiplier for the reduced capitals of cap="small" text.
 * PowerPoint renders small caps at roughly 80% of the full capital height.
 */
const SMALL_CAPS_SCALE = 0.8;
/**
 * Fraction of a line box occupied by the ascent when a font reports no
 * usable metrics (matches the 0.8em ascent fallback in FontResolver).
 */
const DEFAULT_ASCENT_RATIO = 0.8;
/**
 * Splits text into segments for cap="small" rendering: characters with a
 * distinct uppercase form (i.e. lowercase letters) become smallCap segments
 * carrying their uppercased text, drawn at SMALL_CAPS_SCALE of the font
 * size; all other characters (capitals, digits, punctuation, whitespace)
 * stay full-size. Consecutive characters of the same class are grouped.
 * Uses locale-independent Unicode default case conversion.
 *
 * @param text Run text to segment
 * @returns Ordered segments whose concatenated text is the transformed run
 */
export function segmentSmallCaps(text) {
    const segments = [];
    for (const char of text) {
        const upper = char.toUpperCase();
        const smallCap = upper !== char;
        const emit = smallCap ? upper : char;
        const last = segments[segments.length - 1];
        if (last?.smallCap === smallCap) {
            last.text += emit;
        }
        else {
            segments.push({ text: emit, smallCap });
        }
    }
    return segments;
}
/**
 * Text layout engine for measuring and positioning text.
 */
export class TextLayoutEngine {
    logger;
    fontResolver;
    wordWrapper;
    bulletFormatter;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'TextLayoutEngine');
        this.fontResolver = new FontResolver({
            fontScheme: config.fontScheme,
            logger: this.logger.child?.('FontResolver'),
            warnings: config.warnings,
        });
        this.wordWrapper = new WordWrapper({
            fontResolver: this.fontResolver,
            logger: this.logger.child?.('WordWrapper'),
        });
        this.bulletFormatter = new BulletFormatter({
            logger: this.logger.child?.('BulletFormatter'),
        });
    }
    /**
     * Lays out text within the specified bounds.
     *
     * A text body with an empty a:normAutofit (no stored fontScale /
     * lnSpcReduction) is solved iteratively: the text is laid out at
     * decreasing font scales until it fits the box, matching PowerPoint's
     * render-time shrink-to-fit. Stored values, spAutoFit and noAutofit are
     * honored as before.
     *
     * @param ctx Canvas 2D context for text measurement
     * @param textBody Text body to lay out
     * @param shapeBounds Shape bounds in pixels
     * @param scaleX Horizontal scale factor
     * @param scaleY Vertical scale factor
     * @returns Complete text layout
     */
    layoutText(ctx, textBody, shapeBounds, scaleX, scaleY) {
        const bodyProps = textBody.bodyProperties ?? {};
        // Calculate text area with insets
        const textBounds = this.calculateTextBounds(shapeBounds, bodyProps, scaleX, scaleY);
        // Determine wrap mode
        const wrapMode = bodyProps.wrap === false ? 'none' : 'word';
        // An empty a:normAutofit means "shrink to fit at render time"
        if (this.needsAutofitSolve(bodyProps)) {
            return this.solveAndLayout(ctx, textBody, textBounds, wrapMode, scaleX, scaleY);
        }
        // Resolve stored normAutofit shrink values (applied to all paragraphs)
        const autofit = this.getAutofitScales(bodyProps);
        return this.layoutWithScales(ctx, textBody, textBounds, wrapMode, scaleX, scaleY, autofit);
    }
    /**
     * Runs one full layout pass (wrapping, positioning, alignment) with the
     * given autofit multipliers applied. Both the stored-value path and each
     * solver candidate go through this exact code path, so a solved scale
     * renders identically to the same scale stored in the file.
     */
    layoutWithScales(ctx, textBody, textBounds, wrapMode, scaleX, scaleY, autofit) {
        const bodyProps = textBody.bodyProperties ?? {};
        // Layout each paragraph
        const lines = [];
        let currentY = textBounds.y;
        let paragraphBulletIndex = 0;
        let trailingSpaceAfter = 0;
        for (let paraIdx = 0; paraIdx < textBody.paragraphs.length; paraIdx++) {
            const paragraph = textBody.paragraphs[paraIdx];
            if (!paragraph)
                continue;
            const paraProps = paragraph.properties ?? {};
            // Resolve the paragraph's default font and line metrics up front so
            // spacing percentages can be computed against the single-spaced line height
            const lineMetrics = this.getParagraphLineMetrics(ctx, paraProps, scaleY, autofit.fontScale);
            // Get space before
            const spaceBefore = resolveParagraphSpacingPx(paraProps.spaceBefore, lineMetrics.singleLineHeight, scaleY);
            if (paraIdx > 0) {
                currentY += spaceBefore;
            }
            // Layout paragraph
            const paragraphLines = this.layoutParagraph(ctx, paragraph, textBounds, wrapMode, paragraphBulletIndex, scaleX, scaleY, lineMetrics, autofit);
            // Position lines
            for (let lineIdx = 0; lineIdx < paragraphLines.length; lineIdx++) {
                const line = paragraphLines[lineIdx];
                if (!line)
                    continue;
                // Apply horizontal alignment
                const alignedLine = this.applyHorizontalAlignment(line, textBounds, paraProps.alignment ?? 'left');
                // Calculate the Y offset to convert from paragraph-relative to absolute
                const yOffset = currentY - line.y;
                // Update line and all runs/bullets to absolute Y positions
                alignedLine.y = currentY;
                for (const run of alignedLine.runs) {
                    run.y += yOffset;
                }
                if (alignedLine.bullet) {
                    alignedLine.bullet.y += yOffset;
                }
                lines.push(alignedLine);
                currentY += alignedLine.height;
            }
            // Get space after
            const spaceAfter = resolveParagraphSpacingPx(paraProps.spaceAfter, lineMetrics.singleLineHeight, scaleY);
            currentY += spaceAfter;
            trailingSpaceAfter = spaceAfter;
            // Track bullet index for numbered lists
            if (paraProps.bullet && paraProps.bullet.type !== 'none') {
                paragraphBulletIndex++;
            }
            else {
                paragraphBulletIndex = 0;
            }
        }
        // Apply vertical alignment. The final paragraph's space-after has no
        // following paragraph to push, so PowerPoint excludes it from the
        // anchored content height (it neither lowers centered text nor lifts
        // bottom-anchored text).
        const totalTextHeight = Math.max(0, currentY - textBounds.y - trailingSpaceAfter);
        const verticalOffset = this.calculateVerticalOffset(totalTextHeight, textBounds.height, bodyProps.anchor ?? 'top');
        // Shift all lines by vertical offset
        for (const line of lines) {
            line.y += verticalOffset;
            for (const run of line.runs) {
                run.y += verticalOffset;
            }
            if (line.bullet) {
                line.bullet.y += verticalOffset;
            }
        }
        const maxWidth = Math.max(0, ...lines.map((l) => l.width));
        this.logger.debug('Laid out text', {
            paragraphCount: textBody.paragraphs.length,
            lineCount: lines.length,
            totalHeight: totalTextHeight,
            maxWidth,
        });
        return {
            lines,
            totalHeight: totalTextHeight,
            maxWidth,
            bounds: textBounds,
        };
    }
    /**
     * Calculates text bounds with insets applied.
     */
    calculateTextBounds(shapeBounds, bodyProps, scaleX, scaleY) {
        const leftInset = emuToPixels(bodyProps.leftInset ?? DEFAULT_INSET_EMU) * scaleX;
        const rightInset = emuToPixels(bodyProps.rightInset ?? DEFAULT_INSET_EMU) * scaleX;
        const topInset = emuToPixels(bodyProps.topInset ?? DEFAULT_VERTICAL_INSET_EMU) * scaleY;
        const bottomInset = emuToPixels(bodyProps.bottomInset ?? DEFAULT_VERTICAL_INSET_EMU) * scaleY;
        return {
            x: shapeBounds.x + leftInset,
            y: shapeBounds.y + topInset,
            width: Math.max(0, shapeBounds.width - leftInset - rightInset),
            height: Math.max(0, shapeBounds.height - topInset - bottomInset),
        };
    }
    /**
     * Lays out a single paragraph.
     */
    layoutParagraph(ctx, paragraph, textBounds, wrapMode, bulletIndex, scaleX, scaleY, lineMetrics, autofit) {
        const paraProps = paragraph.properties ?? {};
        const { defaultFont, metrics, singleLineHeight } = lineMetrics;
        // Calculate indentation. A resolved marL (explicit pPr attribute or
        // inherited from the level's lvlNpPr in the style chain) IS the
        // paragraph's complete left margin — PowerPoint never stacks a
        // per-level offset on top of it: deeper levels indent further only
        // because their lvlNpPr declares a larger marL. The synthesized
        // level*457200 EMU offset is a fallback for paragraphs whose style
        // chain never resolves a margin (matching the default text style
        // PowerPoint ships, marL = 457200 EMU per level).
        const level = paraProps.level ?? 0;
        const marginLeft = emuToPixels(paraProps.marginLeft ?? 0) * scaleX;
        const indent = emuToPixels(paraProps.indent ?? 0) * scaleX;
        const bulletIndent = paraProps.marginLeft !== undefined ? 0 : this.bulletFormatter.calculateBulletIndent(level);
        // Get bullet
        const bulletProps = this.convertBulletConfig(paraProps.bullet);
        const bullet = this.bulletFormatter.formatBullet(bulletProps, bulletIndex, level);
        // PowerPoint hangs the bullet inside the paragraph's left margin. Lists
        // normally receive marL from the master text styles via the inheritance
        // chain; when nothing in the chain provides one, give the paragraph a
        // default margin so the bullet never overlaps the first glyph.
        let totalLeftMargin = marginLeft + emuToPixels(bulletIndent) * scaleX;
        if (bullet && totalLeftMargin <= 0) {
            totalLeftMargin = DEFAULT_BULLET_MARGIN_PX * scaleX;
        }
        // For bulleted paragraphs the first-line indent zone holds the bullet
        // (hanging indent); the text itself starts at the left margin
        const firstLineIndent = bullet ? 0 : indent;
        // Build text fragments from runs
        const fragments = this.buildFragments(paragraph, paraProps, scaleY, autofit.fontScale);
        // Calculate line height in pixels from the encoded line spacing.
        // normAutofit lnSpcReduction shrinks spacing-derived advances (single and
        // percentage line spacing) but not fixed a:spcPts line heights.
        let lineHeight = resolveLineHeightPx(paraProps.lineSpacing, singleLineHeight, scaleY);
        if (autofit.lineSpaceReduction > 0 &&
            (paraProps.lineSpacing === undefined || paraProps.lineSpacing >= 0)) {
            lineHeight *= Math.max(0, 1 - autofit.lineSpaceReduction);
        }
        // Calculate available width for text
        const availableWidth = textBounds.width - totalLeftMargin - Math.max(0, firstLineIndent);
        // Empty paragraphs (endParaRPr only) occupy one blank line in
        // PowerPoint, sized by the endParaRPr font when present
        if (fragments.length === 0) {
            const endProps = paragraph.endParaRunProperties;
            let emptySingle = singleLineHeight;
            if (endProps?.fontSize !== undefined) {
                const endFont = this.fontResolver.resolveFont(endProps.fontFamily ?? defaultFont.family, fontSizeToPoints(endProps.fontSize) * autofit.fontScale, endProps.bold ?? false, endProps.italic ?? false, scaleY);
                emptySingle = this.fontResolver.calculateLineHeight(endFont.sizePx);
            }
            let emptyHeight = resolveLineHeightPx(paraProps.lineSpacing, emptySingle, scaleY);
            if (autofit.lineSpaceReduction > 0 &&
                (paraProps.lineSpacing === undefined || paraProps.lineSpacing >= 0)) {
                emptyHeight *= Math.max(0, 1 - autofit.lineSpaceReduction);
            }
            return [{ runs: [], y: 0, height: emptyHeight, width: 0 }];
        }
        // Wrap text
        const wrapResult = this.wordWrapper.wrapText(ctx, fragments, availableWidth, wrapMode, lineHeight);
        // Convert wrapped lines to layout lines
        const layoutLines = [];
        let currentY = 0;
        for (let lineIdx = 0; lineIdx < wrapResult.lines.length; lineIdx++) {
            const wrappedLine = wrapResult.lines[lineIdx];
            if (!wrappedLine)
                continue;
            // Per-line metrics from the fragments actually on the line
            // (PowerPoint: line pitch follows the largest run on the line, not
            // the paragraph's default font — e.g. 6pt runs in a paragraph whose
            // defRPr is 18pt must produce 6pt-pitched lines). The single-spaced
            // pitch is PowerPoint's 1.2x-font-size rule; the baseline is the
            // font's ascent scaled proportionally into that 1.2em line box.
            let lineSingle = 0;
            let lineAscent = 0;
            for (const fragment of wrappedLine.fragments) {
                const fm = this.fontResolver.getFontMetrics(ctx, fragment.font);
                const fh = this.fontResolver.calculateLineHeight(fragment.font.sizePx);
                const fa = fm.lineHeight > 0 ? fm.ascent * (fh / fm.lineHeight) : fh * DEFAULT_ASCENT_RATIO;
                if (fh > lineSingle)
                    lineSingle = fh;
                if (fa > lineAscent)
                    lineAscent = fa;
            }
            if (lineSingle <= 0) {
                lineSingle = singleLineHeight;
                lineAscent =
                    metrics.lineHeight > 0
                        ? metrics.ascent * (singleLineHeight / metrics.lineHeight)
                        : singleLineHeight * DEFAULT_ASCENT_RATIO;
            }
            let perLineHeight = resolveLineHeightPx(paraProps.lineSpacing, lineSingle, scaleY);
            if (autofit.lineSpaceReduction > 0 &&
                (paraProps.lineSpacing === undefined || paraProps.lineSpacing >= 0)) {
                perLineHeight *= Math.max(0, 1 - autofit.lineSpaceReduction);
            }
            // Percentage line spacing scales the whole line box, ascent
            // included: at a:spcPct 90% PowerPoint draws the first baseline 10%
            // higher, at 200% the extra leading lands ABOVE the baseline. Fixed
            // a:spcPts spacing keeps the font's natural ascent.
            if ((paraProps.lineSpacing === undefined || paraProps.lineSpacing >= 0) && lineSingle > 0) {
                lineAscent *= perLineHeight / lineSingle;
            }
            // Calculate X position (first line has different indent)
            const isFirstLine = lineIdx === 0;
            const lineIndent = isFirstLine ? firstLineIndent : 0;
            let xPos = textBounds.x + totalLeftMargin + lineIndent;
            // Position text runs
            const positionedRuns = [];
            for (const fragment of wrappedLine.fragments) {
                const runProps = fragment.runProps;
                // Leave the color undefined when no chain source specifies one so the
                // renderer can fall back to the shape's default text color
                const color = runProps?.color;
                // Use pre-computed width from layout, fallback to measuring only if needed
                const fragmentWidth = fragment.width ?? this.fontResolver.measureText(ctx, fragment.text, fragment.font);
                positionedRuns.push({
                    text: fragment.text,
                    x: xPos,
                    y: currentY + lineAscent,
                    width: fragmentWidth,
                    font: fragment.font,
                    color,
                    underline: runProps?.underline,
                    strikethrough: runProps?.strikethrough,
                    baselineOffset: runProps?.baseline,
                });
                xPos += fragmentWidth;
            }
            const isLastLine = lineIdx === wrapResult.lines.length - 1;
            // PowerPoint ignores trailing whitespace when aligning a SOFT-wrapped
            // line: the space at the wrap point stays at the end of the line,
            // hanging invisibly past the wrap edge, so centered/right-aligned
            // lines are positioned by their ink width only. Hard-ended lines
            // (paragraph end or a:br) keep typed trailing spaces in their
            // aligned width — PowerPoint centers those including the space.
            let trailingWhitespaceWidth = 0;
            if (!isLastLine && !wrappedLine.endsWithNewline) {
                for (let runIdx = positionedRuns.length - 1; runIdx >= 0; runIdx--) {
                    const run = positionedRuns[runIdx];
                    const match = /\s+$/.exec(run.text);
                    if (!match)
                        break;
                    if (match[0].length === run.text.length) {
                        // Whole run is whitespace; keep walking into the previous run
                        trailingWhitespaceWidth += run.width;
                    }
                    else {
                        trailingWhitespaceWidth += this.fontResolver.measureText(ctx, match[0], run.font);
                        break;
                    }
                }
            }
            // Justified alignment stretches lines to the full wrap width by
            // widening word gaps (PowerPoint behavior): every line except the
            // paragraph's last and lines ending in a hard break (a:br). The
            // 'distributed' variant stretches the final line too.
            let lineWidth = Math.max(0, wrappedLine.width - trailingWhitespaceWidth);
            if ((paraProps.alignment === 'justify' && !isLastLine && !wrappedLine.endsWithNewline) ||
                paraProps.alignment === 'distributed') {
                const targetWidth = Math.max(0, availableWidth - Math.min(0, lineIndent));
                if (this.justifyLine(ctx, positionedRuns, targetWidth)) {
                    lineWidth = targetWidth;
                }
            }
            // Add bullet to first line
            let positionedBullet;
            if (isFirstLine && bullet) {
                const bulletFont = this.fontResolver.resolveFont(bullet.font ?? defaultFont.family, defaultFont.sizePoints * bullet.sizeMultiplier, defaultFont.bold, defaultFont.italic, scaleY);
                // Bullet color: use explicit bullet color, or inherit from first text run, or default text color
                const firstRunColor = positionedRuns[0]?.color;
                const bulletColor = bullet.color ?? firstRunColor;
                // Position the bullet in the hanging-indent zone: PowerPoint places it
                // at marL + indent (indent is negative for hanging indents). When no
                // hanging indent is defined, back off by the bullet's own width plus a
                // space so the bullet does not overlap the text. Clamp to the text
                // bounds so the bullet never renders outside the shape.
                const textStartX = textBounds.x + totalLeftMargin;
                const bulletGap = this.fontResolver.measureText(ctx, `${bullet.text} `, bulletFont);
                const hangingIndent = indent < 0 ? -indent : bulletGap;
                positionedBullet = {
                    text: bullet.text,
                    x: Math.max(textBounds.x, textStartX - hangingIndent),
                    y: currentY + lineAscent,
                    font: bulletFont,
                    color: bulletColor,
                };
            }
            layoutLines.push({
                runs: positionedRuns,
                bullet: positionedBullet,
                y: currentY,
                height: perLineHeight,
                width: lineWidth,
            });
            currentY += perLineHeight;
        }
        return layoutLines;
    }
    /**
     * Stretches a positioned line to fill targetWidth by distributing the
     * shortfall evenly across its word gaps (maximal whitespace stretches
     * with words on both sides, including gaps spanning run boundaries).
     * The runs array is re-split in place into word- and gap-level runs so
     * each piece carries its own x position; gap runs keep their source
     * run's decorations so underlines span the widened gaps.
     * @returns true when the line was stretched to targetWidth
     */
    justifyLine(ctx, runs, targetWidth) {
        const startX = runs[0]?.x;
        if (startX === undefined)
            return false;
        const pieces = [];
        for (const run of runs) {
            for (const token of run.text.split(/(\s+)/)) {
                if (token === '')
                    continue;
                pieces.push({
                    source: run,
                    text: token,
                    isSpace: /^\s+$/.test(token),
                    width: this.fontResolver.measureText(ctx, token, run.font),
                    gapExtra: 0,
                });
            }
        }
        // PowerPoint ignores whitespace at the edges of a justified line: drop
        // leading/trailing space pieces so they neither shift the first word
        // nor count against the stretch amount
        while (pieces[0]?.isSpace)
            pieces.shift();
        while (pieces[pieces.length - 1]?.isSpace)
            pieces.pop();
        // A gap is a maximal whitespace stretch with words on both sides;
        // consecutive space pieces from adjacent runs form ONE gap
        const gapEnds = [];
        for (let i = 0; i < pieces.length; i++) {
            if (!pieces[i].isSpace)
                continue;
            const gapStart = i;
            while (i + 1 < pieces.length && pieces[i + 1].isSpace)
                i++;
            if (gapStart > 0 && i < pieces.length - 1) {
                gapEnds.push(i);
            }
        }
        if (gapEnds.length === 0)
            return false;
        const naturalWidth = pieces.reduce((sum, p) => sum + p.width, 0);
        const extra = targetWidth - naturalWidth;
        if (extra <= 0)
            return false;
        const extraPerGap = extra / gapEnds.length;
        for (const end of gapEnds) {
            pieces[end].gapExtra = extraPerGap;
        }
        // Rebuild the runs at their stretched positions
        runs.length = 0;
        let cursor = startX;
        for (const piece of pieces) {
            const width = piece.width + piece.gapExtra;
            runs.push({
                text: piece.text,
                x: cursor,
                y: piece.source.y,
                width,
                font: piece.source.font,
                color: piece.source.color,
                underline: piece.source.underline,
                strikethrough: piece.source.strikethrough,
                baselineOffset: piece.source.baselineOffset,
            });
            cursor += width;
        }
        return true;
    }
    /**
     * Builds text fragments from paragraph runs.
     * Fonts are resolved with the render scale so measurement, wrapping and
     * rendering all use the same scaled glyph size; the normAutofit fontScale
     * multiplier is applied to every resolved size.
     *
     * Capitalization (a:rPr cap) is applied here, before wrapping, so line
     * breaks are computed against the transformed text: cap="all" uppercases
     * the fragment text; cap="small" splits the run into full-size and
     * reduced-size uppercase segments. Letter tracking (a:rPr spc) rides on
     * the resolved font so measurement and drawing stay in sync. Tracking is
     * kept absolute (not multiplied by the autofit fontScale), matching how
     * PowerPoint stores spc independently of normAutofit font scaling.
     */
    buildFragments(paragraph, paraProps, scaleY, fontScale) {
        const fragments = [];
        for (const run of paragraph.runs) {
            const runProps = run.properties ?? {};
            const mergedProps = this.mergeRunProperties(paraProps.defaultRunProperties, runProps);
            const fontSize = fontSizeToPoints(mergedProps.fontSize ?? DEFAULT_FONT_SIZE) * fontScale;
            // a:spc is in hundredths of a point; convert to points for resolveFont
            const letterSpacingPoints = (mergedProps.spacing ?? 0) / 100;
            // a:rPr kern (hundredths of a point) is the minimum size that still
            // kerns: sizes below the threshold (or an explicit kern="0") lay out
            // on bare advances, so PowerPoint's small print measures wider than
            // shaped text. Without an inherited value kerning stays on. The check
            // uses each resolved size so reduced small-caps segments gate
            // themselves.
            const kerningMin = mergedProps.kerningMin;
            // Per-glyph (unkerned) placement is only valid for simple scripts:
            // Arabic/Indic contextual forms, combining marks, and ZWJ sequences
            // need shaping across characters, so those runs keep the shaped path.
            const needsShaping = COMPLEX_SHAPING_REGEX.test(run.text);
            const kerns = (sizePoints) => needsShaping ||
                kerningMin === undefined ||
                (kerningMin > 0 && sizePoints * 100 >= kerningMin);
            const resolve = (sizePoints) => this.fontResolver.resolveFont(mergedProps.fontFamily, sizePoints, mergedProps.bold ?? false, mergedProps.italic ?? false, scaleY, letterSpacingPoints, kerns(sizePoints));
            if (mergedProps.caps === 'small') {
                // Small caps: lowercase letters render as capitals at a reduced
                // size; everything else keeps the full-size font. Segments share
                // the run's properties and baseline.
                const segments = segmentSmallCaps(run.text);
                for (let i = 0; i < segments.length; i++) {
                    const segment = segments[i];
                    // Continuation segments that start mid-word must not become line
                    // break opportunities ('Memorandum' -> 'M' + 'EMORANDUM' is one
                    // word); the wrapper glues noBreakBefore fragments to the line.
                    const startsMidWord = i > 0 && !/^\s/.test(segment.text) && !/\s$/.test(segments[i - 1].text);
                    fragments.push({
                        text: segment.text,
                        font: resolve(segment.smallCap ? fontSize * SMALL_CAPS_SCALE : fontSize),
                        runProps: mergedProps,
                        ...(startsMidWord ? { noBreakBefore: true } : {}),
                    });
                }
                continue;
            }
            // cap="all": uppercase before wrapping so measurement sees the
            // transformed text (locale-independent case conversion)
            const text = mergedProps.caps === 'all' ? run.text.toUpperCase() : run.text;
            fragments.push({
                text,
                font: resolve(fontSize),
                runProps: mergedProps,
            });
        }
        return fragments;
    }
    /**
     * Merges default run properties with specific run properties.
     */
    mergeRunProperties(defaults, specific) {
        return {
            fontSize: specific?.fontSize ?? defaults?.fontSize,
            fontFamily: specific?.fontFamily ?? defaults?.fontFamily,
            bold: specific?.bold ?? defaults?.bold,
            italic: specific?.italic ?? defaults?.italic,
            underline: specific?.underline ?? defaults?.underline,
            strikethrough: specific?.strikethrough ?? defaults?.strikethrough,
            color: specific?.color ?? defaults?.color,
            baseline: specific?.baseline ?? defaults?.baseline,
            spacing: specific?.spacing ?? defaults?.spacing,
            caps: specific?.caps ?? defaults?.caps,
            kerningMin: specific?.kerningMin ?? defaults?.kerningMin,
        };
    }
    /**
     * Converts BulletConfig to BulletProps.
     */
    convertBulletConfig(config) {
        if (!config || config.type === 'none') {
            return undefined;
        }
        return {
            type: config.type === 'auto' ? 'autoNum' : config.type === 'picture' ? 'blip' : config.type,
            char: config.char,
            autoNumType: config.autoNumType,
            startAt: config.startAt,
            color: config.color,
            sizePercent: config.sizePercent,
            font: config.font,
        };
    }
    /**
     * Whether the text body requires a render-time shrink-to-fit solve: an
     * a:normAutofit that carries no stored values. PowerPoint persists solved
     * fontScale/lnSpcReduction on save; when either attribute is present the
     * file already holds the solved state and is honored as-is (stored values
     * win — no re-solving). Only a bare `<a:normAutofit/>` is solved here.
     */
    needsAutofitSolve(bodyProps) {
        const parsed = bodyProps;
        return (parsed.autoFit === 'normal' &&
            parsed.fontScale === undefined &&
            parsed.lineSpaceReduction === undefined);
    }
    /**
     * Iteratively solves shrink-to-fit for an empty a:normAutofit: finds the
     * largest candidate scale (see AutofitSolver for the PowerPoint-observed
     * ladder) whose full layout — including re-wrapping at the scaled font
     * sizes — fits the available text-box height, then returns that layout.
     * Layout passes are cached per candidate so the winner is never recomputed;
     * the solver binary-searches the ladder, costing at most ~6 passes.
     */
    solveAndLayout(ctx, textBody, textBounds, wrapMode, scaleX, scaleY) {
        const layouts = new Map();
        const measure = (candidate, index) => {
            let layout = layouts.get(index);
            if (!layout) {
                layout = this.layoutWithScales(ctx, textBody, textBounds, wrapMode, scaleX, scaleY, {
                    fontScale: candidate.fontScale,
                    lineSpaceReduction: candidate.lineSpaceReduction,
                });
                layouts.set(index, layout);
            }
            return layout.totalHeight;
        };
        const index = solveAutofitScale(textBounds.height, measure);
        const candidate = AUTOFIT_CANDIDATES[index] ?? { fontScale: 1, lineSpaceReduction: 0 };
        // solveAutofitScale always measures the index it returns, so this is a
        // cache hit; the measure() call only fills the cache on a defensive miss.
        measure(candidate, index);
        const layout = layouts.get(index);
        if (candidate.fontScale < 1) {
            this.logger.debug('Solved render-time normAutofit shrink', {
                fontScale: candidate.fontScale,
                lineSpaceReduction: candidate.lineSpaceReduction,
                totalHeight: layout.totalHeight,
                availableHeight: textBounds.height,
            });
        }
        return layout;
    }
    /**
     * Resolves the normAutofit multipliers stored on the text body. Only
     * meaningful values are honored: fontScale must be a positive fraction
     * and lnSpcReduction a fraction below 1.
     */
    getAutofitScales(bodyProps) {
        const parsed = bodyProps;
        if (parsed.autoFit !== 'normal') {
            return { fontScale: 1, lineSpaceReduction: 0 };
        }
        const fontScale = parsed.fontScale !== undefined && parsed.fontScale > 0 ? parsed.fontScale : 1;
        const lineSpaceReduction = parsed.lineSpaceReduction !== undefined &&
            parsed.lineSpaceReduction > 0 &&
            parsed.lineSpaceReduction < 1
            ? parsed.lineSpaceReduction
            : 0;
        return { fontScale, lineSpaceReduction };
    }
    /**
     * Resolves the default font and single-spaced line metrics for a paragraph.
     *
     * PowerPoint bases single (and percentage) line spacing on 1.2x the font
     * point size, NOT on the font's own ascent+descent metrics: an 8pt Open
     * Sans paragraph advances 9.6pt per line even though the font's hhea
     * metrics sum to 1.36em. Using font metrics here made every dense text
     * block drift vertically against PowerPoint output.
     */
    getParagraphLineMetrics(ctx, paraProps, scaleY, fontScale) {
        const defaultFont = this.getDefaultFont(paraProps, scaleY, fontScale);
        const metrics = this.fontResolver.getFontMetrics(ctx, defaultFont);
        const singleLineHeight = this.fontResolver.calculateLineHeight(defaultFont.sizePx);
        return { defaultFont, metrics, singleLineHeight };
    }
    /**
     * Gets the default font for a paragraph, scaled to render pixels and by
     * the normAutofit fontScale multiplier.
     */
    getDefaultFont(paraProps, scaleY, fontScale) {
        const defaultProps = paraProps.defaultRunProperties ?? {};
        const fontSize = fontSizeToPoints(defaultProps.fontSize ?? DEFAULT_FONT_SIZE) * fontScale;
        return this.fontResolver.resolveFont(defaultProps.fontFamily, fontSize, defaultProps.bold ?? false, defaultProps.italic ?? false, scaleY);
    }
    /**
     * Applies horizontal alignment to a layout line.
     */
    applyHorizontalAlignment(line, textBounds, alignment) {
        if (alignment === 'left' || line.runs.length === 0) {
            return line;
        }
        const availableWidth = textBounds.width;
        const contentWidth = line.width;
        let offsetX = 0;
        switch (alignment) {
            case 'center':
                offsetX = (availableWidth - contentWidth) / 2;
                break;
            case 'right':
                offsetX = availableWidth - contentWidth;
                break;
            case 'justify':
            case 'distributed':
                // Word-gap stretching is applied during paragraph layout (see
                // justifyLine — it needs the wrap width and paragraph-relative
                // line position); the line is already left-anchored here
                break;
        }
        if (offsetX !== 0) {
            for (const run of line.runs) {
                run.x += offsetX;
            }
            if (line.bullet) {
                line.bullet.x += offsetX;
            }
        }
        return line;
    }
    /**
     * Calculates vertical offset for alignment.
     */
    calculateVerticalOffset(contentHeight, containerHeight, anchor) {
        switch (anchor) {
            case 'top':
                return 0;
            case 'middle':
                return Math.max(0, (containerHeight - contentHeight) / 2);
            case 'bottom':
                return Math.max(0, containerHeight - contentHeight);
            default:
                return 0;
        }
    }
    /**
     * Gets the font resolver for external use.
     */
    getFontResolver() {
        return this.fontResolver;
    }
    /**
     * Gets the bullet formatter for external use.
     */
    getBulletFormatter() {
        return this.bulletFormatter;
    }
}
/**
 * Creates a TextLayoutEngine instance.
 */
export function createTextLayoutEngine(fontScheme, logger) {
    return new TextLayoutEngine({ fontScheme, logger });
}
//# sourceMappingURL=TextLayoutEngine.js.map