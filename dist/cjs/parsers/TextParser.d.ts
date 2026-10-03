/**
 * Parses text body (txBody) elements from shape XML.
 * Extracts paragraphs, runs, and their properties for rendering.
 */
import type { TextBody, TextBodyProperties, TextRunProperties } from '../types/elements.js';
import type { ResolvedTheme, ColorMap } from '../types/theme.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Text run properties extended with the per-script typefaces parsed from
 * a:ea and a:cs. Kept separate from the base TextRunProperties so runs can
 * be split at script boundaries with the correct typeface per segment.
 */
export interface ScriptAwareRunProperties extends TextRunProperties {
    /** East Asian (a:ea) typeface; may be a theme reference (+mj-ea/+mn-ea) */
    eastAsianFontFamily?: string;
    /** Complex script (a:cs) typeface; may be a theme reference (+mj-cs/+mn-cs) */
    complexScriptFontFamily?: string;
}
/**
 * Configuration for TextParser.
 */
export interface TextParserConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /**
     * 1-based number of the slide being rendered. When set, a:fld fields of
     * type "slidenum" resolve to this number instead of their cached a:t text
     * (which is just the value stored at authoring time). Other field types
     * (datetime variants, ...) always keep their cached text, matching
     * PowerPoint's own export behavior.
     */
    slideNumber?: number;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Text body properties extended with parsed a:normAutofit values.
 *
 * PowerPoint stores the result of its shrink-to-fit pass in the file; when
 * present these are honored as-is (no re-solving) by TextLayoutEngine. An
 * empty a:normAutofit (both attributes absent) is instead solved iteratively
 * at render time (see AutofitSolver), matching PowerPoint's behavior.
 */
export interface ParsedTextBodyProperties extends TextBodyProperties {
    /** normAutofit font scale as a fraction (fontScale="62500" -> 0.625) */
    fontScale?: number;
    /** normAutofit line-spacing reduction as a fraction (lnSpcReduction="20000" -> 0.2) */
    lineSpaceReduction?: number;
}
/**
 * Parses text body elements from slide XML.
 */
export declare class TextParser {
    private readonly logger;
    private readonly colorResolver;
    private readonly slideNumber?;
    constructor(config: TextParserConfig);
    /**
     * Parses a text body (txBody) element.
     *
     * @param txBody Text body XML node
     * @param inheritedListStyles Optional inherited list-style source nodes
     *   (each containing a:lvl1pPr..a:lvl9pPr, e.g. the master p:txStyles
     *   bucket and the layout placeholder's a:lstStyle), ordered lowest
     *   precedence first. The shape's own a:lstStyle and each paragraph's pPr
     *   are merged on top of this chain.
     * @returns Parsed TextBody or undefined if invalid
     */
    parseTextBody(txBody: PptxXmlNode | undefined, inheritedListStyles?: PptxXmlNode[]): TextBody | undefined;
    /**
     * Parses body properties (a:bodyPr).
     */
    private parseBodyProperties;
    /**
     * Parses a percentage attribute into a fraction. Handles both the
     * 1000ths-of-a-percent integer form ("62500" -> 0.625) and the transitional
     * percent-string form ("62.5%" -> 0.625).
     */
    private parsePercentAttr;
    /**
     * Parses a paragraph (a:p) element.
     */
    private parseParagraph;
    /**
     * Parses a line break (a:br) into a newline run.
     * Run properties are preserved so paragraphs consisting only of breaks
     * keep their formatting-derived line height.
     */
    private parseLineBreak;
    /**
     * Parses paragraph properties (a:pPr), merging the full inheritance chain
     * per ECMA-376: inherited list styles (master txStyles bucket, layout
     * placeholder lstStyle) < shape's own lstStyle < explicit pPr, all at the
     * paragraph's indent level.
     */
    private parseParagraphProperties;
    /**
     * Parses a single pPr-like node (a:pPr or a:lvlNpPr \u2014 they share the same
     * schema) into sparse paragraph properties with no defaults applied.
     */
    private parsePPrNode;
    /**
     * Merges sparse paragraph properties, with `over` taking precedence over
     * `base` property by property.
     */
    private mergeSparseParagraphProperties;
    /**
     * Parses the bullet aspects present on a pPr-like node. Returns undefined
     * when the node specifies nothing bullet-related.
     */
    private parseBulletProps;
    /**
     * Merges bullet aspects, with `over` taking precedence. The bullet type
     * (with its type-specific fields) comes from the most specific source that
     * declares one; color, size, and font inherit independently.
     */
    private mergeBulletProps;
    /**
     * Converts merged sparse bullet aspects into a final BulletConfig.
     * Aspects without a resolved type (e.g. an inherited color with no bullet
     * anywhere in the chain) yield no bullet.
     */
    private finalizeBullet;
    /**
     * Parses a text run (a:r).
     */
    private parseTextRun;
    /**
     * Merges run properties, with primary overriding defaults.
     */
    private mergeRunProperties;
    /**
     * Parses a text field (a:fld), ECMA-376 §21.1.2.2.4.
     *
     * Field text is resolved type-aware: "slidenum" fields yield the current
     * slide number (when known) because the cached a:t only holds the number
     * the authoring application last saw. All other types — notably the
     * datetime variants — keep their cached literal text: PowerPoint's own
     * image export shows the stored value rather than substituting the
     * render-time date.
     */
    private parseTextField;
    /**
     * Splits a run at script boundaries so each emitted run can be rendered
     * with a single font: CJK segments use the East Asian typeface and
     * complex-script segments (Arabic/Hebrew) the complex-script typeface.
     * Runs without EA/CS characters are returned unchanged (minus the
     * per-script metadata) so pure-Latin content parses exactly as before.
     */
    private splitRunByScript;
    /**
     * Returns run properties with fontFamily switched to the segment's script
     * font. EA/CS segments without an explicit typeface anywhere in the style
     * chain fall back to the theme's minor East Asian / complex-script font
     * (resolved by FontResolver, which also handles +mj-ea/+mn-ea/+mj-cs/
     * +mn-cs references).
     */
    private applyScriptFont;
    /**
     * Removes per-script typeface metadata from run properties. Runs without
     * EA/CS characters are emitted without the extra fields so their parsed
     * output is identical to pre-split behavior.
     */
    private stripScriptFonts;
    /**
     * Parses run properties (a:rPr).
     */
    private parseRunProperties;
    /**
     * Parses spacing value (a:spcBef, a:spcAft).
     *
     * Encoding contract (consumed by TextLayoutEngine.resolveParagraphSpacingPx):
     * - a:spcPts (absolute spacing) is returned as a POSITIVE value in EMU
     * - a:spcPct (percentage of line height) is returned as a NEGATIVE value in
     *   1000ths of a percent (-50000 = 50%)
     *
     * PowerPoint renders a:spcPts paragraph spacing quantized to whole points
     * (a legacy of the binary format storing it as integer points): measured
     * against PowerPoint exports, spcPts val="133" (1.33pt) advances exactly
     * 1pt per gap — using the fractional value made stacked one-line
     * paragraphs drift ~2.5% per paragraph. Values are therefore rounded to
     * the nearest whole point before converting to EMU.
     */
    private parseSpacing;
    /**
     * Parses line spacing (a:lnSpc).
     *
     * Encoding contract (consumed by TextLayoutEngine.resolveLineHeightPx):
     * - a:spcPct (percentage) is returned as a POSITIVE value in 1000ths of a
     *   percent (100000 = 100%)
     * - a:spcPts (fixed line height) is returned as a NEGATIVE value in
     *   hundredths of a point (-2400 = fixed 24pt)
     */
    private parseLineSpacing;
    /**
     * Extracts text content from a text node.
     */
    private extractTextContent;
    /**
     * Parses an integer attribute.
     */
    private parseIntAttr;
    /**
     * Gets a child element, preserving the presence of empty elements. The XML
     * parser yields self-closed childless elements (e.g. `<a:buNone/>`,
     * `<a:noAutofit/>`) as empty strings, which must not be confused with a
     * missing element.
     */
    private getPresentChild;
}
/**
 * Creates a TextParser instance.
 */
export declare function createTextParser(theme: ResolvedTheme, logger?: ILogger): TextParser;
