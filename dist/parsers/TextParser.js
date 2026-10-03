/**
 * Parses text body (txBody) elements from shape XML.
 * Extracts paragraphs, runs, and their properties for rendering.
 */
import { getXmlChild, getXmlChildren, getXmlAttr, getChildrenInDocumentOrder, } from '../core/PptxParser.js';
import { ColorResolver } from '../theme/ColorResolver.js';
import { isCjkChar } from '../text/WordWrapper.js';
import { createLogger } from '../utils/Logger.js';
/**
 * Matches characters that need a non-Latin script font: complex scripts
 * (Arabic/Hebrew and neighbors, U+0590-08FF) and the CJK ranges handled by
 * the word wrapper. Used as a cheap guard so pure-Latin runs are emitted
 * exactly as before.
 */
const NON_LATIN_SCRIPT_REGEX = /[\u0590-\u08FF\u1100-\u11FF\u3000-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFB1D-\uFDFF\uFE70-\uFEFF\uFF00-\uFFEF\u{20000}-\u{3FFFF}]/u;
/**
 * Classifies a character for per-script font selection. Neutral characters
 * (ASCII, whitespace, general punctuation) classify as 'latin'. Bidi
 * reordering is out of scope; 'cs' classification only selects the font.
 */
function classifyScriptChar(char) {
    if (isCjkChar(char))
        return 'ea';
    const code = char.codePointAt(0) ?? 0;
    if ((code >= 0x0590 && code <= 0x08ff) ||
        (code >= 0xfb1d && code <= 0xfdff) || // Hebrew/Arabic Presentation Forms A
        (code >= 0xfe70 && code <= 0xfeff) // Arabic Presentation Forms B
    ) {
        return 'cs';
    }
    return 'latin';
}
/**
 * Parses text body elements from slide XML.
 */
export class TextParser {
    logger;
    colorResolver;
    slideNumber;
    constructor(config) {
        this.logger = config.logger ?? createLogger('warn', 'TextParser');
        this.colorResolver = new ColorResolver(config.theme.colors, config.colorMap);
        this.slideNumber = config.slideNumber;
    }
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
    parseTextBody(txBody, inheritedListStyles) {
        if (!txBody)
            return undefined;
        // Parse body properties
        const bodyPr = getXmlChild(txBody, 'a:bodyPr');
        const bodyProperties = this.parseBodyProperties(bodyPr);
        // Parse list style (for default paragraph/run properties)
        const lstStyle = getXmlChild(txBody, 'a:lstStyle');
        // Parse paragraphs
        const pNodes = getXmlChildren(txBody, 'a:p');
        const paragraphs = [];
        for (const pNode of pNodes) {
            const paragraph = this.parseParagraph(pNode, lstStyle, inheritedListStyles);
            paragraphs.push(paragraph);
        }
        // Return undefined if no paragraphs with content
        if (paragraphs.length === 0) {
            return undefined;
        }
        this.logger.debug('Parsed text body', {
            paragraphCount: paragraphs.length,
        });
        return {
            bodyProperties,
            paragraphs,
        };
    }
    /**
     * Parses body properties (a:bodyPr).
     */
    parseBodyProperties(bodyPr) {
        if (!bodyPr) {
            return {};
        }
        // Parse anchor (vertical alignment)
        const anchorAttr = getXmlAttr(bodyPr, 'anchor');
        let anchor;
        if (anchorAttr === 't')
            anchor = 'top';
        else if (anchorAttr === 'ctr')
            anchor = 'middle';
        else if (anchorAttr === 'b')
            anchor = 'bottom';
        // Parse anchor center (horizontal centering)
        const anchorCtr = getXmlAttr(bodyPr, 'anchorCtr') === '1';
        // Parse wrap
        const wrapAttr = getXmlAttr(bodyPr, 'wrap');
        const wrap = wrapAttr !== 'none';
        // Parse insets
        const leftInset = this.parseIntAttr(bodyPr, 'lIns');
        const rightInset = this.parseIntAttr(bodyPr, 'rIns');
        const topInset = this.parseIntAttr(bodyPr, 'tIns');
        const bottomInset = this.parseIntAttr(bodyPr, 'bIns');
        // Parse rotation
        const rotAttr = getXmlAttr(bodyPr, 'rot');
        const rotation = rotAttr ? parseInt(rotAttr, 10) / 60000 : undefined;
        // Parse auto-fit (presence-checked: these elements are often empty)
        const noAutofit = this.getPresentChild(bodyPr, 'a:noAutofit');
        const normAutofit = this.getPresentChild(bodyPr, 'a:normAutofit');
        const spAutofit = this.getPresentChild(bodyPr, 'a:spAutoFit');
        let autoFit;
        if (noAutofit)
            autoFit = 'none';
        else if (normAutofit)
            autoFit = 'normal';
        else if (spAutofit)
            autoFit = 'shape';
        // Parse the stored normAutofit shrink values (ECMA-376 §21.1.2.1.3):
        // fontScale scales every font size, lnSpcReduction reduces line spacing.
        const fontScale = normAutofit ? this.parsePercentAttr(normAutofit, 'fontScale') : undefined;
        const lineSpaceReduction = normAutofit
            ? this.parsePercentAttr(normAutofit, 'lnSpcReduction')
            : undefined;
        return {
            anchor,
            anchorCenter: anchorCtr || undefined,
            wrap,
            leftInset,
            rightInset,
            topInset,
            bottomInset,
            rotation,
            autoFit,
            fontScale,
            lineSpaceReduction,
        };
    }
    /**
     * Parses a percentage attribute into a fraction. Handles both the
     * 1000ths-of-a-percent integer form ("62500" -> 0.625) and the transitional
     * percent-string form ("62.5%" -> 0.625).
     */
    parsePercentAttr(node, attr) {
        const raw = getXmlAttr(node, attr);
        if (raw === undefined)
            return undefined;
        const value = parseFloat(raw);
        if (!Number.isFinite(value))
            return undefined;
        return raw.trimEnd().endsWith('%') ? value / 100 : value / 100000;
    }
    /**
     * Parses a paragraph (a:p) element.
     */
    parseParagraph(pNode, lstStyle, inheritedListStyles) {
        // Parse paragraph properties
        const pPr = getXmlChild(pNode, 'a:pPr');
        const properties = this.parseParagraphProperties(pPr, lstStyle, inheritedListStyles);
        // Parse runs (a:r), fields (a:fld) and line breaks (a:br) in document
        // order. The relative order of these children is significant: a paragraph
        // authored as 'Line1<a:br/>Line2' must yield the break between the runs.
        const runs = [];
        const orderedChildren = getChildrenInDocumentOrder(pNode, 'a:p', ['a:r', 'a:fld', 'a:br']);
        for (const { tagName, node } of orderedChildren) {
            let run;
            switch (tagName) {
                case 'a:r':
                    run = this.parseTextRun(node, properties.defaultRunProperties);
                    break;
                case 'a:fld':
                    run = this.parseTextField(node, properties.defaultRunProperties);
                    break;
                case 'a:br':
                    run = this.parseLineBreak(node, properties.defaultRunProperties);
                    break;
            }
            if (run) {
                // Split at script boundaries so each emitted run renders with a
                // single font (canvas fillText uses one font per call)
                runs.push(...this.splitRunByScript(run));
            }
        }
        // Parse end paragraph run properties
        const endParaRPr = getXmlChild(pNode, 'a:endParaRPr');
        const endParaRunProperties = endParaRPr ? this.parseRunProperties(endParaRPr) : undefined;
        return {
            properties,
            runs,
            endParaRunProperties,
        };
    }
    /**
     * Parses a line break (a:br) into a newline run.
     * Run properties are preserved so paragraphs consisting only of breaks
     * keep their formatting-derived line height.
     */
    parseLineBreak(brNode, defaultProps) {
        const rPr = getXmlChild(brNode, 'a:rPr');
        const runProps = rPr ? this.parseRunProperties(rPr) : undefined;
        const properties = this.mergeRunProperties(runProps, defaultProps);
        return {
            text: '\n',
            properties,
        };
    }
    /**
     * Parses paragraph properties (a:pPr), merging the full inheritance chain
     * per ECMA-376: inherited list styles (master txStyles bucket, layout
     * placeholder lstStyle) < shape's own lstStyle < explicit pPr, all at the
     * paragraph's indent level.
     */
    parseParagraphProperties(pPr, lstStyle, inheritedListStyles) {
        // Parse level (clamped to the 0-8 range lvl1pPr..lvl9pPr covers)
        const lvlAttr = pPr ? getXmlAttr(pPr, 'lvl') : undefined;
        const level = Math.min(Math.max(lvlAttr ? parseInt(lvlAttr, 10) : 0, 0), 8);
        const levelKey = `a:lvl${level + 1}pPr`;
        // Assemble the chain sources, lowest precedence first
        const sources = [
            ...(inheritedListStyles ?? []).map((style) => getXmlChild(style, levelKey)),
            lstStyle ? getXmlChild(lstStyle, levelKey) : undefined,
            pPr,
        ];
        let merged = {};
        for (const source of sources) {
            if (!source)
                continue;
            merged = this.mergeSparseParagraphProperties(merged, this.parsePPrNode(source));
        }
        return {
            alignment: merged.alignment,
            level,
            indent: merged.indent,
            marginLeft: merged.marginLeft,
            marginRight: merged.marginRight,
            spaceBefore: merged.spaceBefore,
            spaceAfter: merged.spaceAfter,
            lineSpacing: merged.lineSpacing,
            bullet: this.finalizeBullet(merged.bullet),
            defaultRunProperties: merged.defaultRunProperties,
        };
    }
    /**
     * Parses a single pPr-like node (a:pPr or a:lvlNpPr \u2014 they share the same
     * schema) into sparse paragraph properties with no defaults applied.
     */
    parsePPrNode(node) {
        // Parse alignment
        const algn = getXmlAttr(node, 'algn');
        let alignment;
        if (algn === 'l')
            alignment = 'left';
        else if (algn === 'ctr')
            alignment = 'center';
        else if (algn === 'r')
            alignment = 'right';
        else if (algn === 'just')
            alignment = 'justify';
        else if (algn === 'dist')
            alignment = 'distributed';
        // Parse indentation
        const indent = this.parseIntAttr(node, 'indent');
        const marginLeft = this.parseIntAttr(node, 'marL');
        const marginRight = this.parseIntAttr(node, 'marR');
        // Parse spacing
        const spaceBefore = this.parseSpacing(getXmlChild(node, 'a:spcBef'));
        const spaceAfter = this.parseSpacing(getXmlChild(node, 'a:spcAft'));
        const lineSpacing = this.parseLineSpacing(getXmlChild(node, 'a:lnSpc'));
        // Parse bullet aspects
        const bullet = this.parseBulletProps(node);
        // Parse default run properties
        const defRPr = getXmlChild(node, 'a:defRPr');
        const defaultRunProperties = defRPr ? this.parseRunProperties(defRPr) : undefined;
        return {
            alignment,
            indent,
            marginLeft,
            marginRight,
            spaceBefore,
            spaceAfter,
            lineSpacing,
            bullet,
            defaultRunProperties,
        };
    }
    /**
     * Merges sparse paragraph properties, with `over` taking precedence over
     * `base` property by property.
     */
    mergeSparseParagraphProperties(base, over) {
        return {
            alignment: over.alignment ?? base.alignment,
            indent: over.indent ?? base.indent,
            marginLeft: over.marginLeft ?? base.marginLeft,
            marginRight: over.marginRight ?? base.marginRight,
            spaceBefore: over.spaceBefore ?? base.spaceBefore,
            spaceAfter: over.spaceAfter ?? base.spaceAfter,
            lineSpacing: over.lineSpacing ?? base.lineSpacing,
            bullet: this.mergeBulletProps(base.bullet, over.bullet),
            defaultRunProperties: this.mergeRunProperties(over.defaultRunProperties, base.defaultRunProperties),
        };
    }
    /**
     * Parses the bullet aspects present on a pPr-like node. Returns undefined
     * when the node specifies nothing bullet-related.
     */
    parseBulletProps(node) {
        const props = {};
        let hasAny = false;
        // Bullet type: buNone / buAutoNum / buChar / buBlip are mutually
        // exclusive. Presence-checked because buNone (and blip refs) are
        // typically empty elements.
        if (this.getPresentChild(node, 'a:buNone')) {
            props.type = 'none';
            hasAny = true;
        }
        else {
            const buAutoNum = this.getPresentChild(node, 'a:buAutoNum');
            const buChar = this.getPresentChild(node, 'a:buChar');
            const buBlip = this.getPresentChild(node, 'a:buBlip');
            if (buAutoNum) {
                props.type = 'auto';
                props.autoNumType = getXmlAttr(buAutoNum, 'type') ?? 'arabicPeriod';
                const startAtAttr = getXmlAttr(buAutoNum, 'startAt');
                if (startAtAttr !== undefined)
                    props.startAt = parseInt(startAtAttr, 10);
                hasAny = true;
            }
            else if (buChar) {
                props.type = 'char';
                props.char = getXmlAttr(buChar, 'char') ?? '\u2022';
                hasAny = true;
            }
            else if (buBlip) {
                props.type = 'picture';
                hasAny = true;
            }
        }
        // Bullet color
        const buClr = getXmlChild(node, 'a:buClr');
        const color = buClr ? this.colorResolver.resolveColorElement(buClr) : undefined;
        if (color) {
            props.color = color;
            hasAny = true;
        }
        // Bullet size
        const buSzPct = getXmlChild(node, 'a:buSzPct');
        if (buSzPct) {
            props.sizePercent = parseInt(getXmlAttr(buSzPct, 'val') ?? '100000', 10) / 1000;
            hasAny = true;
        }
        // Bullet font
        const buFont = getXmlChild(node, 'a:buFont');
        const font = buFont ? getXmlAttr(buFont, 'typeface') : undefined;
        if (font) {
            props.font = font;
            hasAny = true;
        }
        return hasAny ? props : undefined;
    }
    /**
     * Merges bullet aspects, with `over` taking precedence. The bullet type
     * (with its type-specific fields) comes from the most specific source that
     * declares one; color, size, and font inherit independently.
     */
    mergeBulletProps(base, over) {
        if (!base)
            return over;
        if (!over)
            return base;
        const typeSource = over.type !== undefined ? over : base;
        return {
            type: typeSource.type,
            char: typeSource.char,
            autoNumType: typeSource.autoNumType,
            startAt: typeSource.startAt,
            color: over.color ?? base.color,
            sizePercent: over.sizePercent ?? base.sizePercent,
            font: over.font ?? base.font,
        };
    }
    /**
     * Converts merged sparse bullet aspects into a final BulletConfig.
     * Aspects without a resolved type (e.g. an inherited color with no bullet
     * anywhere in the chain) yield no bullet.
     */
    finalizeBullet(props) {
        if (!props?.type)
            return undefined;
        if (props.type === 'none')
            return { type: 'none' };
        return {
            type: props.type,
            char: props.char,
            autoNumType: props.autoNumType,
            ...(props.startAt !== undefined && { startAt: props.startAt }),
            color: props.color,
            sizePercent: props.sizePercent,
            font: props.font,
        };
    }
    /**
     * Parses a text run (a:r).
     */
    parseTextRun(rNode, defaultProps) {
        // Get text content
        const tNode = getXmlChild(rNode, 'a:t');
        if (!tNode)
            return undefined;
        const text = this.extractTextContent(tNode);
        // Parse run properties
        const rPr = getXmlChild(rNode, 'a:rPr');
        const runProps = rPr ? this.parseRunProperties(rPr) : undefined;
        // Merge run properties with defaults (run props override defaults)
        const properties = this.mergeRunProperties(runProps, defaultProps);
        return {
            text,
            properties,
        };
    }
    /**
     * Merges run properties, with primary overriding defaults.
     */
    mergeRunProperties(primary, defaults) {
        if (!primary && !defaults)
            return undefined;
        if (!defaults)
            return primary;
        if (!primary)
            return defaults;
        // Per-script typefaces inherit through the chain like the Latin font.
        // They are spread conditionally so pure-Latin decks keep byte-identical
        // run properties.
        const eastAsianFontFamily = primary.eastAsianFontFamily ?? defaults.eastAsianFontFamily;
        const complexScriptFontFamily = primary.complexScriptFontFamily ?? defaults.complexScriptFontFamily;
        // Merge with primary taking precedence
        return {
            bold: primary.bold ?? defaults.bold,
            italic: primary.italic ?? defaults.italic,
            underline: primary.underline ?? defaults.underline,
            strikethrough: primary.strikethrough ?? defaults.strikethrough,
            fontFamily: primary.fontFamily ?? defaults.fontFamily,
            fontSize: primary.fontSize ?? defaults.fontSize,
            color: primary.color ?? defaults.color,
            baseline: primary.baseline ?? defaults.baseline,
            spacing: primary.spacing ?? defaults.spacing,
            caps: primary.caps ?? defaults.caps,
            kerningMin: primary.kerningMin ?? defaults.kerningMin,
            ...(eastAsianFontFamily !== undefined && { eastAsianFontFamily }),
            ...(complexScriptFontFamily !== undefined && { complexScriptFontFamily }),
        };
    }
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
    parseTextField(fldNode, defaultProps) {
        // Get cached text content (fallback display text)
        const tNode = getXmlChild(fldNode, 'a:t');
        let text = tNode ? this.extractTextContent(tNode) : '';
        const fieldType = getXmlAttr(fldNode, 'type');
        if (fieldType === 'slidenum' && this.slideNumber !== undefined) {
            text = String(this.slideNumber);
        }
        // Parse run properties
        const rPr = getXmlChild(fldNode, 'a:rPr');
        const runProps = rPr ? this.parseRunProperties(rPr) : undefined;
        // Merge run properties with defaults
        const properties = this.mergeRunProperties(runProps, defaultProps);
        return {
            text,
            properties,
        };
    }
    /**
     * Splits a run at script boundaries so each emitted run can be rendered
     * with a single font: CJK segments use the East Asian typeface and
     * complex-script segments (Arabic/Hebrew) the complex-script typeface.
     * Runs without EA/CS characters are returned unchanged (minus the
     * per-script metadata) so pure-Latin content parses exactly as before.
     */
    splitRunByScript(run) {
        const props = run.properties;
        // Fast path: no EA/CS characters
        if (!NON_LATIN_SCRIPT_REGEX.test(run.text)) {
            const stripped = this.stripScriptFonts(props);
            return stripped === props ? [run] : [{ text: run.text, properties: stripped }];
        }
        // Group consecutive characters of the same script class
        const segments = [];
        for (const char of run.text) {
            const script = classifyScriptChar(char);
            const last = segments[segments.length - 1];
            if (last?.script === script) {
                last.text += char;
            }
            else {
                segments.push({ script, text: char });
            }
        }
        return segments.map(({ script, text }) => ({
            text,
            properties: this.applyScriptFont(props, script),
        }));
    }
    /**
     * Returns run properties with fontFamily switched to the segment's script
     * font. EA/CS segments without an explicit typeface anywhere in the style
     * chain fall back to the theme's minor East Asian / complex-script font
     * (resolved by FontResolver, which also handles +mj-ea/+mn-ea/+mj-cs/
     * +mn-cs references).
     */
    applyScriptFont(props, script) {
        if (script === 'ea') {
            return { ...props, fontFamily: props?.eastAsianFontFamily ?? '+mn-ea' };
        }
        if (script === 'cs') {
            return { ...props, fontFamily: props?.complexScriptFontFamily ?? '+mn-cs' };
        }
        return props;
    }
    /**
     * Removes per-script typeface metadata from run properties. Runs without
     * EA/CS characters are emitted without the extra fields so their parsed
     * output is identical to pre-split behavior.
     */
    stripScriptFonts(props) {
        if (!props)
            return undefined;
        if (props.eastAsianFontFamily === undefined && props.complexScriptFontFamily === undefined) {
            return props;
        }
        const rest = { ...props };
        delete rest.eastAsianFontFamily;
        delete rest.complexScriptFontFamily;
        return rest;
    }
    /**
     * Parses run properties (a:rPr).
     */
    parseRunProperties(rPr) {
        // Parse font size (sz attribute is in hundredths of a point)
        const szAttr = getXmlAttr(rPr, 'sz');
        const fontSize = szAttr ? parseInt(szAttr, 10) : undefined;
        // Parse bold
        const bAttr = getXmlAttr(rPr, 'b');
        const bold = bAttr === '1' || bAttr === 'true'
            ? true
            : bAttr === '0' || bAttr === 'false'
                ? false
                : undefined;
        // Parse italic
        const iAttr = getXmlAttr(rPr, 'i');
        const italic = iAttr === '1' || iAttr === 'true'
            ? true
            : iAttr === '0' || iAttr === 'false'
                ? false
                : undefined;
        // Parse underline
        const uAttr = getXmlAttr(rPr, 'u');
        const underline = uAttr !== undefined && uAttr !== 'none' ? true : undefined;
        // Parse strikethrough
        const strikeAttr = getXmlAttr(rPr, 'strike');
        const strikethrough = strikeAttr !== undefined && strikeAttr !== 'noStrike' ? true : undefined;
        // Parse baseline (for super/subscript)
        const baselineAttr = getXmlAttr(rPr, 'baseline');
        const baseline = baselineAttr ? parseInt(baselineAttr, 10) : undefined;
        // Parse character spacing (letter tracking, in hundredths of a point).
        // May be negative for condensed text.
        const spcAttr = getXmlAttr(rPr, 'spc');
        const spacing = spcAttr ? parseInt(spcAttr, 10) : undefined;
        // Parse kerning threshold (minimum font size to kern, hundredths of a
        // point). PowerPoint's defaultTextStyle typically carries kern="1200",
        // which leaves text under 12pt unkerned — measurably wider than shaped
        // output on all-caps small print.
        const kernAttr = getXmlAttr(rPr, 'kern');
        const kerningMin = kernAttr !== undefined ? parseInt(kernAttr, 10) : undefined;
        // Parse capitalization (cap="all"|"small"|"none", ECMA-376 §21.1.2.3.9).
        // 'none' is preserved so an explicit cap="none" overrides an inherited
        // cap="all" through the merge chain.
        const capAttr = getXmlAttr(rPr, 'cap');
        const caps = capAttr === 'all' || capAttr === 'small' || capAttr === 'none' ? capAttr : undefined;
        // Parse font family
        const latin = getXmlChild(rPr, 'a:latin');
        const ea = getXmlChild(rPr, 'a:ea');
        const cs = getXmlChild(rPr, 'a:cs');
        // Only a:latin sets the run's base font. a:ea/a:cs apply exclusively to
        // their script segments (see splitRunByScript); a Latin-script run whose
        // rPr carries only a:ea/a:cs inherits its latin typeface from the style
        // chain (shape lstStyle, placeholder styles, theme), matching PowerPoint.
        const latinTypeface = latin ? getXmlAttr(latin, 'typeface') : undefined;
        const fontFamily = latinTypeface === '' ? undefined : latinTypeface;
        // Per-script typefaces, used when splitting runs at script boundaries.
        // An empty typeface="" means "nothing specified" and is treated as absent.
        const eaTypeface = ea ? getXmlAttr(ea, 'typeface') : undefined;
        const csTypeface = cs ? getXmlAttr(cs, 'typeface') : undefined;
        const eastAsianFontFamily = eaTypeface === '' ? undefined : eaTypeface;
        const complexScriptFontFamily = csTypeface === '' ? undefined : csTypeface;
        // Parse color
        const solidFill = getXmlChild(rPr, 'a:solidFill');
        const color = solidFill ? this.colorResolver.resolveColorElement(solidFill) : undefined;
        // Parse scheme color reference
        const schemeClr = solidFill ? getXmlChild(solidFill, 'a:schemeClr') : undefined;
        const schemeColor = schemeClr ? getXmlAttr(schemeClr, 'val') : undefined;
        if (schemeColor) {
            this.logger.debug('Parsed text color', { schemeColor, resolvedColor: color });
        }
        // Parse color transforms
        const colorTransforms = schemeClr ? this.colorResolver.extractTransforms(schemeClr) : undefined;
        return {
            fontSize,
            fontFamily,
            bold,
            italic,
            underline,
            strikethrough,
            color,
            schemeColor,
            colorTransforms: colorTransforms && Object.keys(colorTransforms).length > 0 ? colorTransforms : undefined,
            baseline,
            spacing,
            caps,
            ...(kerningMin !== undefined && { kerningMin }),
            ...(eastAsianFontFamily !== undefined && { eastAsianFontFamily }),
            ...(complexScriptFontFamily !== undefined && { complexScriptFontFamily }),
        };
    }
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
    parseSpacing(spcNode) {
        if (!spcNode)
            return undefined;
        // Check for points
        const spcPts = getXmlChild(spcNode, 'a:spcPts');
        if (spcPts) {
            const val = getXmlAttr(spcPts, 'val');
            if (!val)
                return undefined;
            // Centipoints, quantized to whole points, converted to EMU
            // (1 point = 12700 EMU)
            return Math.round(parseInt(val, 10) / 100) * 12700;
        }
        // Check for percentage (of line height)
        const spcPct = getXmlChild(spcNode, 'a:spcPct');
        if (spcPct) {
            // Return as negative to indicate percentage mode
            const val = getXmlAttr(spcPct, 'val');
            return val ? -parseInt(val, 10) : undefined;
        }
        return undefined;
    }
    /**
     * Parses line spacing (a:lnSpc).
     *
     * Encoding contract (consumed by TextLayoutEngine.resolveLineHeightPx):
     * - a:spcPct (percentage) is returned as a POSITIVE value in 1000ths of a
     *   percent (100000 = 100%)
     * - a:spcPts (fixed line height) is returned as a NEGATIVE value in
     *   hundredths of a point (-2400 = fixed 24pt)
     */
    parseLineSpacing(lnSpc) {
        if (!lnSpc)
            return undefined;
        // Check for percentage
        const spcPct = getXmlChild(lnSpc, 'a:spcPct');
        if (spcPct) {
            const val = getXmlAttr(spcPct, 'val');
            return val ? parseInt(val, 10) : undefined;
        }
        // Check for points (treat as fixed line height)
        const spcPts = getXmlChild(lnSpc, 'a:spcPts');
        if (spcPts) {
            const val = getXmlAttr(spcPts, 'val');
            // Return as negative to indicate fixed mode
            return val ? -parseInt(val, 10) : undefined;
        }
        return undefined;
    }
    /**
     * Extracts text content from a text node.
     */
    extractTextContent(tNode) {
        if (typeof tNode === 'string') {
            return tNode;
        }
        // Check for #text property (common in parsed XML)
        const textContent = tNode['#text'];
        if (typeof textContent === 'string') {
            return textContent;
        }
        // Check for direct text
        if (typeof tNode === 'object') {
            const keys = Object.keys(tNode);
            for (const key of keys) {
                if (!key.startsWith('@_') && !key.startsWith('#')) {
                    const value = tNode[key];
                    if (typeof value === 'string') {
                        return value;
                    }
                }
            }
        }
        return '';
    }
    /**
     * Parses an integer attribute.
     */
    parseIntAttr(node, attr) {
        const value = getXmlAttr(node, attr);
        return value !== undefined ? parseInt(value, 10) : undefined;
    }
    /**
     * Gets a child element, preserving the presence of empty elements. The XML
     * parser yields self-closed childless elements (e.g. `<a:buNone/>`,
     * `<a:noAutofit/>`) as empty strings, which must not be confused with a
     * missing element.
     */
    getPresentChild(node, name) {
        if (!node || typeof node !== 'object')
            return undefined;
        const child = node[name];
        if (child === undefined)
            return undefined;
        return typeof child === 'object' && child !== null ? child : {};
    }
}
/**
 * Creates a TextParser instance.
 */
export function createTextParser(theme, logger) {
    return new TextParser({ theme, logger });
}
//# sourceMappingURL=TextParser.js.map