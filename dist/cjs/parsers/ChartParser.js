"use strict";
/**
 * Parses chart XML (c:chartSpace) from PPTX files.
 * Extracts chart type, series data, categories, and styling information.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.ChartParser = void 0;
exports.createChartParser = createChartParser;
const fast_xml_parser_1 = require("fast-xml-parser");
const PptxParser_js_1 = require("../core/PptxParser.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * c:scatterStyle values (ECMA-376 ST_ScatterStyle).
 */
const SCATTER_STYLES = ['none', 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker'];
/**
 * Parses chart XML from PPTX.
 */
class ChartParser {
    logger;
    theme;
    xmlParser;
    constructor(config = {}) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'ChartParser');
        this.theme = config.theme;
        // Create XMLParser once in constructor for reuse across all chart parsing
        this.xmlParser = new fast_xml_parser_1.XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: '@_',
            removeNSPrefix: false,
            parseAttributeValue: false,
            // Preserve significant whitespace in a:t runs (chart titles are rich
            // text) while trimming structural tags and attributes, matching the
            // main PptxParser behavior.
            trimValues: false,
            tagValueProcessor: PptxParser_js_1.trimTagValueExceptText,
            attributeValueProcessor: PptxParser_js_1.trimAttributeValue,
            isArray: (name) => {
                // Elements that can appear multiple times
                // Note: c:strRef and c:numRef are NOT arrays - they're single elements
                return ['c:ser', 'c:pt'].some((tag) => name.endsWith(tag));
            },
        });
    }
    /**
     * Parses a chart XML file and returns structured chart data.
     * @param parser The PPTX parser
     * @param chartPath Path to the chart XML file within the PPTX
     * @returns Parsed chart data or undefined if parsing fails
     */
    async parseChart(parser, chartPath) {
        try {
            // Read the chart XML
            const chartXml = await parser.readBinary(chartPath);
            const chartString = chartXml.toString('utf-8');
            // Parse the XML using the reusable XMLParser instance
            const parsed = this.xmlParser.parse(chartString);
            // Navigate to chart space
            const chartSpace = (0, PptxParser_js_1.getXmlChild)(parsed, 'c:chartSpace');
            if (!chartSpace) {
                this.logger.warn('No chartSpace found in chart XML');
                return undefined;
            }
            return this.parseChartSpace(chartSpace);
        }
        catch (error) {
            this.logger.error('Failed to parse chart', {
                path: chartPath,
                error: error instanceof Error ? error.message : String(error),
            });
            return undefined;
        }
    }
    /**
     * Parses the c:chartSpace element.
     */
    parseChartSpace(chartSpace) {
        const chart = (0, PptxParser_js_1.getXmlChild)(chartSpace, 'c:chart');
        if (!chart) {
            this.logger.warn('No chart element in chartSpace');
            return undefined;
        }
        // Parse title
        const title = this.parseTitle(chart);
        // Parse legend
        const legendData = this.parseLegend(chart);
        const showLegend = legendData !== undefined;
        // Parse plot area (contains the actual chart data)
        const plotArea = (0, PptxParser_js_1.getXmlChild)(chart, 'c:plotArea');
        if (!plotArea) {
            this.logger.warn('No plotArea in chart');
            return undefined;
        }
        // Determine chart type and extract data
        const chartTypeData = this.determineChartType(plotArea);
        if (!chartTypeData) {
            this.logger.warn('Could not determine chart type');
            return undefined;
        }
        const { type, series, categories, showDataLabels, holeSize, scatterStyle, unknownChartTag } = chartTypeData;
        this.logger.debug('Parsed chart', {
            type,
            seriesCount: series.length,
            categoryCount: categories.length,
            title,
        });
        return {
            type,
            series,
            categories,
            title,
            legend: legendData,
            showLegend,
            showDataLabels,
            holeSize,
            scatterStyle,
            unknownChartTag,
        };
    }
    /**
     * Parses the chart title.
     */
    parseTitle(chart) {
        const title = (0, PptxParser_js_1.getXmlChild)(chart, 'c:title');
        if (!title)
            return undefined;
        // Check for auto-deleted title
        const autoTitleDeleted = (0, PptxParser_js_1.getXmlChild)(chart, 'c:autoTitleDeleted');
        if (autoTitleDeleted) {
            const val = (0, PptxParser_js_1.getXmlAttr)(autoTitleDeleted, 'val');
            if (val === '1' || val === 'true')
                return undefined;
        }
        // Extract text from rich text or string reference
        const tx = (0, PptxParser_js_1.getXmlChild)(title, 'c:tx');
        if (!tx)
            return undefined;
        // Try rich text first
        const rich = (0, PptxParser_js_1.getXmlChild)(tx, 'c:rich');
        if (rich) {
            return this.extractTextFromRichText(rich);
        }
        // Try string reference
        const strRef = (0, PptxParser_js_1.getXmlChild)(tx, 'c:strRef');
        if (strRef) {
            return this.extractTextFromStrRef(strRef);
        }
        return undefined;
    }
    /**
     * Extracts text from a rich text element.
     */
    extractTextFromRichText(rich) {
        const texts = [];
        // Navigate through paragraphs
        const paragraphs = (0, PptxParser_js_1.getXmlChildren)(rich, 'a:p');
        for (const p of paragraphs) {
            const runs = (0, PptxParser_js_1.getXmlChildren)(p, 'a:r');
            for (const r of runs) {
                const t = (0, PptxParser_js_1.getXmlChild)(r, 'a:t');
                if (t) {
                    const text = typeof t === 'string' ? t : t['#text'];
                    if (text)
                        texts.push(text);
                }
            }
        }
        return texts.join('');
    }
    /**
     * Extracts text from a string reference.
     */
    extractTextFromStrRef(strRef) {
        const strCache = (0, PptxParser_js_1.getXmlChild)(strRef, 'c:strCache');
        if (!strCache)
            return undefined;
        const points = (0, PptxParser_js_1.getXmlChildren)(strCache, 'c:pt');
        if (points.length === 0)
            return undefined;
        const v = (0, PptxParser_js_1.getXmlChild)(points[0], 'c:v');
        if (!v)
            return undefined;
        return typeof v === 'string' ? v : v['#text'];
    }
    /**
     * Parses the legend configuration.
     */
    parseLegend(chart) {
        const legend = (0, PptxParser_js_1.getXmlChild)(chart, 'c:legend');
        if (!legend)
            return undefined;
        // Parse position
        const legendPos = (0, PptxParser_js_1.getXmlChild)(legend, 'c:legendPos');
        const posVal = legendPos ? (0, PptxParser_js_1.getXmlAttr)(legendPos, 'val') : 'b';
        let position = 'bottom';
        switch (posVal) {
            case 't':
                position = 'top';
                break;
            case 'b':
                position = 'bottom';
                break;
            case 'l':
                position = 'left';
                break;
            case 'r':
            case 'tr':
                position = 'right';
                break;
        }
        return {
            position,
            entries: [], // Entries are determined from series data
        };
    }
    /**
     * Chart-type tags this parser understands, in render-priority order
     * (bars before lines so combo charts default to bar-derived metadata).
     */
    static SUPPORTED_CHART_TAGS = [
        'c:barChart',
        'c:lineChart',
        'c:pieChart',
        'c:doughnutChart',
        'c:areaChart',
        'c:scatterChart',
        'c:bubbleChart',
    ];
    /**
     * Determines the chart type and extracts series/category data.
     * A plot area may contain multiple chart-type nodes sharing axes (a combo
     * chart, e.g. c:barChart + c:lineChart); all of them are parsed and each
     * series is tagged with its plot kind. Chart-type nodes whose c:axId refs
     * pair with a secondary c:valAx get their series flagged secondaryAxis.
     * Unsupported chart types produce an 'unknown' placeholder result instead
     * of failing the whole chart.
     */
    determineChartType(plotArea) {
        // Collect every supported chart-type node in the plot area
        const chartNodes = [];
        for (const tag of ChartParser.SUPPORTED_CHART_TAGS) {
            for (const node of (0, PptxParser_js_1.getXmlChildren)(plotArea, tag)) {
                chartNodes.push({ tag, node });
            }
        }
        if (chartNodes.length === 0) {
            return this.parseUnknownChartType(plotArea);
        }
        const secondaryValAxIds = this.getSecondaryValAxIds(plotArea);
        if (chartNodes.length === 1) {
            const only = chartNodes[0];
            if (!only)
                return undefined;
            const parsed = this.parseChartNode(only.tag, only.node);
            if (parsed && this.referencesSecondaryAxis(only.node, secondaryValAxIds)) {
                this.flagSeriesSecondary(parsed.series);
            }
            return parsed;
        }
        return this.parseComboChart(chartNodes, secondaryValAxIds);
    }
    /**
     * Collects the c:axId values of every non-primary c:valAx in the plot
     * area. PowerPoint writes the primary value axis first; any further
     * c:valAx is a secondary axis, and chart-type nodes referencing one via
     * their own c:axId children plot their series against the secondary
     * scale. Returns an empty set when at most one c:valAx exists, keeping
     * single-axis charts untouched.
     */
    getSecondaryValAxIds(plotArea) {
        const valAxNodes = (0, PptxParser_js_1.getXmlChildren)(plotArea, 'c:valAx');
        // Scatter/bubble plot areas have NO category axis: their x and y axes
        // are BOTH value axes and both primary. Category charts use catAx (or
        // dateAx) + one primary valAx.
        const hasCategoryAxis = (0, PptxParser_js_1.getXmlChildren)(plotArea, 'c:catAx').length > 0 ||
            (0, PptxParser_js_1.getXmlChildren)(plotArea, 'c:dateAx').length > 0;
        const primaryValAxCount = hasCategoryAxis ? 1 : 2;
        const ids = new Set();
        for (let i = primaryValAxCount; i < valAxNodes.length; i++) {
            const id = (0, PptxParser_js_1.getXmlAttr)((0, PptxParser_js_1.getXmlChild)(valAxNodes[i], 'c:axId'), 'val');
            if (id !== undefined)
                ids.add(id);
        }
        return ids;
    }
    /**
     * Returns true when a chart-type node references a secondary value axis
     * through one of its c:axId children. Category-axis ids never match
     * because only c:valAx axIds populate the secondary set.
     */
    referencesSecondaryAxis(node, secondaryValAxIds) {
        if (secondaryValAxIds.size === 0)
            return false;
        return (0, PptxParser_js_1.getXmlChildren)(node, 'c:axId').some((axId) => {
            const id = (0, PptxParser_js_1.getXmlAttr)(axId, 'val');
            return id !== undefined && secondaryValAxIds.has(id);
        });
    }
    /**
     * Marks every series in the list as bound to the secondary value axis.
     */
    flagSeriesSecondary(series) {
        for (const s of series) {
            s.secondaryAxis = true;
        }
    }
    /**
     * Dispatches a single chart-type node to its parser.
     */
    parseChartNode(tag, node) {
        switch (tag) {
            case 'c:barChart':
                return this.parseBarChart(node);
            case 'c:lineChart':
                return this.parseLineChart(node);
            case 'c:pieChart':
                return this.parsePieChart(node);
            case 'c:doughnutChart':
                return this.parseDoughnutChart(node);
            case 'c:areaChart':
                return this.parseAreaChart(node);
            case 'c:scatterChart':
                return this.parseScatterChart(node);
            case 'c:bubbleChart':
                return this.parseBubbleChart(node);
            default:
                return undefined;
        }
    }
    /**
     * Parses a combo chart: multiple chart-type nodes sharing one plot area.
     * Series from every node are merged (each tagged with its own chartKind,
     * and flagged secondaryAxis when their node pairs with a secondary
     * c:valAx); categories come from the first node that provides them.
     */
    parseComboChart(chartNodes, secondaryValAxIds) {
        const series = [];
        let categories = [];
        let showDataLabels = false;
        let holeSize;
        let scatterStyle;
        for (const { tag, node } of chartNodes) {
            const parsed = this.parseChartNode(tag, node);
            if (!parsed)
                continue;
            if (this.referencesSecondaryAxis(node, secondaryValAxIds)) {
                this.flagSeriesSecondary(parsed.series);
            }
            series.push(...parsed.series);
            if (categories.length === 0 && parsed.categories.length > 0) {
                categories = parsed.categories;
            }
            showDataLabels ||= parsed.showDataLabels;
            holeSize ??= parsed.holeSize;
            scatterStyle ??= parsed.scatterStyle;
        }
        return { type: 'combo', series, categories, showDataLabels, holeSize, scatterStyle };
    }
    /**
     * Handles a plot area with no supported chart-type node. Any element
     * ending in 'Chart' (c:radarChart, c:stockChart, c:surface3DChart,
     * 3D variants, ...) yields an 'unknown' placeholder result so the
     * renderer can draw a neutral stand-in instead of nothing.
     */
    parseUnknownChartType(plotArea) {
        const unknownTag = Object.keys(plotArea).find((key) => key.endsWith('Chart'));
        if (!unknownTag) {
            this.logger.warn('No chart-type node found in plot area');
            return undefined;
        }
        this.logger.warn('Unsupported chart type, using placeholder', { tag: unknownTag });
        return {
            type: 'unknown',
            series: [],
            categories: [],
            showDataLabels: false,
            unknownChartTag: unknownTag,
        };
    }
    /**
     * Parses a bar/column chart.
     */
    parseBarChart(barChart) {
        // Determine orientation and grouping
        const barDir = (0, PptxParser_js_1.getXmlChild)(barChart, 'c:barDir');
        const barDirVal = barDir ? (0, PptxParser_js_1.getXmlAttr)(barDir, 'val') : 'col';
        const isHorizontal = barDirVal === 'bar';
        const grouping = (0, PptxParser_js_1.getXmlChild)(barChart, 'c:grouping');
        const groupingVal = grouping ? (0, PptxParser_js_1.getXmlAttr)(grouping, 'val') : 'clustered';
        const isStacked = groupingVal === 'stacked' || groupingVal === 'percentStacked';
        let type;
        if (isHorizontal) {
            type = isStacked ? 'stackedBar' : 'bar';
        }
        else {
            type = isStacked ? 'stackedColumn' : 'column';
        }
        // Parse series
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(barChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseSeries(ser, index, 'bar'));
        // Extract categories from first series
        const firstSeries = seriesNodes[0];
        const categories = firstSeries ? this.extractCategories(firstSeries) : [];
        // Check for data labels
        const showDataLabels = this.hasDataLabels(barChart);
        return { type, series, categories, showDataLabels };
    }
    /**
     * Parses a line chart.
     */
    parseLineChart(lineChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(lineChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseSeries(ser, index, 'line'));
        const firstSeries = seriesNodes[0];
        const categories = firstSeries ? this.extractCategories(firstSeries) : [];
        const showDataLabels = this.hasDataLabels(lineChart);
        return { type: 'line', series, categories, showDataLabels };
    }
    /**
     * Parses a pie chart.
     */
    parsePieChart(pieChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(pieChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseSeries(ser, index, 'pie'));
        const firstSeries = seriesNodes[0];
        const categories = firstSeries ? this.extractCategories(firstSeries) : [];
        const showDataLabels = this.hasDataLabels(pieChart);
        return { type: 'pie', series, categories, showDataLabels };
    }
    /**
     * Parses a doughnut chart: pie data plus a c:holeSize inner cutout
     * (percent of the outer radius, default 50 when unspecified).
     */
    parseDoughnutChart(doughnutChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(doughnutChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseSeries(ser, index, 'doughnut'));
        const firstSeries = seriesNodes[0];
        const categories = firstSeries ? this.extractCategories(firstSeries) : [];
        const showDataLabels = this.hasDataLabels(doughnutChart);
        // c:holeSize val is a percentage of the outer radius
        const holeSizeNode = (0, PptxParser_js_1.getXmlChild)(doughnutChart, 'c:holeSize');
        const holeSizeVal = parseInt((0, PptxParser_js_1.getXmlAttr)(holeSizeNode, 'val') ?? '', 10);
        const holeSize = Number.isInteger(holeSizeVal) && holeSizeVal > 0 && holeSizeVal < 100 ? holeSizeVal : 50;
        return { type: 'doughnut', series, categories, showDataLabels, holeSize };
    }
    /**
     * Parses an area chart (rendered as line).
     */
    parseAreaChart(areaChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(areaChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseSeries(ser, index, 'area'));
        const firstSeries = seriesNodes[0];
        const categories = firstSeries ? this.extractCategories(firstSeries) : [];
        const showDataLabels = this.hasDataLabels(areaChart);
        return { type: 'area', series, categories, showDataLabels };
    }
    /**
     * Parses a scatter (XY) chart. Series use c:xVal/c:yVal numeric pairs
     * instead of c:cat/c:val; x values are stored per series so scatter can
     * render true x/y positions. A c:xVal holding string data (c:strRef or
     * c:strLit) behaves like a category axis: the strings become categories
     * and the series has no xValues (renderers fall back to index-x).
     */
    parseScatterChart(scatterChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(scatterChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => this.parseXYSeries(ser, index));
        // String x data (rare, but legal) acts as a category axis
        const categories = this.extractXValCategories(seriesNodes);
        const showDataLabels = this.hasDataLabels(scatterChart);
        // c:scatterStyle decides whether points are joined by lines
        const styleVal = (0, PptxParser_js_1.getXmlAttr)((0, PptxParser_js_1.getXmlChild)(scatterChart, 'c:scatterStyle'), 'val');
        const scatterStyle = SCATTER_STYLES.includes(styleVal ?? '')
            ? styleVal
            : undefined;
        return { type: 'scatter', series, categories, showDataLabels, scatterStyle };
    }
    /**
     * Parses a bubble chart (c:bubbleChart) as scatter data: c:xVal/c:yVal
     * pairs plus a c:bubbleSize value per point. The renderer scales marker
     * radii from the bubble sizes so marker area encodes the value.
     */
    parseBubbleChart(bubbleChart) {
        const seriesNodes = (0, PptxParser_js_1.getXmlChildren)(bubbleChart, 'c:ser');
        const series = seriesNodes.map((ser, index) => {
            const base = this.parseXYSeries(ser, index);
            const bubbleSizes = this.extractValues((0, PptxParser_js_1.getXmlChild)(ser, 'c:bubbleSize'));
            return bubbleSizes.length > 0 ? { ...base, bubbleSizes } : base;
        });
        // String x data acts as a category axis, same as scatter
        const categories = this.extractXValCategories(seriesNodes);
        const showDataLabels = this.hasDataLabels(bubbleChart);
        return { type: 'scatter', series, categories, showDataLabels };
    }
    /**
     * Parses one XY (scatter/bubble) series: c:yVal numeric values plus
     * optional c:xVal numeric x coordinates. A c:xVal holding string data
     * yields no xValues (renderers fall back to index-x).
     */
    parseXYSeries(ser, index) {
        const name = this.parseSeriesName(ser, index);
        const values = this.extractValues((0, PptxParser_js_1.getXmlChild)(ser, 'c:yVal'));
        const xVal = (0, PptxParser_js_1.getXmlChild)(ser, 'c:xVal');
        const xValues = xVal ? this.extractValues(xVal) : [];
        const color = this.extractSeriesColor(ser, index);
        return {
            name,
            values,
            xValues: xValues.length > 0 ? xValues : undefined,
            chartKind: 'scatter',
            color,
        };
    }
    /**
     * Extracts category labels from the first series' c:xVal when it holds
     * string data (c:strRef/c:strLit), which makes an XY chart behave like a
     * category-axis chart. Numeric c:xVal data yields no categories.
     */
    extractXValCategories(seriesNodes) {
        const firstXVal = seriesNodes[0] ? (0, PptxParser_js_1.getXmlChild)(seriesNodes[0], 'c:xVal') : undefined;
        return firstXVal ? this.extractStringCategories(firstXVal) : [];
    }
    /**
     * Extracts string category labels from a data-source node holding
     * c:strRef/c:strCache or c:strLit (used for scatter c:xVal string data).
     */
    extractStringCategories(source) {
        const strRef = (0, PptxParser_js_1.getXmlChild)(source, 'c:strRef');
        const cache = strRef ? (0, PptxParser_js_1.getXmlChild)(strRef, 'c:strCache') : (0, PptxParser_js_1.getXmlChild)(source, 'c:strLit');
        if (!cache)
            return [];
        const points = (0, PptxParser_js_1.getXmlChildren)(cache, 'c:pt');
        const categories = [];
        for (const pt of points) {
            const v = (0, PptxParser_js_1.getXmlChild)(pt, 'c:v');
            if (v !== undefined && v !== null) {
                // Numeric-looking values are number-coerced by the XML parser
                let text;
                if (typeof v === 'string' || typeof v === 'number') {
                    text = String(v);
                }
                else {
                    text = String(v['#text'] ?? '');
                }
                categories[this.getPointIndex(pt, categories.length)] = text;
            }
        }
        for (let i = 0; i < categories.length; i++) {
            categories[i] ??= '';
        }
        return categories;
    }
    /**
     * Parses a single data series.
     * @param chartKind The plot kind of the owning chart-type node
     */
    parseSeries(ser, index, chartKind) {
        const name = this.parseSeriesName(ser, index);
        // Parse values
        const val = (0, PptxParser_js_1.getXmlChild)(ser, 'c:val');
        const values = this.extractValues(val);
        // Parse color (optional)
        const color = this.extractSeriesColor(ser, index);
        return { name, values, chartKind, color };
    }
    /**
     * Extracts the series name from c:tx (string reference or literal),
     * falling back to a numbered default.
     */
    parseSeriesName(ser, index) {
        let name = `Series ${index + 1}`;
        const tx = (0, PptxParser_js_1.getXmlChild)(ser, 'c:tx');
        if (tx) {
            const strRef = (0, PptxParser_js_1.getXmlChild)(tx, 'c:strRef');
            if (strRef) {
                const extracted = this.extractTextFromStrRef(strRef);
                if (extracted)
                    name = extracted;
            }
            else {
                const v = (0, PptxParser_js_1.getXmlChild)(tx, 'c:v');
                if (v) {
                    const text = typeof v === 'string' ? v : v['#text'];
                    if (text)
                        name = text;
                }
            }
        }
        return name;
    }
    /**
     * Extracts numeric values from a c:val element.
     * Each point is placed at its c:pt idx position: PowerPoint omits c:pt
     * entries for blank cells, so honoring idx keeps series values aligned
     * with categories (holes are handled downstream as gaps).
     */
    extractValues(val) {
        if (!val)
            return [];
        // Values come from a c:numRef cache or an inline c:numLit literal
        const numRef = (0, PptxParser_js_1.getXmlChild)(val, 'c:numRef');
        const cache = numRef ? (0, PptxParser_js_1.getXmlChild)(numRef, 'c:numCache') : (0, PptxParser_js_1.getXmlChild)(val, 'c:numLit');
        if (!cache)
            return [];
        const points = (0, PptxParser_js_1.getXmlChildren)(cache, 'c:pt');
        const values = [];
        for (const pt of points) {
            const v = (0, PptxParser_js_1.getXmlChild)(pt, 'c:v');
            if (v !== undefined && v !== null) {
                // v can be a number, string, or object with #text
                let num;
                if (typeof v === 'number') {
                    num = v;
                }
                else if (typeof v === 'string') {
                    num = parseFloat(v);
                }
                else {
                    const text = v['#text'];
                    num = typeof text === 'number' ? text : parseFloat(text || '0');
                }
                values[this.getPointIndex(pt, values.length)] = isNaN(num) ? 0 : num;
            }
        }
        return values;
    }
    /**
     * Resolves the target array position of a c:pt from its idx attribute.
     * Falls back to the next dense position when idx is missing, invalid, or
     * implausibly large (a crafted/corrupt idx would otherwise allocate and
     * densify a multi-million-entry array — a denial-of-service vector).
     */
    getPointIndex(pt, fallback) {
        const idx = parseInt((0, PptxParser_js_1.getXmlAttr)(pt, 'idx') ?? '', 10);
        return Number.isInteger(idx) && idx >= 0 && idx <= ChartParser.MAX_POINT_INDEX ? idx : fallback;
    }
    /**
     * Upper bound for a c:pt idx this parser honors. Real-world charts stay
     * far below this; anything larger is treated as corrupt and placed at the
     * next dense position instead.
     */
    static MAX_POINT_INDEX = 100_000;
    /**
     * Extracts category labels from a series.
     * Supports string (c:strRef/c:strLit) and numeric/date (c:numRef/c:numLit)
     * category axes; numeric categories are formatted plainly with no
     * formatCode interpretation. Each point is placed at its c:pt idx
     * position, and idx gaps (blank cells) become empty labels so categories
     * stay aligned with series values.
     */
    extractCategories(ser) {
        const cat = (0, PptxParser_js_1.getXmlChild)(ser, 'c:cat');
        if (!cat)
            return [];
        const cache = this.findCategoryCache(cat);
        if (!cache)
            return [];
        const points = (0, PptxParser_js_1.getXmlChildren)(cache, 'c:pt');
        const categories = [];
        for (const pt of points) {
            const v = (0, PptxParser_js_1.getXmlChild)(pt, 'c:v');
            if (v !== undefined && v !== null) {
                // v can be a number, string, or object with #text
                let text;
                if (typeof v === 'string') {
                    text = v;
                }
                else if (typeof v === 'number') {
                    text = String(v);
                }
                else {
                    const inner = v['#text'];
                    text = inner !== undefined && inner !== null ? String(inner) : '';
                }
                categories[this.getPointIndex(pt, categories.length)] = text;
            }
        }
        // Fill idx gaps (blank category cells) with empty labels
        for (let i = 0; i < categories.length; i++) {
            categories[i] ??= '';
        }
        return categories;
    }
    /**
     * Finds the cache node holding category points inside c:cat.
     * Categories may be a string reference (c:strRef/c:strCache), a numeric
     * reference (c:numRef/c:numCache), or inline literals (c:strLit/c:numLit).
     */
    findCategoryCache(cat) {
        const strRef = (0, PptxParser_js_1.getXmlChild)(cat, 'c:strRef');
        if (strRef)
            return (0, PptxParser_js_1.getXmlChild)(strRef, 'c:strCache');
        const numRef = (0, PptxParser_js_1.getXmlChild)(cat, 'c:numRef');
        if (numRef)
            return (0, PptxParser_js_1.getXmlChild)(numRef, 'c:numCache');
        return (0, PptxParser_js_1.getXmlChild)(cat, 'c:strLit') ?? (0, PptxParser_js_1.getXmlChild)(cat, 'c:numLit');
    }
    /**
     * Extracts the color for a series. The theme-accent fallback uses the
     * series' own c:idx (global across the plot area) so combo charts do not
     * restart the palette per chart-type node.
     */
    extractSeriesColor(ser, index) {
        const idxAttr = (0, PptxParser_js_1.getXmlAttr)((0, PptxParser_js_1.getXmlChild)(ser, 'c:idx'), 'val');
        const globalIdx = idxAttr !== undefined ? parseInt(idxAttr, 10) : NaN;
        const paletteIdx = Number.isInteger(globalIdx) && globalIdx >= 0 ? globalIdx : index;
        const spPr = (0, PptxParser_js_1.getXmlChild)(ser, 'c:spPr');
        if (!spPr)
            return this.getThemeAccentColor(paletteIdx);
        // Check for solid fill
        const solidFill = (0, PptxParser_js_1.getXmlChild)(spPr, 'a:solidFill');
        if (solidFill) {
            // Try RGB color
            const srgbClr = (0, PptxParser_js_1.getXmlChild)(solidFill, 'a:srgbClr');
            if (srgbClr) {
                const val = (0, PptxParser_js_1.getXmlAttr)(srgbClr, 'val');
                if (val) {
                    return this.hexToRgba(val);
                }
            }
            // Try scheme color
            const schemeClr = (0, PptxParser_js_1.getXmlChild)(solidFill, 'a:schemeClr');
            if (schemeClr && this.theme) {
                const val = (0, PptxParser_js_1.getXmlAttr)(schemeClr, 'val');
                if (val) {
                    return this.resolveSchemeColor(val);
                }
            }
        }
        return this.getThemeAccentColor(index);
    }
    /**
     * Gets a theme accent color by index.
     */
    getThemeAccentColor(index) {
        if (!this.theme)
            return undefined;
        const colors = this.theme.colors;
        const accents = [
            colors.accent1,
            colors.accent2,
            colors.accent3,
            colors.accent4,
            colors.accent5,
            colors.accent6,
        ];
        return accents[index % accents.length];
    }
    /**
     * Resolves a scheme color to RGBA.
     */
    resolveSchemeColor(schemeValue) {
        if (!this.theme)
            return undefined;
        const colors = this.theme.colors;
        const colorMap = {
            dk1: colors.dark1,
            lt1: colors.light1,
            dk2: colors.dark2,
            lt2: colors.light2,
            tx1: colors.dark1,
            tx2: colors.dark2,
            bg1: colors.light1,
            bg2: colors.light2,
            accent1: colors.accent1,
            accent2: colors.accent2,
            accent3: colors.accent3,
            accent4: colors.accent4,
            accent5: colors.accent5,
            accent6: colors.accent6,
            hlink: colors.hyperlink,
            folHlink: colors.followedHyperlink,
        };
        return colorMap[schemeValue];
    }
    /**
     * Converts a hex color string to RGBA.
     */
    hexToRgba(hex) {
        const cleanHex = hex.replace('#', '');
        const r = parseInt(cleanHex.substring(0, 2), 16);
        const g = parseInt(cleanHex.substring(2, 4), 16);
        const b = parseInt(cleanHex.substring(4, 6), 16);
        return { r, g, b, a: 255 };
    }
    /**
     * Checks if data labels are enabled for the chart.
     */
    hasDataLabels(chartNode) {
        const dLbls = (0, PptxParser_js_1.getXmlChild)(chartNode, 'c:dLbls');
        if (!dLbls)
            return false;
        const showVal = (0, PptxParser_js_1.getXmlChild)(dLbls, 'c:showVal');
        if (showVal) {
            const val = (0, PptxParser_js_1.getXmlAttr)(showVal, 'val');
            return val === '1' || val === 'true';
        }
        return false;
    }
}
exports.ChartParser = ChartParser;
/**
 * Creates a ChartParser instance.
 */
function createChartParser(theme, logger) {
    return new ChartParser({ theme, logger });
}
