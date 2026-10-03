/**
 * Parses chart XML (c:chartSpace) from PPTX files.
 * Extracts chart type, series data, categories, and styling information.
 */
import type { PptxParser } from '../core/PptxParser.js';
import type { Rgba } from '../types/geometry.js';
import type { ResolvedTheme } from '../types/theme.js';
import type { ILogger } from '../utils/Logger.js';
/**
 * Chart types supported for rendering.
 * 'combo' marks a plot area with multiple chart-type nodes sharing axes;
 * 'unknown' marks an unsupported type rendered as a neutral placeholder.
 */
export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'stackedBar' | 'stackedColumn' | 'combo' | 'unknown';
/**
 * The plot kind of an individual series. Matches its owning chart-type node,
 * which matters for combo charts where bar and line series share one plot area.
 */
export type SeriesChartKind = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
/**
 * c:scatterStyle values (ECMA-376 ST_ScatterStyle).
 */
declare const SCATTER_STYLES: readonly ["none", "line", "lineMarker", "marker", "smooth", "smoothMarker"];
/**
 * Scatter plot style: which of connecting lines and point markers are drawn.
 */
export type ScatterStyle = (typeof SCATTER_STYLES)[number];
/**
 * A single data series in a chart.
 */
export interface ChartSeries {
    /** Series name/label */
    name: string;
    /** Data values, indexed by c:pt idx (may contain holes for blank cells) */
    values: number[];
    /**
     * X values for scatter series (c:xVal), aligned with values by index.
     * Absent for category-axis charts and for scatter series without c:xVal
     * (renderers fall back to index-x).
     */
    xValues?: number[];
    /**
     * Plot kind of this series (defaults to the chart type when absent).
     * In combo charts each series keeps the kind of its chart-type node.
     */
    chartKind?: SeriesChartKind;
    /**
     * Bubble sizes (c:bubbleSize) aligned with values by index, set for
     * bubble charts (parsed as scatter). Marker radii scale with the square
     * root of the size so marker area encodes the value.
     */
    bubbleSizes?: number[];
    /**
     * True when the series' chart-type node references a secondary value
     * axis (any c:valAx after the first in the plot area). Such series scale
     * against an independent right-hand axis.
     */
    secondaryAxis?: boolean;
    /** Series color (optional, uses theme if not specified) */
    color?: Rgba;
}
/**
 * Legend configuration.
 */
export interface LegendData {
    /** Legend position */
    position: 'top' | 'bottom' | 'left' | 'right';
    /** Legend entries */
    entries: LegendEntry[];
}
/**
 * A single legend entry.
 */
export interface LegendEntry {
    /** Entry name */
    name: string;
    /** Entry color */
    color?: Rgba;
}
/**
 * Rectangle for positioning.
 */
export interface ChartRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
/**
 * Complete chart data structure.
 */
export interface ChartData {
    /** Chart type */
    type: ChartType;
    /** Data series */
    series: ChartSeries[];
    /** Category labels (x-axis for bar/line, slice labels for pie) */
    categories: string[];
    /** Chart title */
    title?: string;
    /** Legend configuration */
    legend?: LegendData;
    /** Whether to show the legend */
    showLegend: boolean;
    /** Whether to show data labels */
    showDataLabels: boolean;
    /**
     * Doughnut hole size as a percentage of the outer radius (c:holeSize val,
     * default 50). Only set for doughnut charts.
     */
    holeSize?: number;
    /**
     * Scatter line/marker style (c:scatterStyle val). Only set for scatter
     * charts; the line/smooth styles draw connecting polylines between
     * points in data order.
     */
    scatterStyle?: ScatterStyle;
    /**
     * The unsupported chart-type tag (e.g. 'c:radarChart') that produced a
     * placeholder. Only set when type is 'unknown'.
     */
    unknownChartTag?: string;
}
/**
 * Configuration for ChartParser.
 */
export interface ChartParserConfig {
    /** Resolved theme for color resolution */
    theme?: ResolvedTheme;
    /** Logger instance */
    logger?: ILogger;
}
/**
 * Parses chart XML from PPTX.
 */
export declare class ChartParser {
    private readonly logger;
    private readonly theme?;
    private readonly xmlParser;
    constructor(config?: ChartParserConfig);
    /**
     * Parses a chart XML file and returns structured chart data.
     * @param parser The PPTX parser
     * @param chartPath Path to the chart XML file within the PPTX
     * @returns Parsed chart data or undefined if parsing fails
     */
    parseChart(parser: PptxParser, chartPath: string): Promise<ChartData | undefined>;
    /**
     * Parses the c:chartSpace element.
     */
    private parseChartSpace;
    /**
     * Parses the chart title.
     */
    private parseTitle;
    /**
     * Extracts text from a rich text element.
     */
    private extractTextFromRichText;
    /**
     * Extracts text from a string reference.
     */
    private extractTextFromStrRef;
    /**
     * Parses the legend configuration.
     */
    private parseLegend;
    /**
     * Chart-type tags this parser understands, in render-priority order
     * (bars before lines so combo charts default to bar-derived metadata).
     */
    private static readonly SUPPORTED_CHART_TAGS;
    /**
     * Determines the chart type and extracts series/category data.
     * A plot area may contain multiple chart-type nodes sharing axes (a combo
     * chart, e.g. c:barChart + c:lineChart); all of them are parsed and each
     * series is tagged with its plot kind. Chart-type nodes whose c:axId refs
     * pair with a secondary c:valAx get their series flagged secondaryAxis.
     * Unsupported chart types produce an 'unknown' placeholder result instead
     * of failing the whole chart.
     */
    private determineChartType;
    /**
     * Collects the c:axId values of every non-primary c:valAx in the plot
     * area. PowerPoint writes the primary value axis first; any further
     * c:valAx is a secondary axis, and chart-type nodes referencing one via
     * their own c:axId children plot their series against the secondary
     * scale. Returns an empty set when at most one c:valAx exists, keeping
     * single-axis charts untouched.
     */
    private getSecondaryValAxIds;
    /**
     * Returns true when a chart-type node references a secondary value axis
     * through one of its c:axId children. Category-axis ids never match
     * because only c:valAx axIds populate the secondary set.
     */
    private referencesSecondaryAxis;
    /**
     * Marks every series in the list as bound to the secondary value axis.
     */
    private flagSeriesSecondary;
    /**
     * Dispatches a single chart-type node to its parser.
     */
    private parseChartNode;
    /**
     * Parses a combo chart: multiple chart-type nodes sharing one plot area.
     * Series from every node are merged (each tagged with its own chartKind,
     * and flagged secondaryAxis when their node pairs with a secondary
     * c:valAx); categories come from the first node that provides them.
     */
    private parseComboChart;
    /**
     * Handles a plot area with no supported chart-type node. Any element
     * ending in 'Chart' (c:radarChart, c:stockChart, c:surface3DChart,
     * 3D variants, ...) yields an 'unknown' placeholder result so the
     * renderer can draw a neutral stand-in instead of nothing.
     */
    private parseUnknownChartType;
    /**
     * Parses a bar/column chart.
     */
    private parseBarChart;
    /**
     * Parses a line chart.
     */
    private parseLineChart;
    /**
     * Parses a pie chart.
     */
    private parsePieChart;
    /**
     * Parses a doughnut chart: pie data plus a c:holeSize inner cutout
     * (percent of the outer radius, default 50 when unspecified).
     */
    private parseDoughnutChart;
    /**
     * Parses an area chart (rendered as line).
     */
    private parseAreaChart;
    /**
     * Parses a scatter (XY) chart. Series use c:xVal/c:yVal numeric pairs
     * instead of c:cat/c:val; x values are stored per series so scatter can
     * render true x/y positions. A c:xVal holding string data (c:strRef or
     * c:strLit) behaves like a category axis: the strings become categories
     * and the series has no xValues (renderers fall back to index-x).
     */
    private parseScatterChart;
    /**
     * Parses a bubble chart (c:bubbleChart) as scatter data: c:xVal/c:yVal
     * pairs plus a c:bubbleSize value per point. The renderer scales marker
     * radii from the bubble sizes so marker area encodes the value.
     */
    private parseBubbleChart;
    /**
     * Parses one XY (scatter/bubble) series: c:yVal numeric values plus
     * optional c:xVal numeric x coordinates. A c:xVal holding string data
     * yields no xValues (renderers fall back to index-x).
     */
    private parseXYSeries;
    /**
     * Extracts category labels from the first series' c:xVal when it holds
     * string data (c:strRef/c:strLit), which makes an XY chart behave like a
     * category-axis chart. Numeric c:xVal data yields no categories.
     */
    private extractXValCategories;
    /**
     * Extracts string category labels from a data-source node holding
     * c:strRef/c:strCache or c:strLit (used for scatter c:xVal string data).
     */
    private extractStringCategories;
    /**
     * Parses a single data series.
     * @param chartKind The plot kind of the owning chart-type node
     */
    private parseSeries;
    /**
     * Extracts the series name from c:tx (string reference or literal),
     * falling back to a numbered default.
     */
    private parseSeriesName;
    /**
     * Extracts numeric values from a c:val element.
     * Each point is placed at its c:pt idx position: PowerPoint omits c:pt
     * entries for blank cells, so honoring idx keeps series values aligned
     * with categories (holes are handled downstream as gaps).
     */
    private extractValues;
    /**
     * Resolves the target array position of a c:pt from its idx attribute.
     * Falls back to the next dense position when idx is missing, invalid, or
     * implausibly large (a crafted/corrupt idx would otherwise allocate and
     * densify a multi-million-entry array — a denial-of-service vector).
     */
    private getPointIndex;
    /**
     * Upper bound for a c:pt idx this parser honors. Real-world charts stay
     * far below this; anything larger is treated as corrupt and placed at the
     * next dense position instead.
     */
    private static readonly MAX_POINT_INDEX;
    /**
     * Extracts category labels from a series.
     * Supports string (c:strRef/c:strLit) and numeric/date (c:numRef/c:numLit)
     * category axes; numeric categories are formatted plainly with no
     * formatCode interpretation. Each point is placed at its c:pt idx
     * position, and idx gaps (blank cells) become empty labels so categories
     * stay aligned with series values.
     */
    private extractCategories;
    /**
     * Finds the cache node holding category points inside c:cat.
     * Categories may be a string reference (c:strRef/c:strCache), a numeric
     * reference (c:numRef/c:numCache), or inline literals (c:strLit/c:numLit).
     */
    private findCategoryCache;
    /**
     * Extracts the color for a series. The theme-accent fallback uses the
     * series' own c:idx (global across the plot area) so combo charts do not
     * restart the palette per chart-type node.
     */
    private extractSeriesColor;
    /**
     * Gets a theme accent color by index.
     */
    private getThemeAccentColor;
    /**
     * Resolves a scheme color to RGBA.
     */
    private resolveSchemeColor;
    /**
     * Converts a hex color string to RGBA.
     */
    private hexToRgba;
    /**
     * Checks if data labels are enabled for the chart.
     */
    private hasDataLabels;
}
/**
 * Creates a ChartParser instance.
 */
export declare function createChartParser(theme?: ResolvedTheme, logger?: ILogger): ChartParser;
export {};
//# sourceMappingURL=ChartParser.d.ts.map