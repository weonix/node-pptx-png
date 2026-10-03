/**
 * Renders charts to canvas using basic canvas primitives.
 * Supports bar, column, line, pie, doughnut (single-series and concentric
 * multi-series), scatter (with optional connecting lines, string-x
 * categories, and bubble sizing), and combo charts with independent
 * primary/secondary value axes, plus a neutral placeholder for unsupported
 * chart types.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rect } from '../types/geometry.js';
import type { ChartData } from '../parsers/ChartParser.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * A run of consecutive defined data points within a series.
 */
export interface SeriesSegment {
    /** Category index of the point */
    index: number;
    /** The defined value at that index */
    value: number;
}
/**
 * Splits a (possibly sparse) values array into contiguous runs of defined
 * points. Holes — sparse slots, explicit undefined, or non-finite values —
 * terminate the current run, so line charts break at gaps instead of
 * bridging them and single isolated points render as markers only.
 * @param values The series values (may contain holes)
 * @param limit Maximum number of indices to consider (e.g. category count)
 * @returns Runs of consecutive defined points, in index order
 */
export declare function segmentSeriesValues(values: number[], limit: number): SeriesSegment[][];
/**
 * Configuration for ChartRenderer.
 */
export interface ChartRendererConfig {
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector for chart fidelity events */
    warnings?: WarningCollector;
}
/**
 * Renders charts to canvas.
 */
export declare class ChartRenderer {
    private readonly logger;
    private readonly warnings;
    constructor(config?: ChartRendererConfig);
    /**
     * Renders a chart to the canvas.
     * @param ctx Canvas 2D context
     * @param chartData Parsed chart data
     * @param bounds The bounds to render within
     */
    renderChart(ctx: CanvasRenderingContext2D, chartData: ChartData, bounds: Rect): void;
    /**
     * Calculates the layout areas for the chart.
     */
    private calculateLayout;
    /**
     * Renders the chart title.
     */
    private renderTitle;
    /**
     * Renders a bar or column chart.
     */
    private renderBarChart;
    /**
     * Draws the X and Y axis lines along the plot area edges, plus the
     * secondary value-axis line on the right edge when requested.
     */
    private drawAxisLines;
    /**
     * Renders a line chart. Holes in the data (blank cells) break the line:
     * no point and no connecting segment is drawn across a gap.
     */
    private renderLineChart;
    /**
     * Draws one line series as gap-separated polylines with point markers.
     * Each contiguous run of defined values becomes its own polyline; isolated
     * points render as a marker with no connecting line.
     */
    private drawLineSeries;
    /**
     * Renders a pie or doughnut chart. Holes in the data (blank cells) are
     * skipped entirely — no slice is drawn and the total excludes them.
     * Doughnut charts leave an inner cutout of holeSize percent (default 50)
     * of the outer radius empty and render multiple series as concentric
     * rings (first series innermost, matching PowerPoint), each ring
     * (outerR - holeR) / seriesCount thick. Multi-series pies render only
     * the first series.
     */
    private renderPieChart;
    /**
     * Draws one series as pie slices between the given radii (full wedges
     * when innerRadius is 0, annular segments otherwise). Slice colors vary
     * by point index; holes in the data draw no slice and are excluded from
     * the series total.
     */
    private drawPieSeries;
    /**
     * Renders a scatter (XY) chart as marker circles, covering bubble charts
     * (parsed as scatter with per-point bubble sizes) as well.
     * - Series without xValues fall back to index-x (0, 1, 2, ...). When the
     *   chart carries string-x categories, points sit at category band
     *   centers and the x-axis shows the category labels instead of numeric
     *   ticks.
     * - Holes in either coordinate skip the point; a line/lineMarker/smooth/
     *   smoothMarker scatterStyle connects surviving points in data order
     *   with straight polylines that break at holes (smooth curves are
     *   approximated straight). The pure line/smooth styles hide markers.
     * - Bubble radii use square-root scaling so marker area encodes the
     *   bubble value, capped at 8% of the plot width; blank or non-positive
     *   bubble sizes hide the point, matching PowerPoint's default.
     */
    private renderScatterChart;
    /**
     * Maximum bubble-chart marker radius as a fraction of the plot width.
     */
    private static readonly MAX_BUBBLE_RADIUS_FRACTION;
    /**
     * Default scatter marker radius in pixels (non-bubble series).
     */
    private static readonly SCATTER_MARKER_RADIUS;
    /**
     * Builds the data-space point for one scatter/bubble entry, or undefined
     * when it is a hole: y or x missing/non-finite, or (for bubble series) a
     * blank or non-positive bubble size. In category mode x is always the
     * data index.
     */
    private buildScatterPoint;
    /**
     * Strokes straight polylines through contiguous runs of defined points
     * in data order. Holes break the line; runs of a single point draw
     * nothing (the marker pass covers them).
     */
    private drawScatterPolylines;
    /**
     * Renders a combo chart: bar-kind series as clustered columns first, then
     * line-kind series on top. Series share the primary value-axis scale
     * unless flagged secondaryAxis, in which case they scale independently
     * against the right-hand secondary axis.
     */
    private renderComboChart;
    /**
     * Renders a neutral placeholder for unsupported chart types: a light gray
     * rounded rectangle with the chart title (or 'Chart') centered inside.
     */
    private renderUnknownPlaceholder;
    /**
     * Chart types that render a secondary right-hand value axis. Horizontal
     * bar charts (value axis is horizontal), stacked charts (series share
     * one stack scale), and scatter charts keep single-axis rendering even
     * when series carry a secondaryAxis flag.
     */
    private static readonly SECONDARY_AXIS_TYPES;
    /**
     * Whether the chart renders a secondary right-hand value axis: at least
     * one series is bound to a secondary axis and the chart type supports
     * drawing one.
     */
    private hasSecondaryValueAxis;
    /**
     * The series scaled by the primary value axis: every series when no
     * secondary axis is rendered, otherwise only those not flagged secondary.
     */
    private primaryAxisSeries;
    /**
     * Computes the secondary value-axis maximum (with the standard positive
     * floor and 10% headroom, matching the primary axis) over the
     * secondary-axis series, or undefined when the chart renders no
     * secondary axis.
     */
    private computeSecondaryMax;
    /**
     * Collects every defined, finite value across the given series. Holes
     * (sparse slots, explicit undefined, NaN) are excluded so scale
     * calculations ignore them instead of treating them as 0.
     */
    private collectDefinedValues;
    /**
     * Returns the series value at the given category index, or undefined when
     * the slot is a hole (blank cell) or out of range.
     */
    private getDefinedValue;
    /**
     * Resolves the plot kind of a series, defaulting to a kind derived from
     * the overall chart type when the series carries none.
     */
    private getSeriesKind;
    /**
     * Renders the legend.
     */
    private renderLegend;
    /**
     * Renders horizontal category labels (for column charts).
     */
    private renderCategoryLabelsHorizontal;
    /**
     * Renders vertical category labels (for bar charts).
     */
    private renderCategoryLabelsVertical;
    /**
     * Renders vertical value labels (Y-axis for column/line charts). The
     * 'left' side right-aligns labels against the plot area (primary axis);
     * 'right' left-aligns them (secondary axis on the right of the plot).
     */
    private renderValueLabelsVertical;
    /**
     * Renders horizontal value labels (X-axis for bar charts).
     */
    private renderValueLabelsHorizontal;
    /**
     * Gets the color for a series.
     */
    private getSeriesColor;
    /**
     * Converts RGBA to CSS color string.
     */
    private rgbaToString;
    /**
     * Truncates a label to max length.
     */
    private truncateLabel;
    /**
     * Formats a value-axis tick label. Precision derives from the tick step
     * so fractional steps keep enough decimals to stay distinct (step 0.25
     * -> "0.25") while integer steps render clean whole numbers, and
     * floating-point noise (e.g. 0.30000000000000004) never leaks into a
     * label. Large magnitudes use K/M suffixes.
     */
    private formatValue;
    /**
     * Formats a number with just enough decimals to resolve the given step,
     * trimming trailing zeros (step 0.25 -> 2 decimals, step 2.2 -> 1,
     * step 22 -> 0).
     */
    private formatWithStepPrecision;
}
/**
 * Creates a ChartRenderer instance.
 */
export declare function createChartRenderer(logger?: ILogger): ChartRenderer;
