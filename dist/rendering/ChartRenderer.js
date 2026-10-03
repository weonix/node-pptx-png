/**
 * Renders charts to canvas using basic canvas primitives.
 * Supports bar, column, line, pie, doughnut (single-series and concentric
 * multi-series), scatter (with optional connecting lines, string-x
 * categories, and bubble sizing), and combo charts with independent
 * primary/secondary value axes, plus a neutral placeholder for unsupported
 * chart types.
 */
import { createLogger } from '../utils/Logger.js';
/**
 * Splits a (possibly sparse) values array into contiguous runs of defined
 * points. Holes — sparse slots, explicit undefined, or non-finite values —
 * terminate the current run, so line charts break at gaps instead of
 * bridging them and single isolated points render as markers only.
 * @param values The series values (may contain holes)
 * @param limit Maximum number of indices to consider (e.g. category count)
 * @returns Runs of consecutive defined points, in index order
 */
export function segmentSeriesValues(values, limit) {
    const segments = [];
    let current = [];
    const count = Math.min(values.length, limit);
    for (let i = 0; i < count; i++) {
        const value = values[i];
        if (value !== undefined && Number.isFinite(value)) {
            current.push({ index: i, value });
        }
        else if (current.length > 0) {
            segments.push(current);
            current = [];
        }
    }
    if (current.length > 0)
        segments.push(current);
    return segments;
}
/**
 * Default colors for chart series (similar to Office theme accents).
 */
const DEFAULT_SERIES_COLORS = [
    { r: 68, g: 114, b: 196, a: 255 }, // Blue
    { r: 237, g: 125, b: 49, a: 255 }, // Orange
    { r: 165, g: 165, b: 165, a: 255 }, // Gray
    { r: 255, g: 192, b: 0, a: 255 }, // Yellow
    { r: 91, g: 155, b: 213, a: 255 }, // Light Blue
    { r: 112, g: 173, b: 71, a: 255 }, // Green
    { r: 158, g: 72, b: 14, a: 255 }, // Dark Orange
    { r: 153, g: 115, b: 0, a: 255 }, // Dark Yellow
];
/**
 * Layout constants for chart rendering.
 */
const LAYOUT = {
    titleHeight: 30,
    legendHeight: 25,
    axisLabelWidth: 50,
    axisLabelHeight: 25,
    padding: 10,
    barGap: 0.2, // Gap between bars as fraction of bar width
    fontSize: {
        title: 14,
        axis: 9,
        legend: 10,
    },
};
/**
 * Renders charts to canvas.
 */
export class ChartRenderer {
    logger;
    warnings;
    constructor(config = {}) {
        this.logger = config.logger ?? createLogger('warn', 'ChartRenderer');
        this.warnings = config.warnings;
    }
    /**
     * Renders a chart to the canvas.
     * @param ctx Canvas 2D context
     * @param chartData Parsed chart data
     * @param bounds The bounds to render within
     */
    renderChart(ctx, chartData, bounds) {
        if (chartData.type === 'unknown') {
            this.logger.warn('Unsupported chart type, rendering placeholder', {
                tag: chartData.unknownChartTag,
            });
            this.warnings?.push({
                code: 'unsupported-chart-type',
                message: `Unsupported chart type "${chartData.unknownChartTag ?? 'unknown'}"; placeholder rendered`,
                detail: { tag: chartData.unknownChartTag },
            });
            this.renderUnknownPlaceholder(ctx, chartData, bounds);
            return;
        }
        if (chartData.series.length === 0) {
            this.logger.debug('No series data, skipping chart render');
            return;
        }
        this.logger.debug('Rendering chart', {
            type: chartData.type,
            seriesCount: chartData.series.length,
            categoryCount: chartData.categories.length,
        });
        // Calculate layout areas
        const layout = this.calculateLayout(chartData, bounds);
        // Fill background with white
        ctx.save();
        ctx.fillStyle = 'white';
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        // Render title if present
        if (chartData.title) {
            this.renderTitle(ctx, chartData.title, layout.titleArea);
        }
        // Render the chart based on type
        switch (chartData.type) {
            case 'bar':
            case 'stackedBar':
                this.renderBarChart(ctx, chartData, layout, true);
                break;
            case 'column':
            case 'stackedColumn':
                this.renderBarChart(ctx, chartData, layout, false);
                break;
            case 'line':
            case 'area':
                this.renderLineChart(ctx, chartData, layout);
                break;
            case 'pie':
            case 'doughnut':
                this.renderPieChart(ctx, chartData, layout);
                break;
            case 'scatter':
                this.renderScatterChart(ctx, chartData, layout);
                break;
            case 'combo':
                this.renderComboChart(ctx, chartData, layout);
                break;
            default:
                // Default to column chart
                this.renderBarChart(ctx, chartData, layout, false);
        }
        // Render legend if enabled
        if (chartData.showLegend && chartData.series.length > 0) {
            this.renderLegend(ctx, chartData, layout.legendArea);
        }
        ctx.restore();
    }
    /**
     * Calculates the layout areas for the chart.
     */
    calculateLayout(chartData, bounds) {
        let top = bounds.y + LAYOUT.padding;
        let bottom = bounds.y + bounds.height - LAYOUT.padding;
        let left = bounds.x + LAYOUT.padding;
        let right = bounds.x + bounds.width - LAYOUT.padding;
        const layout = {
            titleArea: { x: left, y: top, width: 0, height: 0 },
            legendArea: { x: left, y: top, width: 0, height: 0 },
            plotArea: { x: left, y: top, width: 0, height: 0 },
            xAxisArea: { x: left, y: top, width: 0, height: 0 },
            yAxisArea: { x: left, y: top, width: 0, height: 0 },
            y2AxisArea: { x: left, y: top, width: 0, height: 0 },
        };
        // Reserve space for title
        if (chartData.title) {
            layout.titleArea = {
                x: left,
                y: top,
                width: right - left,
                height: LAYOUT.titleHeight,
            };
            top += LAYOUT.titleHeight;
        }
        // Reserve space for legend
        if (chartData.showLegend && chartData.series.length > 0) {
            const legendPos = chartData.legend?.position ?? 'bottom';
            if (legendPos === 'bottom') {
                layout.legendArea = {
                    x: left,
                    y: bottom - LAYOUT.legendHeight,
                    width: right - left,
                    height: LAYOUT.legendHeight,
                };
                bottom -= LAYOUT.legendHeight;
            }
            else if (legendPos === 'top') {
                layout.legendArea = {
                    x: left,
                    y: top,
                    width: right - left,
                    height: LAYOUT.legendHeight,
                };
                top += LAYOUT.legendHeight;
            }
            else if (legendPos === 'right') {
                const legendWidth = 100;
                layout.legendArea = {
                    x: right - legendWidth,
                    y: top,
                    width: legendWidth,
                    height: bottom - top,
                };
                right -= legendWidth;
            }
            else if (legendPos === 'left') {
                const legendWidth = 100;
                layout.legendArea = {
                    x: left,
                    y: top,
                    width: legendWidth,
                    height: bottom - top,
                };
                left += legendWidth;
            }
        }
        // Reserve space for axes (except pie/doughnut charts)
        if (chartData.type !== 'pie' && chartData.type !== 'doughnut') {
            // Y-axis labels on left
            layout.yAxisArea = {
                x: left,
                y: top,
                width: LAYOUT.axisLabelWidth,
                height: bottom - top - LAYOUT.axisLabelHeight,
            };
            left += LAYOUT.axisLabelWidth;
            // Secondary value-axis labels on the right, only when one is rendered
            if (this.hasSecondaryValueAxis(chartData)) {
                layout.y2AxisArea = {
                    x: right - LAYOUT.axisLabelWidth,
                    y: top,
                    width: LAYOUT.axisLabelWidth,
                    height: bottom - top - LAYOUT.axisLabelHeight,
                };
                right -= LAYOUT.axisLabelWidth;
            }
            // X-axis labels on bottom
            layout.xAxisArea = {
                x: left,
                y: bottom - LAYOUT.axisLabelHeight,
                width: right - left,
                height: LAYOUT.axisLabelHeight,
            };
            bottom -= LAYOUT.axisLabelHeight;
        }
        // Remaining area is the plot area
        layout.plotArea = {
            x: left,
            y: top,
            width: right - left,
            height: bottom - top,
        };
        return layout;
    }
    /**
     * Renders the chart title.
     */
    renderTitle(ctx, title, area) {
        ctx.save();
        ctx.font = `${LAYOUT.fontSize.title}px Calibri, Arial, sans-serif`;
        ctx.fillStyle = 'black';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        const x = area.x + area.width / 2;
        const y = area.y + area.height / 2;
        ctx.fillText(title, x, y);
        ctx.restore();
    }
    /**
     * Renders a bar or column chart.
     */
    renderBarChart(ctx, chartData, layout, isHorizontal) {
        if (chartData.categories.length === 0)
            return;
        const plotArea = layout.plotArea;
        const isStacked = chartData.type === 'stackedBar' || chartData.type === 'stackedColumn';
        // Calculate value range. Holes (blank cells) are ignored rather than
        // treated as 0 so they cannot distort the scale. When a secondary value
        // axis is rendered, the primary scale covers only primary-axis series.
        const minValue = 0;
        let maxValue;
        if (isStacked) {
            // For stacked charts, max is the sum of all series at each category;
            // a hole simply contributes nothing to the stack.
            maxValue = Math.max(...chartData.categories.map((_, i) => chartData.series.reduce((sum, s) => {
                const v = i < s.values.length ? s.values[i] : undefined;
                return sum + (v !== undefined && Number.isFinite(v) ? v : 0);
            }, 0)), 0);
        }
        else {
            maxValue = Math.max(...this.collectDefinedValues(this.primaryAxisSeries(chartData)), 0);
        }
        // Ensure we have a range
        if (maxValue <= minValue)
            maxValue = minValue + 1;
        maxValue *= 1.1; // Add 10% padding
        // Independent secondary-axis scale (undefined when single-axis)
        const secondaryMax = this.computeSecondaryMax(chartData);
        const axisMaxFor = (series) => series.secondaryAxis && secondaryMax !== undefined ? secondaryMax : maxValue;
        const categoryCount = chartData.categories.length;
        const seriesCount = chartData.series.length;
        ctx.save();
        if (isHorizontal) {
            // Horizontal bar chart
            const barHeight = plotArea.height / categoryCount;
            const barGroupHeight = barHeight * (1 - LAYOUT.barGap);
            const singleBarHeight = isStacked ? barGroupHeight : barGroupHeight / seriesCount;
            for (let catIndex = 0; catIndex < categoryCount; catIndex++) {
                const categoryY = plotArea.y + catIndex * barHeight + (barHeight - barGroupHeight) / 2;
                if (isStacked) {
                    let currentX = plotArea.x;
                    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        const series = chartData.series[seriesIndex];
                        if (!series)
                            continue;
                        const value = this.getDefinedValue(series, catIndex);
                        if (value === undefined)
                            continue; // Hole: no bar segment
                        const barWidth = (value / maxValue) * plotArea.width;
                        const color = this.getSeriesColor(series, seriesIndex);
                        ctx.fillStyle = this.rgbaToString(color);
                        ctx.fillRect(currentX, categoryY, barWidth, singleBarHeight);
                        currentX += barWidth;
                    }
                }
                else {
                    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        const series = chartData.series[seriesIndex];
                        if (!series)
                            continue;
                        const value = this.getDefinedValue(series, catIndex);
                        if (value === undefined)
                            continue; // Hole: no bar
                        const barWidth = (value / maxValue) * plotArea.width;
                        const color = this.getSeriesColor(series, seriesIndex);
                        const barY = categoryY + seriesIndex * singleBarHeight;
                        ctx.fillStyle = this.rgbaToString(color);
                        ctx.fillRect(plotArea.x, barY, barWidth, singleBarHeight * 0.9);
                    }
                }
            }
            // Render axes
            this.renderCategoryLabelsVertical(ctx, chartData.categories, layout.yAxisArea, plotArea);
            this.renderValueLabelsHorizontal(ctx, minValue, maxValue, layout.xAxisArea, plotArea);
        }
        else {
            // Vertical column chart
            const barWidth = plotArea.width / categoryCount;
            const barGroupWidth = barWidth * (1 - LAYOUT.barGap);
            const singleBarWidth = isStacked ? barGroupWidth : barGroupWidth / seriesCount;
            for (let catIndex = 0; catIndex < categoryCount; catIndex++) {
                const categoryX = plotArea.x + catIndex * barWidth + (barWidth - barGroupWidth) / 2;
                if (isStacked) {
                    let currentY = plotArea.y + plotArea.height;
                    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        const series = chartData.series[seriesIndex];
                        if (!series)
                            continue;
                        const value = this.getDefinedValue(series, catIndex);
                        if (value === undefined)
                            continue; // Hole: no bar segment
                        const barHeight = (value / maxValue) * plotArea.height;
                        const color = this.getSeriesColor(series, seriesIndex);
                        ctx.fillStyle = this.rgbaToString(color);
                        ctx.fillRect(categoryX, currentY - barHeight, singleBarWidth, barHeight);
                        currentY -= barHeight;
                    }
                }
                else {
                    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        const series = chartData.series[seriesIndex];
                        if (!series)
                            continue;
                        const value = this.getDefinedValue(series, catIndex);
                        if (value === undefined)
                            continue; // Hole: no bar
                        const barHeight = (value / axisMaxFor(series)) * plotArea.height;
                        const color = this.getSeriesColor(series, seriesIndex);
                        const barX = categoryX + seriesIndex * singleBarWidth;
                        ctx.fillStyle = this.rgbaToString(color);
                        ctx.fillRect(barX, plotArea.y + plotArea.height - barHeight, singleBarWidth * 0.9, barHeight);
                    }
                }
            }
            // Render axes
            this.renderCategoryLabelsHorizontal(ctx, chartData.categories, layout.xAxisArea, plotArea);
            this.renderValueLabelsVertical(ctx, minValue, maxValue, layout.yAxisArea, plotArea);
            if (secondaryMax !== undefined) {
                this.renderValueLabelsVertical(ctx, minValue, secondaryMax, layout.y2AxisArea, plotArea, 'right');
            }
        }
        // Draw axis lines
        this.drawAxisLines(ctx, plotArea, secondaryMax !== undefined);
        ctx.restore();
    }
    /**
     * Draws the X and Y axis lines along the plot area edges, plus the
     * secondary value-axis line on the right edge when requested.
     */
    drawAxisLines(ctx, plotArea, withSecondaryAxis = false) {
        ctx.strokeStyle = 'gray';
        ctx.lineWidth = 1;
        // X-axis
        ctx.beginPath();
        ctx.moveTo(plotArea.x, plotArea.y + plotArea.height);
        ctx.lineTo(plotArea.x + plotArea.width, plotArea.y + plotArea.height);
        ctx.stroke();
        // Y-axis
        ctx.beginPath();
        ctx.moveTo(plotArea.x, plotArea.y);
        ctx.lineTo(plotArea.x, plotArea.y + plotArea.height);
        ctx.stroke();
        // Secondary Y-axis on the right edge
        if (withSecondaryAxis) {
            ctx.beginPath();
            ctx.moveTo(plotArea.x + plotArea.width, plotArea.y);
            ctx.lineTo(plotArea.x + plotArea.width, plotArea.y + plotArea.height);
            ctx.stroke();
        }
    }
    /**
     * Renders a line chart. Holes in the data (blank cells) break the line:
     * no point and no connecting segment is drawn across a gap.
     */
    renderLineChart(ctx, chartData, layout) {
        if (chartData.categories.length === 0)
            return;
        const plotArea = layout.plotArea;
        // Calculate value range, ignoring holes; series on a secondary value
        // axis scale independently against their own maximum
        const minValue = 0;
        let maxValue = Math.max(...this.collectDefinedValues(this.primaryAxisSeries(chartData)), 0);
        if (maxValue <= minValue)
            maxValue = minValue + 1;
        maxValue *= 1.1;
        const secondaryMax = this.computeSecondaryMax(chartData);
        const categoryCount = chartData.categories.length;
        const pointSpacing = plotArea.width / Math.max(1, categoryCount - 1);
        ctx.save();
        // Draw each series
        for (let seriesIndex = 0; seriesIndex < chartData.series.length; seriesIndex++) {
            const series = chartData.series[seriesIndex];
            if (!series)
                continue;
            const color = this.getSeriesColor(series, seriesIndex);
            this.drawLineSeries(ctx, series, color, plotArea, {
                maxValue: series.secondaryAxis && secondaryMax !== undefined ? secondaryMax : maxValue,
                categoryCount,
                pointSpacing,
            });
        }
        // Render axes
        this.renderCategoryLabelsHorizontal(ctx, chartData.categories, layout.xAxisArea, plotArea);
        this.renderValueLabelsVertical(ctx, minValue, maxValue, layout.yAxisArea, plotArea);
        if (secondaryMax !== undefined) {
            this.renderValueLabelsVertical(ctx, minValue, secondaryMax, layout.y2AxisArea, plotArea, 'right');
        }
        // Draw axis lines
        this.drawAxisLines(ctx, plotArea, secondaryMax !== undefined);
        ctx.restore();
    }
    /**
     * Draws one line series as gap-separated polylines with point markers.
     * Each contiguous run of defined values becomes its own polyline; isolated
     * points render as a marker with no connecting line.
     */
    drawLineSeries(ctx, series, color, plotArea, scale) {
        const segments = segmentSeriesValues(series.values, scale.categoryCount);
        const xOffset = scale.xOffset ?? 0;
        for (const segment of segments) {
            const points = segment.map(({ index, value }) => ({
                x: plotArea.x + xOffset + index * scale.pointSpacing,
                y: plotArea.y + plotArea.height - (value / scale.maxValue) * plotArea.height,
            }));
            const firstPoint = points[0];
            if (points.length > 1 && firstPoint) {
                // Draw the polyline for this contiguous run
                ctx.strokeStyle = this.rgbaToString(color);
                ctx.lineWidth = 2;
                ctx.beginPath();
                ctx.moveTo(firstPoint.x, firstPoint.y);
                for (let i = 1; i < points.length; i++) {
                    const pt = points[i];
                    if (pt)
                        ctx.lineTo(pt.x, pt.y);
                }
                ctx.stroke();
            }
            // Draw point markers (also for isolated single points)
            ctx.fillStyle = this.rgbaToString(color);
            for (const point of points) {
                ctx.beginPath();
                ctx.arc(point.x, point.y, 4, 0, Math.PI * 2);
                ctx.fill();
            }
        }
    }
    /**
     * Renders a pie or doughnut chart. Holes in the data (blank cells) are
     * skipped entirely — no slice is drawn and the total excludes them.
     * Doughnut charts leave an inner cutout of holeSize percent (default 50)
     * of the outer radius empty and render multiple series as concentric
     * rings (first series innermost, matching PowerPoint), each ring
     * (outerR - holeR) / seriesCount thick. Multi-series pies render only
     * the first series.
     */
    renderPieChart(ctx, chartData, layout) {
        if (chartData.series.length === 0)
            return;
        const plotArea = layout.plotArea;
        const centerX = plotArea.x + plotArea.width / 2;
        const centerY = plotArea.y + plotArea.height / 2;
        const radius = (Math.min(plotArea.width, plotArea.height) / 2) * 0.9;
        if (chartData.type === 'doughnut') {
            const holeRadius = radius * ((chartData.holeSize ?? 50) / 100);
            const ringThickness = (radius - holeRadius) / chartData.series.length;
            ctx.save();
            for (let ringIndex = 0; ringIndex < chartData.series.length; ringIndex++) {
                const series = chartData.series[ringIndex];
                if (!series)
                    continue;
                const innerRadius = holeRadius + ringIndex * ringThickness;
                this.drawPieSeries(ctx, series, centerX, centerY, innerRadius + ringThickness, innerRadius);
            }
            ctx.restore();
            return;
        }
        if (chartData.series.length > 1) {
            // Multi-series pies have no concentric-ring form; only the first
            // series is drawn.
            this.logger.warn('Multi-series pie: rendering first series only', {
                seriesCount: chartData.series.length,
            });
            this.warnings?.push({
                code: 'unsupported-chart-type',
                message: `Multi-series pie chart: only the first of ${chartData.series.length} series was rendered`,
                detail: { seriesCount: chartData.series.length },
            });
        }
        const series = chartData.series[0];
        if (!series || series.values.length === 0)
            return;
        ctx.save();
        this.drawPieSeries(ctx, series, centerX, centerY, radius, 0);
        ctx.restore();
    }
    /**
     * Draws one series as pie slices between the given radii (full wedges
     * when innerRadius is 0, annular segments otherwise). Slice colors vary
     * by point index; holes in the data draw no slice and are excluded from
     * the series total.
     */
    drawPieSeries(ctx, series, centerX, centerY, outerRadius, innerRadius) {
        // Total over defined values only; holes contribute no slice
        const total = this.collectDefinedValues([series]).reduce((sum, v) => sum + v, 0);
        if (total <= 0)
            return;
        let currentAngle = -Math.PI / 2; // Start from top
        for (let i = 0; i < series.values.length; i++) {
            const value = series.values[i];
            if (value === undefined || !Number.isFinite(value))
                continue; // Hole: no slice
            const sweepAngle = (value / total) * Math.PI * 2;
            const color = this.getSeriesColor(undefined, i);
            // Draw pie slice (annular segment when there is an inner radius)
            ctx.fillStyle = this.rgbaToString(color);
            ctx.beginPath();
            if (innerRadius > 0) {
                ctx.arc(centerX, centerY, outerRadius, currentAngle, currentAngle + sweepAngle);
                ctx.arc(centerX, centerY, innerRadius, currentAngle + sweepAngle, currentAngle, true);
            }
            else {
                ctx.moveTo(centerX, centerY);
                ctx.arc(centerX, centerY, outerRadius, currentAngle, currentAngle + sweepAngle);
            }
            ctx.closePath();
            ctx.fill();
            // Draw slice border
            ctx.strokeStyle = 'white';
            ctx.lineWidth = 1;
            ctx.stroke();
            currentAngle += sweepAngle;
        }
    }
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
    renderScatterChart(ctx, chartData, layout) {
        const plotArea = layout.plotArea;
        const categoryCount = chartData.categories.length;
        const categoryMode = categoryCount > 0;
        // Largest bubble size across all series, for marker-radius scaling
        let maxBubbleSize = 0;
        for (const series of chartData.series) {
            for (const size of series.bubbleSizes ?? []) {
                if (Number.isFinite(size) && size > maxBubbleSize)
                    maxBubbleSize = size;
            }
        }
        const maxBubbleRadius = plotArea.width * ChartRenderer.MAX_BUBBLE_RADIUS_FRACTION;
        // Data-space points per series, aligned by data index; holes stay
        // undefined so connecting lines can break at gaps
        const seriesPoints = [];
        for (let seriesIndex = 0; seriesIndex < chartData.series.length; seriesIndex++) {
            const series = chartData.series[seriesIndex];
            if (!series)
                continue;
            const points = [];
            for (let i = 0; i < series.values.length; i++) {
                points.push(this.buildScatterPoint(series, i, categoryMode, maxBubbleSize, maxBubbleRadius));
            }
            seriesPoints.push({ seriesIndex, points });
        }
        const allPoints = seriesPoints.flatMap((sp) => sp.points.filter((p) => p !== undefined));
        if (allPoints.length === 0)
            return;
        // X range spans the data (category mode maps indices to band centers
        // instead); Y range starts at 0 like the other chart types
        let xMin = 0;
        let xMax = 1;
        if (!categoryMode) {
            xMin = Math.min(...allPoints.map((p) => p.x));
            xMax = Math.max(...allPoints.map((p) => p.x));
            const xPad = (xMax - xMin || 1) * 0.05;
            xMin -= xPad;
            xMax += xPad;
        }
        // Y range: anchor at 0 for positive data (matching the other chart
        // types) but extend to the data minimum when values are negative so
        // points never paint outside the plot area.
        let yMin = Math.min(0, ...allPoints.map((p) => p.y));
        let yMax = Math.max(...allPoints.map((p) => p.y), 0);
        if (yMax <= yMin)
            yMax = yMin + 1;
        const yPad = (yMax - yMin) * 0.1;
        yMax += yPad;
        if (yMin < 0)
            yMin -= yPad;
        const projectX = (x) => categoryMode
            ? // Clamp: series with more values than categories must not paint
                // beyond the plot area
                plotArea.x + ((Math.min(x, categoryCount - 1) + 0.5) / categoryCount) * plotArea.width
            : plotArea.x + ((x - xMin) / (xMax - xMin)) * plotArea.width;
        const projectY = (y) => plotArea.y + plotArea.height - ((y - yMin) / (yMax - yMin)) * plotArea.height;
        const style = chartData.scatterStyle;
        const drawLines = style === 'line' || style === 'lineMarker' || style === 'smooth' || style === 'smoothMarker';
        const drawMarkers = style !== 'line' && style !== 'smooth';
        ctx.save();
        for (const { seriesIndex, points } of seriesPoints) {
            const series = chartData.series[seriesIndex];
            const color = this.getSeriesColor(series, seriesIndex);
            if (drawLines) {
                this.drawScatterPolylines(ctx, points, color, projectX, projectY);
            }
            if (drawMarkers) {
                ctx.fillStyle = this.rgbaToString(color);
                for (const point of points) {
                    if (!point)
                        continue;
                    ctx.beginPath();
                    ctx.arc(projectX(point.x), projectY(point.y), point.radius, 0, Math.PI * 2);
                    ctx.fill();
                }
            }
        }
        // Render axes: category labels in category mode, numeric ticks otherwise
        this.renderValueLabelsVertical(ctx, yMin, yMax, layout.yAxisArea, plotArea);
        if (categoryMode) {
            this.renderCategoryLabelsHorizontal(ctx, chartData.categories, layout.xAxisArea, plotArea);
        }
        else {
            this.renderValueLabelsHorizontal(ctx, xMin, xMax, layout.xAxisArea, plotArea);
        }
        // Draw axis lines
        this.drawAxisLines(ctx, plotArea);
        ctx.restore();
    }
    /**
     * Maximum bubble-chart marker radius as a fraction of the plot width.
     */
    static MAX_BUBBLE_RADIUS_FRACTION = 0.08;
    /**
     * Default scatter marker radius in pixels (non-bubble series).
     */
    static SCATTER_MARKER_RADIUS = 4;
    /**
     * Builds the data-space point for one scatter/bubble entry, or undefined
     * when it is a hole: y or x missing/non-finite, or (for bubble series) a
     * blank or non-positive bubble size. In category mode x is always the
     * data index.
     */
    buildScatterPoint(series, index, categoryMode, maxBubbleSize, maxBubbleRadius) {
        const y = series.values[index];
        if (y === undefined || !Number.isFinite(y))
            return undefined;
        // Fall back to index-x when the series has no x values
        const x = !categoryMode && series.xValues ? series.xValues[index] : index;
        if (x === undefined || !Number.isFinite(x))
            return undefined;
        let radius = ChartRenderer.SCATTER_MARKER_RADIUS;
        if (series.bubbleSizes) {
            const size = series.bubbleSizes[index];
            // Blank and non-positive bubbles are hidden (PowerPoint's default)
            if (size === undefined || !Number.isFinite(size) || size <= 0 || maxBubbleSize <= 0) {
                return undefined;
            }
            radius = maxBubbleRadius * Math.sqrt(size / maxBubbleSize);
        }
        return { x, y, radius };
    }
    /**
     * Strokes straight polylines through contiguous runs of defined points
     * in data order. Holes break the line; runs of a single point draw
     * nothing (the marker pass covers them).
     */
    drawScatterPolylines(ctx, points, color, projectX, projectY) {
        ctx.strokeStyle = this.rgbaToString(color);
        ctx.lineWidth = 2;
        let run = [];
        const flushRun = () => {
            const first = run[0];
            if (run.length > 1 && first) {
                ctx.beginPath();
                ctx.moveTo(first.px, first.py);
                for (let i = 1; i < run.length; i++) {
                    const pt = run[i];
                    if (pt)
                        ctx.lineTo(pt.px, pt.py);
                }
                ctx.stroke();
            }
            run = [];
        };
        for (const point of points) {
            if (!point) {
                flushRun();
                continue;
            }
            run.push({ px: projectX(point.x), py: projectY(point.y) });
        }
        flushRun();
    }
    /**
     * Renders a combo chart: bar-kind series as clustered columns first, then
     * line-kind series on top. Series share the primary value-axis scale
     * unless flagged secondaryAxis, in which case they scale independently
     * against the right-hand secondary axis.
     */
    renderComboChart(ctx, chartData, layout) {
        const plotArea = layout.plotArea;
        // Category count falls back to the longest series when categories are absent
        const categoryCount = Math.max(chartData.categories.length, ...chartData.series.map((s) => s.values.length));
        if (categoryCount === 0)
            return;
        // Primary value-axis scale over the primary-axis series, ignoring holes
        const minValue = 0;
        let maxValue = Math.max(...this.collectDefinedValues(this.primaryAxisSeries(chartData)), 0);
        if (maxValue <= minValue)
            maxValue = minValue + 1;
        maxValue *= 1.1;
        // Independent secondary-axis scale (undefined when single-axis)
        const secondaryMax = this.computeSecondaryMax(chartData);
        const axisMaxFor = (series) => series.secondaryAxis && secondaryMax !== undefined ? secondaryMax : maxValue;
        const barSeriesIndices = [];
        const lineSeriesIndices = [];
        for (let i = 0; i < chartData.series.length; i++) {
            const series = chartData.series[i];
            if (!series)
                continue;
            if (this.getSeriesKind(series, chartData) === 'bar') {
                barSeriesIndices.push(i);
            }
            else {
                lineSeriesIndices.push(i);
            }
        }
        ctx.save();
        // Bars first
        const barWidth = plotArea.width / categoryCount;
        const barGroupWidth = barWidth * (1 - LAYOUT.barGap);
        const singleBarWidth = barGroupWidth / Math.max(1, barSeriesIndices.length);
        for (let catIndex = 0; catIndex < categoryCount; catIndex++) {
            const categoryX = plotArea.x + catIndex * barWidth + (barWidth - barGroupWidth) / 2;
            for (let slot = 0; slot < barSeriesIndices.length; slot++) {
                const seriesIndex = barSeriesIndices[slot];
                if (seriesIndex === undefined)
                    continue;
                const series = chartData.series[seriesIndex];
                if (!series)
                    continue;
                const value = this.getDefinedValue(series, catIndex);
                if (value === undefined)
                    continue; // Hole: no bar
                const barHeight = (value / axisMaxFor(series)) * plotArea.height;
                const color = this.getSeriesColor(series, seriesIndex);
                const barX = categoryX + slot * singleBarWidth;
                ctx.fillStyle = this.rgbaToString(color);
                ctx.fillRect(barX, plotArea.y + plotArea.height - barHeight, singleBarWidth * 0.9, barHeight);
            }
        }
        // Lines on top, scaled by their own value axis and centered in the same
        // category bands as the bar clusters so points align with their bars
        for (const seriesIndex of lineSeriesIndices) {
            const series = chartData.series[seriesIndex];
            if (!series)
                continue;
            const color = this.getSeriesColor(series, seriesIndex);
            this.drawLineSeries(ctx, series, color, plotArea, {
                maxValue: axisMaxFor(series),
                categoryCount,
                pointSpacing: barWidth,
                xOffset: barWidth / 2,
            });
        }
        // Render axes
        this.renderCategoryLabelsHorizontal(ctx, chartData.categories, layout.xAxisArea, plotArea);
        this.renderValueLabelsVertical(ctx, minValue, maxValue, layout.yAxisArea, plotArea);
        if (secondaryMax !== undefined) {
            this.renderValueLabelsVertical(ctx, minValue, secondaryMax, layout.y2AxisArea, plotArea, 'right');
        }
        this.drawAxisLines(ctx, plotArea, secondaryMax !== undefined);
        ctx.restore();
    }
    /**
     * Renders a neutral placeholder for unsupported chart types: a light gray
     * rounded rectangle with the chart title (or 'Chart') centered inside.
     */
    renderUnknownPlaceholder(ctx, chartData, bounds) {
        ctx.save();
        // White background matching regular chart rendering
        ctx.fillStyle = 'white';
        ctx.fillRect(bounds.x, bounds.y, bounds.width, bounds.height);
        // Light gray rounded rectangle inset from the bounds
        const inset = LAYOUT.padding;
        const x = bounds.x + inset;
        const y = bounds.y + inset;
        const width = Math.max(0, bounds.width - inset * 2);
        const height = Math.max(0, bounds.height - inset * 2);
        const cornerRadius = Math.min(8, width / 2, height / 2);
        ctx.fillStyle = 'rgb(242, 242, 242)';
        ctx.strokeStyle = 'rgb(200, 200, 200)';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(x + cornerRadius, y);
        ctx.arcTo(x + width, y, x + width, y + height, cornerRadius);
        ctx.arcTo(x + width, y + height, x, y + height, cornerRadius);
        ctx.arcTo(x, y + height, x, y, cornerRadius);
        ctx.arcTo(x, y, x + width, y, cornerRadius);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        // Centered label: chart title if present, else a generic 'Chart'
        ctx.font = `${LAYOUT.fontSize.title}px Calibri, Arial, sans-serif`;
        ctx.fillStyle = 'rgb(120, 120, 120)';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(chartData.title ?? 'Chart', x + width / 2, y + height / 2);
        ctx.restore();
    }
    /**
     * Chart types that render a secondary right-hand value axis. Horizontal
     * bar charts (value axis is horizontal), stacked charts (series share
     * one stack scale), and scatter charts keep single-axis rendering even
     * when series carry a secondaryAxis flag.
     */
    static SECONDARY_AXIS_TYPES = new Set([
        'column',
        'line',
        'area',
        'combo',
    ]);
    /**
     * Whether the chart renders a secondary right-hand value axis: at least
     * one series is bound to a secondary axis and the chart type supports
     * drawing one.
     */
    hasSecondaryValueAxis(chartData) {
        return (ChartRenderer.SECONDARY_AXIS_TYPES.has(chartData.type) &&
            chartData.series.some((s) => s.secondaryAxis === true));
    }
    /**
     * The series scaled by the primary value axis: every series when no
     * secondary axis is rendered, otherwise only those not flagged secondary.
     */
    primaryAxisSeries(chartData) {
        if (!this.hasSecondaryValueAxis(chartData))
            return chartData.series;
        return chartData.series.filter((s) => !s.secondaryAxis);
    }
    /**
     * Computes the secondary value-axis maximum (with the standard positive
     * floor and 10% headroom, matching the primary axis) over the
     * secondary-axis series, or undefined when the chart renders no
     * secondary axis.
     */
    computeSecondaryMax(chartData) {
        if (!this.hasSecondaryValueAxis(chartData))
            return undefined;
        const secondarySeries = chartData.series.filter((s) => s.secondaryAxis === true);
        let max = Math.max(...this.collectDefinedValues(secondarySeries), 0);
        if (max <= 0)
            max = 1;
        return max * 1.1;
    }
    /**
     * Collects every defined, finite value across the given series. Holes
     * (sparse slots, explicit undefined, NaN) are excluded so scale
     * calculations ignore them instead of treating them as 0.
     */
    collectDefinedValues(series) {
        const values = [];
        for (const s of series) {
            for (const v of s.values) {
                if (v !== undefined && Number.isFinite(v))
                    values.push(v);
            }
        }
        return values;
    }
    /**
     * Returns the series value at the given category index, or undefined when
     * the slot is a hole (blank cell) or out of range.
     */
    getDefinedValue(series, catIndex) {
        if (catIndex >= series.values.length)
            return undefined;
        const value = series.values[catIndex];
        return value !== undefined && Number.isFinite(value) ? value : undefined;
    }
    /**
     * Resolves the plot kind of a series, defaulting to a kind derived from
     * the overall chart type when the series carries none.
     */
    getSeriesKind(series, chartData) {
        if (series.chartKind)
            return series.chartKind;
        switch (chartData.type) {
            case 'bar':
            case 'column':
            case 'stackedBar':
            case 'stackedColumn':
                return 'bar';
            case 'pie':
                return 'pie';
            case 'doughnut':
                return 'doughnut';
            case 'area':
                return 'area';
            case 'scatter':
                return 'scatter';
            default:
                return 'line';
        }
    }
    /**
     * Renders the legend.
     */
    renderLegend(ctx, chartData, area) {
        const boxSize = 12;
        const spacing = 10;
        const itemWidth = 80;
        let x = area.x + spacing;
        let y = area.y + (area.height - boxSize) / 2;
        ctx.save();
        for (let i = 0; i < chartData.series.length; i++) {
            const series = chartData.series[i];
            if (!series)
                continue;
            const color = this.getSeriesColor(series, i);
            // Draw color box
            ctx.fillStyle = this.rgbaToString(color);
            ctx.fillRect(x, y, boxSize, boxSize);
            // Draw label
            ctx.font = `${LAYOUT.fontSize.legend}px Arial, sans-serif`;
            ctx.fillStyle = 'black';
            ctx.textBaseline = 'middle';
            ctx.textAlign = 'left';
            ctx.fillText(this.truncateLabel(series.name, 10), x + boxSize + 4, y + boxSize / 2);
            x += itemWidth;
            if (x + itemWidth > area.x + area.width) {
                x = area.x + spacing;
                y += boxSize + 4;
            }
        }
        ctx.restore();
    }
    /**
     * Renders horizontal category labels (for column charts).
     */
    renderCategoryLabelsHorizontal(ctx, categories, area, plotArea) {
        if (categories.length === 0)
            return;
        const labelWidth = plotArea.width / categories.length;
        ctx.save();
        ctx.font = `${LAYOUT.fontSize.axis}px Arial, sans-serif`;
        ctx.fillStyle = 'gray';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        for (let i = 0; i < categories.length; i++) {
            const x = plotArea.x + i * labelWidth + labelWidth / 2;
            const y = area.y + area.height / 2;
            ctx.fillText(this.truncateLabel(categories[i] ?? '', 10), x, y);
        }
        ctx.restore();
    }
    /**
     * Renders vertical category labels (for bar charts).
     */
    renderCategoryLabelsVertical(ctx, categories, area, plotArea) {
        if (categories.length === 0)
            return;
        const labelHeight = plotArea.height / categories.length;
        ctx.save();
        ctx.font = `${LAYOUT.fontSize.axis}px Arial, sans-serif`;
        ctx.fillStyle = 'gray';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        for (let i = 0; i < categories.length; i++) {
            const x = area.x + area.width - 4;
            const y = plotArea.y + i * labelHeight + labelHeight / 2;
            ctx.fillText(this.truncateLabel(categories[i] ?? '', 8), x, y);
        }
        ctx.restore();
    }
    /**
     * Renders vertical value labels (Y-axis for column/line charts). The
     * 'left' side right-aligns labels against the plot area (primary axis);
     * 'right' left-aligns them (secondary axis on the right of the plot).
     */
    renderValueLabelsVertical(ctx, minValue, maxValue, area, plotArea, side = 'left') {
        const tickCount = 5;
        const step = (maxValue - minValue) / tickCount;
        ctx.save();
        ctx.font = `${LAYOUT.fontSize.axis}px Arial, sans-serif`;
        ctx.fillStyle = 'gray';
        ctx.textAlign = side === 'left' ? 'right' : 'left';
        ctx.textBaseline = 'middle';
        const x = side === 'left' ? area.x + area.width - 4 : area.x + 4;
        for (let i = 0; i <= tickCount; i++) {
            const value = minValue + ((maxValue - minValue) * i) / tickCount;
            const y = plotArea.y + plotArea.height - (i * plotArea.height) / tickCount;
            ctx.fillText(this.formatValue(value, step), x, y);
        }
        ctx.restore();
    }
    /**
     * Renders horizontal value labels (X-axis for bar charts).
     */
    renderValueLabelsHorizontal(ctx, minValue, maxValue, area, plotArea) {
        const tickCount = 5;
        const step = (maxValue - minValue) / tickCount;
        ctx.save();
        ctx.font = `${LAYOUT.fontSize.axis}px Arial, sans-serif`;
        ctx.fillStyle = 'gray';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        for (let i = 0; i <= tickCount; i++) {
            const value = minValue + ((maxValue - minValue) * i) / tickCount;
            const x = plotArea.x + (i * plotArea.width) / tickCount;
            const y = area.y + area.height / 2;
            ctx.fillText(this.formatValue(value, step), x, y);
        }
        ctx.restore();
    }
    /**
     * Gets the color for a series.
     */
    getSeriesColor(series, index) {
        if (series?.color) {
            return series.color;
        }
        const defaultColor = DEFAULT_SERIES_COLORS[index % DEFAULT_SERIES_COLORS.length];
        // Fallback to first color if somehow undefined (should never happen)
        return defaultColor ?? DEFAULT_SERIES_COLORS[0] ?? { r: 68, g: 114, b: 196, a: 255 };
    }
    /**
     * Converts RGBA to CSS color string.
     */
    rgbaToString(color) {
        return `rgba(${color.r}, ${color.g}, ${color.b}, ${color.a / 255})`;
    }
    /**
     * Truncates a label to max length.
     */
    truncateLabel(label, maxLength) {
        if (!label)
            return '';
        if (label.length <= maxLength)
            return label;
        return label.substring(0, maxLength - 1) + '...';
    }
    /**
     * Formats a value-axis tick label. Precision derives from the tick step
     * so fractional steps keep enough decimals to stay distinct (step 0.25
     * -> "0.25") while integer steps render clean whole numbers, and
     * floating-point noise (e.g. 0.30000000000000004) never leaks into a
     * label. Large magnitudes use K/M suffixes.
     */
    formatValue(value, step) {
        const abs = Math.abs(value);
        if (abs >= 1_000_000) {
            return this.formatWithStepPrecision(value / 1_000_000, step / 1_000_000) + 'M';
        }
        if (abs >= 1_000) {
            return this.formatWithStepPrecision(value / 1_000, step / 1_000) + 'K';
        }
        return this.formatWithStepPrecision(value, step);
    }
    /**
     * Formats a number with just enough decimals to resolve the given step,
     * trimming trailing zeros (step 0.25 -> 2 decimals, step 2.2 -> 1,
     * step 22 -> 0).
     */
    formatWithStepPrecision(value, step) {
        let decimals = 0;
        if (Number.isFinite(step) && step > 0) {
            decimals = Math.min(6, Math.max(0, 1 - Math.floor(Math.log10(step))));
        }
        return String(parseFloat(value.toFixed(decimals)));
    }
}
/**
 * Creates a ChartRenderer instance.
 */
export function createChartRenderer(logger) {
    return new ChartRenderer({ logger });
}
//# sourceMappingURL=ChartRenderer.js.map