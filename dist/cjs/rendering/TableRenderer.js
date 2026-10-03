"use strict";
/**
 * Renders tables from PPTX graphicFrame elements.
 * Tables are contained in a:graphic/a:graphicData[@uri="...table"]/a:tbl.
 *
 * Handles:
 * - Column widths from a:tblGrid/a:gridCol[@w]
 * - Row heights from a:tr[@h]
 * - Cell backgrounds from a:tc/a:tcPr/a:solidFill and a:gradFill (linear
 *   gradients render as true gradients in the cell rect; radial/path
 *   gradients are approximated by their lowest-position stop)
 * - Cell borders from a:tc/a:tcPr/a:ln*, including the diagonals
 *   a:lnTlToBr / a:lnBlToTr stroked corner-to-corner
 * - Cell text from a:tc/a:txBody (using TextParser and TextRenderer)
 * - Cell margins from a:tc/a:tcPr[@marL/marR/marT/marB]
 * - Table styles from a:tblPr/a:tableStyleId resolved against
 *   ppt/tableStyles.xml (see TableStyleParser): the a:tblBg background is
 *   painted beneath all cell fills, then per-cell fills, borders (outer
 *   edges vs insideH/insideV), and text defaults are layered as
 *   wholeTbl < band1H/band2H < band1V/band2V < lastCol < firstCol <
 *   lastRow < seCell/swCell < firstRow < neCell/nwCell, with explicit tcPr
 *   formatting on top.
 * - Right-to-left tables (a:tblPr rtl="1"): columns are mirrored
 *   horizontally, and each cell's sided borders/diagonals mirror with them.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.TableRenderer = void 0;
exports.isGradientCellFill = isGradientCellFill;
exports.mirrorRectHorizontally = mirrorRectHorizontally;
exports.createTableRenderer = createTableRenderer;
const geometry_js_1 = require("../types/geometry.js");
const PptxParser_js_1 = require("../core/PptxParser.js");
const UnitConverter_js_1 = require("../core/UnitConverter.js");
const ColorResolver_js_1 = require("../theme/ColorResolver.js");
const TextParser_js_1 = require("../parsers/TextParser.js");
const TableStyleParser_js_1 = require("../parsers/TableStyleParser.js");
const FillRenderer_js_1 = require("./FillRenderer.js");
const TextRenderer_js_1 = require("./TextRenderer.js");
const Logger_js_1 = require("../utils/Logger.js");
/**
 * Default cell margin in EMU (91440 EMU = 0.1 inches).
 */
const DEFAULT_CELL_MARGIN_EMU = 91440;
/**
 * Default top/bottom cell margin in EMU (45720 EMU = 0.05 inches,
 * ECMA-376 CT_TableCellProperties marT/marB defaults).
 */
const DEFAULT_CELL_VERTICAL_MARGIN_EMU = 45720;
/**
 * Default border width in EMU (12700 EMU = 1 point).
 */
const DEFAULT_BORDER_WIDTH_EMU = 12700;
/**
 * Preset dash patterns mapped to canvas dash arrays.
 * Values are relative to stroke width (kept in sync with StrokeRenderer).
 */
const PRESET_DASH_PATTERNS = {
    solid: [],
    dot: [1, 2],
    dash: [4, 3],
    lgDash: [8, 3],
    dashDot: [4, 3, 1, 3],
    lgDashDot: [8, 3, 1, 3],
    lgDashDotDot: [8, 3, 1, 3, 1, 3],
    sysDash: [3, 1],
    sysDot: [1, 1],
    sysDashDot: [3, 1, 1, 1],
    sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
/**
 * Narrows a cell fill to its gradient variant.
 * @param fill The cell fill
 * @returns True when the fill is a linear gradient
 */
function isGradientCellFill(fill) {
    return 'type' in fill && fill.type === 'gradient';
}
/**
 * Renders tables to canvas.
 */
class TableRenderer {
    logger;
    theme;
    scaleX;
    scaleY;
    colorResolver;
    unitConverter;
    textParser;
    textRenderer;
    tableStyles;
    constructor(config) {
        this.logger = config.logger ?? (0, Logger_js_1.createLogger)('warn', 'TableRenderer');
        this.theme = config.theme;
        this.scaleX = config.scaleX;
        this.scaleY = config.scaleY;
        this.tableStyles = config.tableStyles;
        this.colorResolver = new ColorResolver_js_1.ColorResolver(config.theme.colors, config.colorMap);
        this.unitConverter = new UnitConverter_js_1.UnitConverter();
        this.textParser = new TextParser_js_1.TextParser({
            theme: config.theme,
            colorMap: config.colorMap,
            slideNumber: config.slideNumber,
            logger: this.logger.child?.('TextParser'),
        });
        this.textRenderer = new TextRenderer_js_1.TextRenderer({
            theme: config.theme,
            scaleX: config.scaleX,
            scaleY: config.scaleY,
            logger: this.logger.child?.('TextRenderer'),
            warnings: config.warnings,
        });
    }
    /**
     * Renders a table to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param tableNode The a:tbl XML node
     * @param bounds The bounds to render within (in pixels)
     */
    renderTable(ctx, tableNode, bounds) {
        // Parse the table structure
        const table = this.parseTable(tableNode);
        if (!table || table.rows.length === 0) {
            this.logger.debug('No table data to render');
            return;
        }
        this.logger.debug('Rendering table', {
            rows: table.rows.length,
            columns: table.columnWidths.length,
            styleId: table.properties.tableStyleId,
            bounds,
        });
        // Resolve the referenced table style, if any
        const tableStyle = table.properties.tableStyleId
            ? this.lookupTableStyle(table.properties.tableStyleId)
            : undefined;
        ctx.save();
        // Calculate pixel positions for each cell
        const columnPixels = this.calculatePixelPositions(table.columnWidths, bounds.x, bounds.width);
        const rowPixels = this.calculatePixelPositions(table.rowHeights, bounds.y, bounds.height);
        // Paint the table-style background (a:tblBg) beneath all cell fills
        const tblBg = tableStyle?.tblBg;
        if (tblBg && tblBg.type !== 'none') {
            this.paintFill(ctx, tblBg.type === 'solid' ? tblBg.color : tblBg, bounds);
        }
        // RTL tables mirror the column layout horizontally; styles are computed
        // in logical grid space, then each cell rect is mirrored at paint time
        const rtl = table.properties.rtl === true;
        // Render each cell
        for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
            const row = table.rows[rowIndex];
            if (!row)
                continue;
            for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
                const cell = row.cells[colIndex];
                if (!cell)
                    continue;
                // Skip merged cells
                if (cell.hMerge || cell.vMerge) {
                    continue;
                }
                // Calculate cell bounds
                let cellBounds = this.calculateCellBounds(columnPixels, rowPixels, colIndex, rowIndex, cell.colSpan, cell.rowSpan);
                if (rtl) {
                    cellBounds = mirrorRectHorizontally(cellBounds, bounds);
                }
                // Compute the effective table-style formatting for this cell
                const styled = tableStyle
                    ? this.computeEffectiveCellStyle(tableStyle, table.properties, rowIndex, colIndex, cell.rowSpan, cell.colSpan, table.rows.length, table.columnWidths.length)
                    : undefined;
                // Render cell
                this.renderCell(ctx, cell, cellBounds, styled, rtl);
            }
        }
        ctx.restore();
    }
    /**
     * Looks up a table style by its GUID in the configured style collection.
     * PowerPoint frequently references built-in styles (e.g. Medium Style 2
     * Accent 1, {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}) that tableStyles.xml
     * does not define; those render unstyled with a debug log.
     * @param styleId The style GUID from a:tableStyleId
     * @returns The parsed style, or undefined when it is not defined
     */
    lookupTableStyle(styleId) {
        const style = this.tableStyles?.styles.get(styleId);
        if (!style) {
            this.logger.debug('Table style not found in tableStyles.xml; rendering unstyled', {
                styleId,
            });
            return undefined;
        }
        return style;
    }
    /**
     * Computes the effective table-style formatting for a cell by layering the
     * style's conditional formatting parts in precedence order (low to high),
     * following the CT_TableStyle element sequence (ECMA-376 §20.1.4.2.26):
     * wholeTbl < band1V/band2V < band1H/band2H < lastCol < firstCol <
     * lastRow < seCell/swCell < firstRow < neCell/nwCell. Parts apply only
     * when the corresponding a:tblPr flag is enabled; corner parts require
     * both of their row and column flags (e.g. nwCell needs firstRow and
     * firstCol).
     *
     * Band indexing: data rows are counted from after the header row when
     * firstRow is enabled; the first data row gets band1H, the second band2H,
     * alternating. The last row is excluded from banding when lastRow is
     * enabled. Column banding is analogous with firstCol/lastCol.
     *
     * @param style The resolved table style
     * @param props The table's a:tblPr flags
     * @param rowIndex Cell row index (grid row of the cell's top edge)
     * @param colIndex Cell column index (grid column of the cell's left edge)
     * @param rowSpan Cell row span
     * @param colSpan Cell column span
     * @param rowCount Total number of table rows
     * @param colCount Total number of table columns
     * @returns The effective cell style
     */
    computeEffectiveCellStyle(style, props, rowIndex, colIndex, rowSpan, colSpan, rowCount, colCount) {
        const lastRowIdx = rowCount - 1;
        const lastColIdx = colCount - 1;
        const cellTop = rowIndex;
        const cellBottom = Math.min(rowIndex + rowSpan - 1, lastRowIdx);
        const cellLeft = colIndex;
        const cellRight = Math.min(colIndex + colSpan - 1, lastColIdx);
        const layers = [];
        if (style.wholeTbl) {
            layers.push({ part: style.wholeTbl, top: 0, bottom: lastRowIdx, left: 0, right: lastColIdx });
        }
        // Row banding (below column banding per the CT_TableStyle sequence)
        if (props.bandRow) {
            const headerRows = props.firstRow ? 1 : 0;
            // Match the lastRow part gate (cellBottom); see inLastCol above
            const inLastRow = props.lastRow === true && cellBottom === lastRowIdx;
            if (cellTop >= headerRows && !inLastRow) {
                const dataRowIndex = cellTop - headerRows;
                const band = dataRowIndex % 2 === 0 ? style.band1H : style.band2H;
                if (band) {
                    layers.push({ part: band, top: cellTop, bottom: cellBottom, left: 0, right: lastColIdx });
                }
            }
        }
        // Column banding (above row banding: band1V/band2V follow band1H/band2H
        // in the CT_TableStyle sequence, and later parts layer higher)
        if (props.bandCol) {
            const headerCols = props.firstCol ? 1 : 0;
            // Match the lastCol part gate (cellRight) so a spanning cell never
            // receives both a band layer and the lastCol layer
            const inLastCol = props.lastCol === true && cellRight === lastColIdx;
            if (cellLeft >= headerCols && !inLastCol) {
                const dataColIndex = cellLeft - headerCols;
                const band = dataColIndex % 2 === 0 ? style.band1V : style.band2V;
                if (band) {
                    layers.push({ part: band, top: 0, bottom: lastRowIdx, left: cellLeft, right: cellRight });
                }
            }
        }
        if (props.lastCol && cellRight === lastColIdx && style.lastCol) {
            layers.push({
                part: style.lastCol,
                top: 0,
                bottom: lastRowIdx,
                left: lastColIdx,
                right: lastColIdx,
            });
        }
        if (props.firstCol && cellLeft === 0 && style.firstCol) {
            layers.push({ part: style.firstCol, top: 0, bottom: lastRowIdx, left: 0, right: 0 });
        }
        if (props.lastRow && cellBottom === lastRowIdx && style.lastRow) {
            layers.push({
                part: style.lastRow,
                top: lastRowIdx,
                bottom: lastRowIdx,
                left: 0,
                right: lastColIdx,
            });
        }
        // Corner parts cover their single corner cell and require both of the
        // corresponding row/column flags (schema order: seCell < swCell <
        // firstRow < neCell < nwCell)
        const inLastRow = props.lastRow === true && cellBottom === lastRowIdx;
        const inFirstRow = props.firstRow === true && cellTop === 0;
        const inFirstCol = props.firstCol === true && cellLeft === 0;
        const inLastCol = props.lastCol === true && cellRight === lastColIdx;
        if (inLastRow && inLastCol && style.seCell) {
            layers.push({
                part: style.seCell,
                top: lastRowIdx,
                bottom: lastRowIdx,
                left: lastColIdx,
                right: lastColIdx,
            });
        }
        if (inLastRow && inFirstCol && style.swCell) {
            layers.push({ part: style.swCell, top: lastRowIdx, bottom: lastRowIdx, left: 0, right: 0 });
        }
        if (props.firstRow && cellTop === 0 && style.firstRow) {
            layers.push({ part: style.firstRow, top: 0, bottom: 0, left: 0, right: lastColIdx });
        }
        if (inFirstRow && inLastCol && style.neCell) {
            layers.push({
                part: style.neCell,
                top: 0,
                bottom: 0,
                left: lastColIdx,
                right: lastColIdx,
            });
        }
        if (inFirstRow && inFirstCol && style.nwCell) {
            layers.push({ part: style.nwCell, top: 0, bottom: 0, left: 0, right: 0 });
        }
        // Merge the layers: higher-precedence parts override per property
        const effective = { borders: {} };
        const borders = {};
        let fill;
        for (const layer of layers) {
            if (layer.part.fill) {
                fill = layer.part.fill;
            }
            // Region-relative border selection: a cell at the part's region edge
            // takes the part's outer border, inner edges take insideH/insideV.
            // Diagonals have no inside variant and apply to every cell in the
            // part's region.
            const b = layer.part.borders;
            const left = cellLeft === layer.left ? b.left : b.insideV;
            const right = cellRight === layer.right ? b.right : b.insideV;
            const top = cellTop === layer.top ? b.top : b.insideH;
            const bottom = cellBottom === layer.bottom ? b.bottom : b.insideH;
            if (left)
                borders.left = left;
            if (right)
                borders.right = right;
            if (top)
                borders.top = top;
            if (bottom)
                borders.bottom = bottom;
            if (b.tlToBr)
                borders.tlToBr = b.tlToBr;
            if (b.blToTr)
                borders.blToTr = b.blToTr;
            const t = layer.part.text;
            if (t.bold !== undefined)
                effective.bold = t.bold;
            if (t.italic !== undefined)
                effective.italic = t.italic;
            if (t.color !== undefined)
                effective.color = t.color;
            if (t.fontFamily !== undefined)
                effective.fontFamily = t.fontFamily;
        }
        if (fill?.type === 'solid') {
            effective.fill = fill.color;
        }
        else if (fill?.type === 'gradient') {
            effective.fill = fill;
        }
        effective.borders = {
            left: toParsedCellBorder(borders.left),
            right: toParsedCellBorder(borders.right),
            top: toParsedCellBorder(borders.top),
            bottom: toParsedCellBorder(borders.bottom),
            tlToBr: toParsedCellBorder(borders.tlToBr),
            blToTr: toParsedCellBorder(borders.blToTr),
        };
        return effective;
    }
    /**
     * Parses a table XML node into a structured format.
     */
    parseTable(tableNode) {
        if (!tableNode)
            return undefined;
        // Parse table properties
        const tblPr = (0, PptxParser_js_1.getXmlChild)(tableNode, 'a:tblPr');
        const properties = this.parseTableProperties(tblPr);
        // Parse column widths from tblGrid
        const tblGrid = (0, PptxParser_js_1.getXmlChild)(tableNode, 'a:tblGrid');
        const columnWidths = this.parseColumnWidths(tblGrid);
        if (columnWidths.length === 0) {
            this.logger.debug('No columns defined in table');
            return undefined;
        }
        // Parse rows
        const trNodes = (0, PptxParser_js_1.getXmlChildren)(tableNode, 'a:tr');
        const rows = [];
        for (const trNode of trNodes) {
            const row = this.parseRow(trNode, columnWidths.length);
            if (row) {
                rows.push(row);
            }
        }
        // Extract row heights
        const rowHeights = rows.map((r) => r.height);
        return {
            columnWidths,
            rowHeights,
            rows,
            properties,
        };
    }
    /**
     * Parses table properties (a:tblPr flags and the a:tableStyleId child).
     */
    parseTableProperties(tblPr) {
        if (!tblPr)
            return {};
        const flag = (name) => {
            const value = (0, PptxParser_js_1.getXmlAttr)(tblPr, name);
            return value === '1' || value === 'true';
        };
        return {
            firstRow: flag('firstRow'),
            firstCol: flag('firstCol'),
            lastRow: flag('lastRow'),
            lastCol: flag('lastCol'),
            bandRow: flag('bandRow'),
            bandCol: flag('bandCol'),
            rtl: flag('rtl'),
            tableStyleId: parseTableStyleId(tblPr),
        };
    }
    /**
     * Parses column widths from tblGrid.
     */
    parseColumnWidths(tblGrid) {
        if (!tblGrid)
            return [];
        const gridCols = (0, PptxParser_js_1.getXmlChildren)(tblGrid, 'a:gridCol');
        const widths = [];
        for (const gridCol of gridCols) {
            const w = (0, PptxParser_js_1.getXmlAttr)(gridCol, 'w');
            if (w) {
                widths.push(parseInt(w, 10));
            }
        }
        return widths;
    }
    /**
     * Parses a table row.
     */
    parseRow(trNode, _columnCount) {
        // Get row height
        const hAttr = (0, PptxParser_js_1.getXmlAttr)(trNode, 'h');
        const height = hAttr ? parseInt(hAttr, 10) : 0;
        // Parse cells
        const tcNodes = (0, PptxParser_js_1.getXmlChildren)(trNode, 'a:tc');
        const cells = [];
        for (const tcNode of tcNodes) {
            const cell = this.parseCell(tcNode);
            cells.push(cell);
        }
        return { height, cells };
    }
    /**
     * Parses a table cell.
     */
    parseCell(tcNode) {
        // Parse cell properties
        const tcPr = (0, PptxParser_js_1.getXmlChild)(tcNode, 'a:tcPr');
        // Parse text body
        const txBody = (0, PptxParser_js_1.getXmlChild)(tcNode, 'a:txBody');
        const textBody = this.textParser.parseTextBody(txBody);
        // Parse background color
        const backgroundColor = this.parseCellBackground(tcPr);
        // Parse margins
        const margins = this.parseCellMargins(tcPr);
        // Parse borders
        const borders = this.parseCellBorders(tcPr);
        // Parse span attributes
        const gridSpan = (0, PptxParser_js_1.getXmlAttr)(tcNode, 'gridSpan');
        const rowSpan = (0, PptxParser_js_1.getXmlAttr)(tcNode, 'rowSpan');
        const hMerge = (0, PptxParser_js_1.getXmlAttr)(tcNode, 'hMerge') === '1';
        const vMerge = (0, PptxParser_js_1.getXmlAttr)(tcNode, 'vMerge') === '1';
        // Parse anchor (vertical alignment)
        let anchor;
        if (tcPr) {
            const anchorAttr = (0, PptxParser_js_1.getXmlAttr)(tcPr, 'anchor');
            if (anchorAttr === 't')
                anchor = 'top';
            else if (anchorAttr === 'ctr')
                anchor = 'middle';
            else if (anchorAttr === 'b')
                anchor = 'bottom';
        }
        // Also check body properties for anchor
        if (!anchor && txBody) {
            const bodyPr = (0, PptxParser_js_1.getXmlChild)(txBody, 'a:bodyPr');
            if (bodyPr) {
                const anchorAttr = (0, PptxParser_js_1.getXmlAttr)(bodyPr, 'anchor');
                if (anchorAttr === 't')
                    anchor = 'top';
                else if (anchorAttr === 'ctr')
                    anchor = 'middle';
                else if (anchorAttr === 'b')
                    anchor = 'bottom';
            }
        }
        return {
            textBody,
            backgroundColor,
            margins,
            borders,
            colSpan: gridSpan ? parseInt(gridSpan, 10) : 1,
            rowSpan: rowSpan ? parseInt(rowSpan, 10) : 1,
            hMerge,
            vMerge,
            anchor,
        };
    }
    /**
     * Parses the cell background fill. Returns 'none' for an explicit
     * a:noFill, which suppresses any table-style fill. Linear a:gradFill
     * fills are preserved as gradients; radial/path gradients degrade to
     * their lowest-position stop.
     */
    parseCellBackground(tcPr) {
        if (!tcPr)
            return undefined;
        // Check for solid fill
        const solidFill = (0, PptxParser_js_1.getXmlChild)(tcPr, 'a:solidFill');
        if (solidFill) {
            return this.colorResolver.resolveColorElement(solidFill);
        }
        // Check for no fill (transparent). hasXmlChild: a self-closed
        // <a:noFill/> parses to '' which is falsy.
        if ((0, PptxParser_js_1.hasXmlChild)(tcPr, 'a:noFill')) {
            return 'none';
        }
        // Explicit gradient fill; any other explicit fill (pattern/picture)
        // suppresses the table-style fill rather than letting a band color
        // leak through.
        const gradFill = (0, PptxParser_js_1.getXmlChild)(tcPr, 'a:gradFill');
        if (gradFill && typeof gradFill === 'object') {
            const fill = (0, TableStyleParser_js_1.parseTableGradientFill)(gradFill, this.colorResolver, this.logger);
            if (!fill)
                return 'none';
            return fill.type === 'solid' ? fill.color : fill;
        }
        if ((0, PptxParser_js_1.hasXmlChild)(tcPr, 'a:gradFill') ||
            (0, PptxParser_js_1.hasXmlChild)(tcPr, 'a:pattFill') ||
            (0, PptxParser_js_1.hasXmlChild)(tcPr, 'a:blipFill')) {
            return 'none';
        }
        return undefined;
    }
    /**
     * Parses cell margins.
     */
    parseCellMargins(tcPr) {
        const defaultMargin = DEFAULT_CELL_MARGIN_EMU;
        if (!tcPr) {
            return {
                left: defaultMargin,
                right: defaultMargin,
                top: DEFAULT_CELL_VERTICAL_MARGIN_EMU,
                bottom: DEFAULT_CELL_VERTICAL_MARGIN_EMU,
            };
        }
        const marL = (0, PptxParser_js_1.getXmlAttr)(tcPr, 'marL');
        const marR = (0, PptxParser_js_1.getXmlAttr)(tcPr, 'marR');
        const marT = (0, PptxParser_js_1.getXmlAttr)(tcPr, 'marT');
        const marB = (0, PptxParser_js_1.getXmlAttr)(tcPr, 'marB');
        return {
            left: marL !== undefined ? parseInt(marL, 10) : defaultMargin,
            right: marR !== undefined ? parseInt(marR, 10) : defaultMargin,
            top: marT !== undefined ? parseInt(marT, 10) : DEFAULT_CELL_VERTICAL_MARGIN_EMU,
            bottom: marB !== undefined ? parseInt(marB, 10) : DEFAULT_CELL_VERTICAL_MARGIN_EMU,
        };
    }
    /**
     * Parses cell borders.
     */
    parseCellBorders(tcPr) {
        if (!tcPr)
            return {};
        const borders = {};
        // Parse each border direction
        borders.left = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnL'));
        borders.right = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnR'));
        borders.top = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnT'));
        borders.bottom = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnB'));
        borders.tlToBr = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnTlToBr'));
        borders.blToTr = this.parseBorder((0, PptxParser_js_1.getXmlChild)(tcPr, 'a:lnBlToTr'));
        return borders;
    }
    /**
     * Parses a single border line. Returns 'none' for an explicit a:noFill,
     * which suppresses any table-style border.
     */
    parseBorder(lnNode) {
        if (!lnNode || typeof lnNode !== 'object')
            return undefined;
        // Check for no fill (no border)
        if ((0, PptxParser_js_1.hasXmlChild)(lnNode, 'a:noFill')) {
            return 'none';
        }
        // Get border width
        const wAttr = (0, PptxParser_js_1.getXmlAttr)(lnNode, 'w');
        const width = wAttr ? parseInt(wAttr, 10) : DEFAULT_BORDER_WIDTH_EMU;
        // Get border color from solid fill
        const solidFill = (0, PptxParser_js_1.getXmlChild)(lnNode, 'a:solidFill');
        const color = solidFill
            ? this.colorResolver.resolveColorElement(solidFill)
            : { ...geometry_js_1.Colors.black }; // Default to black
        if (!color)
            return undefined;
        // Get preset dash
        const prstDash = (0, PptxParser_js_1.getXmlChild)(lnNode, 'a:prstDash');
        const dash = prstDash ? (0, PptxParser_js_1.getXmlAttr)(prstDash, 'val') : undefined;
        return { width, color, dash };
    }
    /**
     * Calculates pixel positions from EMU sizes.
     */
    calculatePixelPositions(sizes, startPixel, totalPixels) {
        // Calculate total EMU size
        const totalEmu = sizes.reduce((sum, s) => sum + s, 0);
        if (totalEmu === 0)
            return sizes.map(() => startPixel);
        // Scale factor to fit into total pixels
        const scale = totalPixels / (this.unitConverter.emuToPixels(totalEmu) * this.scaleX);
        // Calculate cumulative positions
        const positions = [startPixel];
        let current = startPixel;
        for (const size of sizes) {
            const pixelSize = this.unitConverter.emuToPixels(size) * this.scaleX * scale;
            current += pixelSize;
            positions.push(current);
        }
        return positions;
    }
    /**
     * Calculates bounds for a cell accounting for spans.
     */
    calculateCellBounds(columnPixels, rowPixels, colIndex, rowIndex, colSpan, rowSpan) {
        const x = columnPixels[colIndex] ?? 0;
        const y = rowPixels[rowIndex] ?? 0;
        const endColIndex = Math.min(colIndex + colSpan, columnPixels.length - 1);
        const endRowIndex = Math.min(rowIndex + rowSpan, rowPixels.length - 1);
        const width = (columnPixels[endColIndex] ?? x) - x;
        const height = (rowPixels[endRowIndex] ?? y) - y;
        return { x, y, width, height };
    }
    /**
     * Merges a cell's explicit tcPr formatting over the effective table-style
     * formatting. Explicit values win; explicit 'none' suppresses the style.
     * @param cell The parsed cell
     * @param styled The effective table-style formatting, if any
     * @returns The fill color and per-edge borders to render
     */
    mergeCellWithStyle(cell, styled) {
        const fill = cell.backgroundColor === 'none' ? undefined : (cell.backgroundColor ?? styled?.fill);
        return {
            fill,
            borders: {
                left: cell.borders.left ?? styled?.borders.left,
                right: cell.borders.right ?? styled?.borders.right,
                top: cell.borders.top ?? styled?.borders.top,
                bottom: cell.borders.bottom ?? styled?.borders.bottom,
                tlToBr: cell.borders.tlToBr ?? styled?.borders.tlToBr,
                blToTr: cell.borders.blToTr ?? styled?.borders.blToTr,
            },
        };
    }
    /**
     * Renders a single cell. When rtl is set the cell rect has been mirrored
     * horizontally, so sided borders swap edges and the diagonals swap
     * directions.
     */
    renderCell(ctx, cell, bounds, styled, rtl = false) {
        const { fill, borders } = this.mergeCellWithStyle(cell, styled);
        // Render cell background
        if (fill) {
            this.paintFill(ctx, fill, bounds);
        }
        // Render cell borders
        const drawnBorders = rtl
            ? {
                left: borders.right,
                right: borders.left,
                top: borders.top,
                bottom: borders.bottom,
                tlToBr: borders.blToTr,
                blToTr: borders.tlToBr,
            }
            : borders;
        this.renderCellBorders(ctx, drawnBorders, bounds);
        // Render cell text
        if (cell.textBody && cell.textBody.paragraphs.length > 0) {
            // Apply table-style text defaults where runs specify nothing
            if (styled) {
                this.applyStyleTextDefaults(cell.textBody, styled);
            }
            // Calculate text bounds with margins. marL/marR are logical cell
            // properties: in a mirrored RTL cell the logical left margin applies
            // to the visual right edge.
            const marginLeft = this.unitConverter.emuToPixels(rtl ? cell.margins.right : cell.margins.left) * this.scaleX;
            const marginRight = this.unitConverter.emuToPixels(rtl ? cell.margins.left : cell.margins.right) * this.scaleX;
            const marginTop = this.unitConverter.emuToPixels(cell.margins.top) * this.scaleY;
            const marginBottom = this.unitConverter.emuToPixels(cell.margins.bottom) * this.scaleY;
            const textBounds = {
                x: bounds.x + marginLeft,
                y: bounds.y + marginTop,
                width: Math.max(0, bounds.width - marginLeft - marginRight),
                height: Math.max(0, bounds.height - marginTop - marginBottom),
            };
            // Apply vertical alignment if specified
            if (cell.anchor && cell.textBody.bodyProperties) {
                cell.textBody.bodyProperties.anchor = cell.anchor;
            }
            this.textRenderer.renderText(ctx, cell.textBody, textBounds);
        }
    }
    /**
     * Fills table-style text defaults (bold/italic/color/font) into runs that
     * do not specify those properties themselves, so explicit run formatting
     * keeps winning over the style.
     */
    applyStyleTextDefaults(textBody, styled) {
        if (styled.bold === undefined &&
            styled.italic === undefined &&
            styled.color === undefined &&
            styled.fontFamily === undefined) {
            return;
        }
        for (const paragraph of textBody.paragraphs) {
            for (const run of paragraph.runs) {
                const props = (run.properties ??= {});
                if (props.bold === undefined && styled.bold !== undefined) {
                    props.bold = styled.bold;
                }
                if (props.italic === undefined && styled.italic !== undefined) {
                    props.italic = styled.italic;
                }
                if (props.color === undefined && props.schemeColor === undefined && styled.color) {
                    props.color = { ...styled.color };
                }
                if (props.fontFamily === undefined && styled.fontFamily !== undefined) {
                    props.fontFamily = styled.fontFamily;
                }
            }
        }
    }
    /**
     * Paints a cell or table background fill (solid color or linear gradient)
     * into the given rect. Linear gradients are anchored in the rect itself:
     * the gradient line runs through the rect center at the fill's angle.
     * @param ctx Canvas 2D context
     * @param fill The fill to paint
     * @param rect The rect to fill (in pixels)
     */
    paintFill(ctx, fill, rect) {
        if (isGradientCellFill(fill)) {
            const style = this.createLinearGradient(ctx, fill, rect);
            if (!style)
                return;
            ctx.fillStyle = style;
        }
        else {
            ctx.fillStyle = this.colorResolver.rgbaToCss(fill);
        }
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    /**
     * Creates a canvas linear gradient for a cell fill within the given rect
     * (degrades to the first stop's solid color when under two stops).
     */
    createLinearGradient(ctx, fill, rect) {
        const firstStop = fill.stops[0];
        if (!firstStop)
            return undefined;
        if (fill.stops.length < 2) {
            return this.colorResolver.rgbaToCss(firstStop.color);
        }
        const centerX = rect.x + rect.width / 2;
        const centerY = rect.y + rect.height / 2;
        const halfDiagonal = Math.sqrt(rect.width * rect.width + rect.height * rect.height) / 2;
        const { x0, y0, x1, y1 } = (0, FillRenderer_js_1.computeLinearGradientPoints)(centerX, centerY, halfDiagonal, fill.angle);
        const gradient = ctx.createLinearGradient(x0, y0, x1, y1);
        for (const stop of fill.stops) {
            const position = Math.min(Math.max(stop.position, 0), 1);
            gradient.addColorStop(position, this.colorResolver.rgbaToCss(stop.color));
        }
        return gradient;
    }
    /**
     * Renders cell borders (edges and diagonals).
     */
    renderCellBorders(ctx, borders, bounds) {
        ctx.save();
        const right = bounds.x + bounds.width;
        const bottom = bounds.y + bounds.height;
        const diagScale = (this.scaleX + this.scaleY) / 2;
        // Diagonals first so edge borders paint over their endpoints
        this.drawBorderEdge(ctx, borders.tlToBr, bounds.x, bounds.y, right, bottom, diagScale);
        this.drawBorderEdge(ctx, borders.blToTr, bounds.x, bottom, right, bounds.y, diagScale);
        this.drawBorderEdge(ctx, borders.left, bounds.x, bounds.y, bounds.x, bottom, this.scaleX);
        this.drawBorderEdge(ctx, borders.right, right, bounds.y, right, bottom, this.scaleX);
        this.drawBorderEdge(ctx, borders.top, bounds.x, bounds.y, right, bounds.y, this.scaleY);
        this.drawBorderEdge(ctx, borders.bottom, bounds.x, bottom, right, bottom, this.scaleY);
        ctx.restore();
    }
    /**
     * Draws a single border edge as a line segment.
     */
    drawBorderEdge(ctx, border, x1, y1, x2, y2, scale) {
        if (!border || border === 'none')
            return;
        const widthPixels = Math.max(this.unitConverter.emuToPixels(border.width) * scale, 0.5);
        ctx.strokeStyle = this.colorResolver.rgbaToCss(border.color);
        ctx.lineWidth = widthPixels;
        // Dash pattern scales with stroke width (as in StrokeRenderer)
        const pattern = border.dash ? (PRESET_DASH_PATTERNS[border.dash] ?? []) : [];
        ctx.setLineDash(pattern.length > 0 ? pattern.map((v) => v * widthPixels) : []);
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
    }
}
exports.TableRenderer = TableRenderer;
/**
 * Reads the a:tableStyleId child text (the style GUID) from a:tblPr.
 * Handles both standard-parsed nodes (text stored as the value itself) and
 * ordered-parse conversions (text stored under '#text').
 */
function parseTableStyleId(tblPr) {
    const raw = tblPr['a:tableStyleId'];
    let text;
    if (typeof raw === 'string') {
        text = raw;
    }
    else if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
        text = raw['#text'];
    }
    if (typeof text !== 'string')
        return undefined;
    const trimmed = text.trim();
    return trimmed.length > 0 ? trimmed : undefined;
}
/**
 * Mirrors a rect horizontally within the given table bounds. RTL tables
 * (a:tblPr rtl="1") lay their columns out right-to-left, so each logical
 * cell rect flips around the table's vertical center line.
 * @param rect The cell rect in logical (LTR) position
 * @param tableBounds The full table bounds
 * @returns The mirrored rect
 */
function mirrorRectHorizontally(rect, tableBounds) {
    return {
        x: 2 * tableBounds.x + tableBounds.width - rect.x - rect.width,
        y: rect.y,
        width: rect.width,
        height: rect.height,
    };
}
/**
 * Converts a table-style border to a renderable cell border ('none' for an
 * explicit removal, undefined when unspecified).
 */
function toParsedCellBorder(border) {
    if (!border)
        return undefined;
    if (border.none)
        return 'none';
    return {
        width: border.width ?? DEFAULT_BORDER_WIDTH_EMU,
        color: border.color ?? { ...geometry_js_1.Colors.black },
        dash: border.dash,
    };
}
/**
 * Creates a TableRenderer instance.
 */
function createTableRenderer(theme, scaleX, scaleY, logger) {
    return new TableRenderer({ theme, scaleX, scaleY, logger });
}
