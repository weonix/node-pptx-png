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
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { Rect, Rgba } from '../types/geometry.js';
import type { ResolvedTheme, ColorMap } from '../types/theme.js';
import type { TextBody } from '../types/elements.js';
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { TableStyle, TableStyleCollection, TableStyleGradientFill } from '../parsers/TableStyleParser.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * Parsed table structure.
 */
export interface ParsedTable {
    /** Column widths in EMU */
    columnWidths: number[];
    /** Row heights in EMU */
    rowHeights: number[];
    /** Table cells organized by row */
    rows: ParsedTableRow[];
    /** Table properties */
    properties: TableProperties;
}
/**
 * Table-level properties.
 */
export interface TableProperties {
    /** First row has special formatting */
    firstRow?: boolean;
    /** First column has special formatting */
    firstCol?: boolean;
    /** Last row has special formatting */
    lastRow?: boolean;
    /** Last column has special formatting */
    lastCol?: boolean;
    /** Banded rows (alternating row colors) */
    bandRow?: boolean;
    /** Banded columns (alternating column colors) */
    bandCol?: boolean;
    /** Right-to-left table (columns mirror horizontally) */
    rtl?: boolean;
    /** Table style GUID from a:tableStyleId */
    tableStyleId?: string;
}
/**
 * A concrete cell fill: a solid color or a linear gradient (from an
 * explicit tcPr fill or a table-style part fill).
 */
export type CellFill = Rgba | TableStyleGradientFill;
/**
 * Narrows a cell fill to its gradient variant.
 * @param fill The cell fill
 * @returns True when the fill is a linear gradient
 */
export declare function isGradientCellFill(fill: CellFill): fill is TableStyleGradientFill;
/**
 * Parsed table row.
 */
export interface ParsedTableRow {
    /** Row height in EMU */
    height: number;
    /** Cells in this row */
    cells: ParsedTableCell[];
}
/**
 * Parsed table cell.
 */
export interface ParsedTableCell {
    /** Cell text body (parsed) */
    textBody?: TextBody;
    /** Cell background fill ('none' marks an explicit a:noFill) */
    backgroundColor?: CellFill | 'none';
    /** Cell margins in EMU */
    margins: CellMargins;
    /** Cell borders */
    borders: CellBorders;
    /** Column span (gridSpan attribute) */
    colSpan: number;
    /** Row span (rowSpan attribute) */
    rowSpan: number;
    /** Whether this cell is merged horizontally (hMerge) */
    hMerge: boolean;
    /** Whether this cell is merged vertically (vMerge) */
    vMerge: boolean;
    /** Vertical text alignment */
    anchor?: 'top' | 'middle' | 'bottom';
}
/**
 * Cell margins in EMU.
 */
export interface CellMargins {
    left: number;
    right: number;
    top: number;
    bottom: number;
}
/**
 * A parsed cell border: a drawable border, or 'none' when the border is
 * explicitly removed (a:noFill), which must suppress table-style borders.
 */
export type ParsedCellBorder = CellBorder | 'none';
/**
 * Cell border definitions. Undefined edges are unspecified (a table style
 * may supply them); 'none' edges are explicitly removed.
 */
export interface CellBorders {
    left?: ParsedCellBorder;
    right?: ParsedCellBorder;
    top?: ParsedCellBorder;
    bottom?: ParsedCellBorder;
    /** Diagonal border from top-left to bottom-right (a:lnTlToBr) */
    tlToBr?: ParsedCellBorder;
    /** Diagonal border from bottom-left to top-right (a:lnBlToTr) */
    blToTr?: ParsedCellBorder;
}
/**
 * Single border definition.
 */
export interface CellBorder {
    /** Border width in EMU */
    width: number;
    /** Border color */
    color: Rgba;
    /** Preset dash value (solid, dash, dot, ...) */
    dash?: string;
}
/**
 * The effective table-style formatting for one cell, computed by layering
 * the style's conditional formatting parts. Explicit tcPr formatting is
 * applied on top of this at render time.
 */
export interface EffectiveCellStyle {
    /** Cell fill (undefined when the style yields no fill) */
    fill?: CellFill;
    /** Cell borders from the style */
    borders: CellBorders;
    /** Default text bold flag */
    bold?: boolean;
    /** Default text italic flag */
    italic?: boolean;
    /** Default text color */
    color?: Rgba;
    /** Default font family */
    fontFamily?: string;
}
/**
 * Configuration for TableRenderer.
 */
export interface TableRendererConfig {
    /** Resolved theme for color/font resolution */
    theme: ResolvedTheme;
    /** Color map for scheme color remapping (bg1/tx1/...) */
    colorMap?: ColorMap;
    /** Horizontal scale factor */
    scaleX: number;
    /** Vertical scale factor */
    scaleY: number;
    /** Parsed table styles (ppt/tableStyles.xml) for a:tableStyleId lookup */
    tableStyles?: TableStyleCollection;
    /** 1-based slide number for a:fld type="slidenum" resolution */
    slideNumber?: number;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector, forwarded to the text pipeline */
    warnings?: WarningCollector;
}
/**
 * Renders tables to canvas.
 */
export declare class TableRenderer {
    private readonly logger;
    private readonly theme;
    private readonly scaleX;
    private readonly scaleY;
    private readonly colorResolver;
    private readonly unitConverter;
    private readonly textParser;
    private readonly textRenderer;
    private readonly tableStyles?;
    constructor(config: TableRendererConfig);
    /**
     * Renders a table to the canvas.
     *
     * @param ctx Canvas 2D context
     * @param tableNode The a:tbl XML node
     * @param bounds The bounds to render within (in pixels)
     */
    renderTable(ctx: CanvasRenderingContext2D, tableNode: PptxXmlNode, bounds: Rect): void;
    /**
     * Looks up a table style by its GUID in the configured style collection.
     * PowerPoint frequently references built-in styles (e.g. Medium Style 2
     * Accent 1, {5C22544A-7EE6-4342-B048-85BDC9FD1C3A}) that tableStyles.xml
     * does not define; those render unstyled with a debug log.
     * @param styleId The style GUID from a:tableStyleId
     * @returns The parsed style, or undefined when it is not defined
     */
    lookupTableStyle(styleId: string): TableStyle | undefined;
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
    computeEffectiveCellStyle(style: TableStyle, props: TableProperties, rowIndex: number, colIndex: number, rowSpan: number, colSpan: number, rowCount: number, colCount: number): EffectiveCellStyle;
    /**
     * Parses a table XML node into a structured format.
     */
    parseTable(tableNode: PptxXmlNode): ParsedTable | undefined;
    /**
     * Parses table properties (a:tblPr flags and the a:tableStyleId child).
     */
    private parseTableProperties;
    /**
     * Parses column widths from tblGrid.
     */
    private parseColumnWidths;
    /**
     * Parses a table row.
     */
    private parseRow;
    /**
     * Parses a table cell.
     */
    private parseCell;
    /**
     * Parses the cell background fill. Returns 'none' for an explicit
     * a:noFill, which suppresses any table-style fill. Linear a:gradFill
     * fills are preserved as gradients; radial/path gradients degrade to
     * their lowest-position stop.
     */
    private parseCellBackground;
    /**
     * Parses cell margins.
     */
    private parseCellMargins;
    /**
     * Parses cell borders.
     */
    private parseCellBorders;
    /**
     * Parses a single border line. Returns 'none' for an explicit a:noFill,
     * which suppresses any table-style border.
     */
    private parseBorder;
    /**
     * Calculates pixel positions from EMU sizes.
     */
    private calculatePixelPositions;
    /**
     * Calculates bounds for a cell accounting for spans.
     */
    private calculateCellBounds;
    /**
     * Merges a cell's explicit tcPr formatting over the effective table-style
     * formatting. Explicit values win; explicit 'none' suppresses the style.
     * @param cell The parsed cell
     * @param styled The effective table-style formatting, if any
     * @returns The fill color and per-edge borders to render
     */
    mergeCellWithStyle(cell: ParsedTableCell, styled?: EffectiveCellStyle): {
        fill?: CellFill;
        borders: CellBorders;
    };
    /**
     * Renders a single cell. When rtl is set the cell rect has been mirrored
     * horizontally, so sided borders swap edges and the diagonals swap
     * directions.
     */
    private renderCell;
    /**
     * Fills table-style text defaults (bold/italic/color/font) into runs that
     * do not specify those properties themselves, so explicit run formatting
     * keeps winning over the style.
     */
    private applyStyleTextDefaults;
    /**
     * Paints a cell or table background fill (solid color or linear gradient)
     * into the given rect. Linear gradients are anchored in the rect itself:
     * the gradient line runs through the rect center at the fill's angle.
     * @param ctx Canvas 2D context
     * @param fill The fill to paint
     * @param rect The rect to fill (in pixels)
     */
    private paintFill;
    /**
     * Creates a canvas linear gradient for a cell fill within the given rect
     * (degrades to the first stop's solid color when under two stops).
     */
    private createLinearGradient;
    /**
     * Renders cell borders (edges and diagonals).
     */
    private renderCellBorders;
    /**
     * Draws a single border edge as a line segment.
     */
    private drawBorderEdge;
}
/**
 * Mirrors a rect horizontally within the given table bounds. RTL tables
 * (a:tblPr rtl="1") lay their columns out right-to-left, so each logical
 * cell rect flips around the table's vertical center line.
 * @param rect The cell rect in logical (LTR) position
 * @param tableBounds The full table bounds
 * @returns The mirrored rect
 */
export declare function mirrorRectHorizontally(rect: Rect, tableBounds: Rect): Rect;
/**
 * Creates a TableRenderer instance.
 */
export declare function createTableRenderer(theme: ResolvedTheme, scaleX: number, scaleY: number, logger?: ILogger): TableRenderer;
