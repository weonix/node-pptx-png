/**
 * SmartArt diagram rendering via the pre-rendered drawing part.
 *
 * PowerPoint persists a pre-baked rendering of every SmartArt graphic as a
 * "diagram drawing" part (MS-ODRAWXML, dsp: namespace): the slide's
 * graphicFrame carries <dgm:relIds r:dm=.../> pointing at the diagram data
 * part (diagrams/dataN.xml), and the drawing part (diagrams/drawingN.xml) is
 * reachable either through the data part's own relationships or through the
 * <dsp:dataModelExt relId=.../> extension inside the data model, whose relId
 * resolves against the referencing (slide) part. The drawing part's
 * <dsp:spTree> holds plain DrawingML shapes (dsp:sp with spPr/style/txBody,
 * structurally identical to p:sp), so no diagram layout engine is needed.
 * dsp:grpSp groups nest exactly like p:grpSp (grpSpPr xfrm with
 * off/ext/chOff/chExt) and are flattened here; a dsp:txXfrm on a shape lays
 * its text out in a rectangle independent of the shape geometry.
 */
import type { CanvasRenderingContext2D } from 'skia-canvas';
import type { ResolvedTheme } from '../types/index.js';
import type { Rect } from '../types/geometry.js';
import type { PptxParser, PptxXmlNode } from '../core/PptxParser.js';
import type { ILogger } from '../utils/Logger.js';
import type { WarningCollector } from '../utils/WarningCollector.js';
/**
 * The a:graphicData uri identifying a SmartArt diagram graphic frame.
 */
export declare const DIAGRAM_GRAPHIC_DATA_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
/**
 * Configuration for DiagramRenderer.
 */
export interface DiagramRendererConfig {
    /** Resolved theme for color resolution */
    theme: ResolvedTheme;
    /** PPTX parser for accessing diagram parts */
    parser: PptxParser;
    /** Path of the slide referencing the diagram (e.g., ppt/slides/slide1.xml) */
    slidePath: string;
    /** Slide node (p:sld) for color map override resolution */
    slideNode?: PptxXmlNode;
    /** Slide layout node for color map resolution */
    layoutNode?: PptxXmlNode;
    /** Slide master node for color map resolution */
    masterNode?: PptxXmlNode;
    /** Logger instance */
    logger?: ILogger;
    /** Structured warning collector for diagram fidelity events */
    warnings?: WarningCollector;
}
/**
 * Renders SmartArt diagrams by drawing the shapes of the pre-rendered
 * drawing part (dsp:spTree) through the standard ShapeRenderer.
 */
export declare class DiagramRenderer {
    private readonly logger;
    private readonly theme;
    private readonly parser;
    private readonly slidePath;
    private readonly slideNode?;
    private readonly layoutNode?;
    private readonly masterNode?;
    private readonly relationshipParser;
    private readonly unitConverter;
    private readonly transformCalculator;
    private readonly groupShapeRenderer;
    private readonly warnings;
    constructor(config: DiagramRendererConfig);
    /**
     * Renders a SmartArt diagram graphic frame.
     *
     * Diagram shape coordinates live in the drawing part's own EMU space,
     * anchored at the graphicFrame origin and 1:1 with the frame's EMU extents
     * (PowerPoint regenerates the drawing part whenever the frame is resized).
     * Rendering therefore translates to the frame's pixel origin and scales
     * shapes by pixels-per-EMU derived from the frame bounds, which also
     * absorbs any group scaling already applied to the bounds.
     *
     * Missing drawing parts (files saved by producers that never pre-render)
     * degrade gracefully: a warning is logged and nothing is drawn.
     *
     * @param ctx Canvas 2D context
     * @param graphicData The a:graphicData node containing dgm:relIds
     * @param bounds The frame bounds in canvas pixels
     * @param frameExtEmu The frame's own EMU extents (p:xfrm/a:ext cx/cy)
     */
    renderDiagram(ctx: CanvasRenderingContext2D, graphicData: PptxXmlNode, bounds: Rect, frameExtEmu: {
        cx: number;
        cy: number;
    }): Promise<void>;
    /**
     * Renders one collected diagram shape: composes the enclosing group's
     * child-space mapping into the shape transform, and honors a dsp:txXfrm
     * text rectangle by rendering the geometry and the text as separate passes
     * when it differs from the shape's own xfrm.
     */
    private renderDiagramShape;
    /**
     * Parses a dsp:sp's text transform (dsp:txXfrm, MS-ODRAWXML). Returns the
     * drawing-space EMU rectangle to lay text out in (composed through the
     * enclosing group mapping), or undefined when the shape carries none, when
     * it is degenerate, or when it matches the shape xfrm (where the normal
     * in-shape text path is already correct).
     */
    private parseTextXfrm;
    /**
     * Renders a diagram shape's text body inside an EMU rectangle independent
     * of the shape geometry (dsp:txXfrm), applying the rectangle's own
     * rotation/flip. The default text color follows ShapeRenderer's
     * precedence: the style fontRef color, then the fill-contrast heuristic.
     */
    private renderTextInRect;
    /**
     * Resolves the diagram drawing part path for a diagram data part.
     *
     * Tries, in order:
     * 1. A diagramDrawing relationship in the data part's own rels
     *    (spec-conformant producers).
     * 2. The dsp:dataModelExt relId extension inside the data model, resolved
     *    against the data part rels and then the slide rels (PowerPoint stores
     *    the diagramDrawing relationship on the slide).
     *
     * @param dataPath Path of the diagram data part (e.g., ppt/diagrams/data1.xml)
     * @returns The drawing part path, or undefined when no drawing part exists
     */
    private resolveDrawingPart;
    /**
     * Checks that a relationship is a diagram drawing relationship, so a
     * dataModelExt relId is never resolved through an unrelated relationship.
     */
    private isDiagramDrawingRel;
    /**
     * Reads the relId attribute of the dsp:dataModelExt extension inside a
     * diagram data part (dgm:dataModel > dgm:extLst > a:ext > dsp:dataModelExt).
     * Returns undefined when the data model carries no such extension.
     */
    private readDataModelExtRelId;
    /**
     * Reads the drawing part and returns its shapes in document order, with
     * dsp:grpSp subtrees flattened into shape entries carrying the composed
     * group child-space mapping. Flattening preserves document order, so
     * z-order is unchanged.
     */
    private readDrawingShapes;
    /**
     * Flattens a dsp:grpSp subtree into shape entries, composing each nesting
     * level's child-space mapping (chOff/chExt onto off/ext, plus the group
     * xfrm's rotation/flip) exactly like GroupShapeRenderer does for p:grpSp
     * trees.
     */
    private collectGroupShapes;
    /**
     * Composes a nested group's transform with its parent's, mirroring
     * GroupShapeRenderer.renderGroupShape: the inner group's bounds are mapped
     * through the parent's child-space transform, and the effective child
     * scale factors are recomputed from the transformed bounds so the parent's
     * scaling is inherited by the inner group's children.
     */
    private composeGroupTransform;
    /**
     * Navigates the ordered drawing XML (dsp:drawing > dsp:spTree) to the
     * spTree children array, preserving document order.
     */
    private navigateToSpTree;
    /**
     * Rewraps a dsp:sp node as a p:sp-shaped node so ShapeRenderer can render
     * it unchanged. Child nodes are shared by reference, so ordered-source
     * metadata on spPr/txBody content (path segment order, run interleaving)
     * is preserved.
     */
    private toPShapeNode;
}
/**
 * Creates a DiagramRenderer instance.
 */
export declare function createDiagramRenderer(config: DiagramRendererConfig): DiagramRenderer;
