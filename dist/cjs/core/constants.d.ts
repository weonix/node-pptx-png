/**
 * Shared constants for PPTX element types and other common values.
 */
/**
 * Element types that can appear in a shape tree (p:spTree) or group shape (p:grpSp).
 * These represent the different kinds of visual elements in a PowerPoint slide.
 * Order is significant for iteration but not for z-order (document order determines z-order).
 */
export declare const SHAPE_ELEMENT_TYPES: readonly ["p:sp", "p:cxnSp", "p:pic", "p:grpSp", "p:graphicFrame", "mc:AlternateContent"];
/**
 * Type representing a valid shape element type.
 */
export type ShapeElementType = (typeof SHAPE_ELEMENT_TYPES)[number];
