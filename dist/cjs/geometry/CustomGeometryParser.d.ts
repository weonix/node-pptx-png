/**
 * Parses custom geometry (a:custGeom) XML into a GeometryDefinition that the
 * shared geometry engine can evaluate — the same representation used for
 * preset shapes, since custom geometry uses the identical drawing language
 * (ECMA-376 §20.1.9.8).
 *
 * Path segments are read in true document order via
 * getChildrenInDocumentOrder, which is order-preserving for nodes produced
 * by the ordered slide parse (interleaved moveTo/lnTo/arcTo sequences and
 * multiple subpaths are handled correctly).
 */
import type { PptxXmlNode } from '../core/PptxParser.js';
import type { GeometryDefinition } from './GeometryEngine.js';
/**
 * Parses an a:custGeom node into a GeometryDefinition for the geometry
 * engine: adjust value defaults (a:avLst), guide formulas (a:gdLst, in
 * document order), the text rectangle (a:rect), and all paths (a:pathLst)
 * with per-path coordinate spaces and fill/stroke flags.
 *
 * @param custGeom The a:custGeom XML node
 * @returns The geometry definition, or undefined if no drawable path exists
 */
export declare function parseCustomGeometryDefinition(custGeom: PptxXmlNode): GeometryDefinition | undefined;
