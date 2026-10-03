/**
 * Evaluator for the OOXML shape guide formula language (ECMA-376 §20.1.9.11).
 *
 * Guide formulas compute named values from shape dimensions, adjust values,
 * and previously defined guides. They power both preset geometry definitions
 * (presetShapeDefinitions.xml) and custom geometry (a:custGeom/a:gdLst).
 *
 * Angles are expressed in 60000ths of a degree throughout (so 5400000 = 90°).
 */
/** OOXML angle units per degree (60000ths of a degree). */
export declare const ANGLE_UNITS_PER_DEGREE = 60000;
/** OOXML angle units in a full circle (360° × 60000). */
export declare const ANGLE_UNITS_PER_CIRCLE = 21600000;
/**
 * Converts an OOXML angle (60000ths of a degree) to radians.
 */
export declare function ooxmlAngleToRadians(angle: number): number;
/**
 * Converts radians to an OOXML angle (60000ths of a degree).
 */
export declare function radiansToOoxmlAngle(radians: number): number;
/**
 * Evaluation context for guide formulas.
 * Holds the shape's guide coordinate space dimensions and all named guide
 * values resolved so far (adjust values and evaluated gdLst entries).
 */
export interface GuideContext {
    /** Width of the guide coordinate space. */
    readonly w: number;
    /** Height of the guide coordinate space. */
    readonly h: number;
    /** Named guide values (adjust values and evaluated guides). */
    readonly guides: Map<string, number>;
}
/**
 * Creates a guide evaluation context for the given coordinate space.
 * @param w Width of the guide coordinate space
 * @param h Height of the guide coordinate space
 * @param adjustments Optional initial named values (adjust values)
 */
export declare function createGuideContext(w: number, h: number, adjustments?: Iterable<readonly [string, number]>): GuideContext;
/**
 * Resolves a formula argument or coordinate token to its numeric value.
 * Tokens may be numeric literals, adjust/guide names, or built-in guides
 * (w, h, l, t, r, b, hc, vc, ss, ls, wd2.., hd2.., ssd2.., cd2, 3cd4, ...).
 * @throws Error if the token is not a number, guide, or built-in
 */
export declare function resolveGuideValue(token: string | number, ctx: GuideContext): number;
/**
 * Evaluates a single guide formula (the fmla attribute of an a:gd element).
 *
 * Supported operations (x, y, z are literals or guide references):
 * - `*\/ x y z` → (x * y) / z
 * - `+- x y z` → x + y - z
 * - `+/ x y z` → (x + y) / z
 * - `?: x y z` → x > 0 ? y : z
 * - `abs x`, `max x y`, `min x y`, `sqrt x`, `val x`
 * - `mod x y z` → sqrt(x² + y² + z²)
 * - `pin x y z` → y clamped to [x, z]
 * - `sin x y`, `cos x y`, `tan x y` → x * trig(y), y an OOXML angle
 * - `at2 x y` → atan2(y, x) as an OOXML angle
 * - `cat2 x y z` → x * cos(atan2(z, y)); `sat2 x y z` → x * sin(atan2(z, y))
 *
 * Division by zero yields 0 (matching renderer-friendly behavior for
 * degenerate shapes; ECMA-376 leaves it undefined).
 *
 * @throws Error on malformed formulas or unknown operations/references
 */
export declare function evaluateFormula(fmla: string, ctx: GuideContext): number;
/**
 * Evaluates an ordered list of guide definitions into the context.
 * Each guide may reference adjust values, built-ins, and previously
 * evaluated guides (document order matters, per ECMA-376).
 *
 * @param entries Ordered [name, formula] pairs from a gdLst
 * @param ctx Context to evaluate into (mutated)
 */
export declare function evaluateGuideList(entries: ReadonlyArray<readonly [string, string]>, ctx: GuideContext): void;
//# sourceMappingURL=GuideEvaluator.d.ts.map