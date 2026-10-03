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
export const ANGLE_UNITS_PER_DEGREE = 60000;
/** OOXML angle units in a full circle (360° × 60000). */
export const ANGLE_UNITS_PER_CIRCLE = 21600000;
/**
 * Converts an OOXML angle (60000ths of a degree) to radians.
 */
export function ooxmlAngleToRadians(angle) {
    return (angle / ANGLE_UNITS_PER_DEGREE) * (Math.PI / 180);
}
/**
 * Converts radians to an OOXML angle (60000ths of a degree).
 */
export function radiansToOoxmlAngle(radians) {
    return (radians * 180 * ANGLE_UNITS_PER_DEGREE) / Math.PI;
}
/**
 * Creates a guide evaluation context for the given coordinate space.
 * @param w Width of the guide coordinate space
 * @param h Height of the guide coordinate space
 * @param adjustments Optional initial named values (adjust values)
 */
export function createGuideContext(w, h, adjustments) {
    const guides = new Map();
    if (adjustments) {
        for (const [name, value] of adjustments) {
            guides.set(name, value);
        }
    }
    return { w, h, guides };
}
/** Matches a numeric literal token (integer or decimal, optionally signed). */
const NUMERIC_LITERAL = /^[-+]?\d+(?:\.\d+)?$/;
/** Matches derived built-in guides: wd2..wdN, hd2..hdN, ssd2..ssdN. */
const DERIVED_GUIDE = /^(wd|hd|ssd)(\d+)$/;
/** Matches angular built-in guides: cdN and McdN (e.g. cd4, 3cd8). */
const ANGULAR_GUIDE = /^(\d*)cd(\d+)$/;
/**
 * Resolves a formula argument or coordinate token to its numeric value.
 * Tokens may be numeric literals, adjust/guide names, or built-in guides
 * (w, h, l, t, r, b, hc, vc, ss, ls, wd2.., hd2.., ssd2.., cd2, 3cd4, ...).
 * @throws Error if the token is not a number, guide, or built-in
 */
export function resolveGuideValue(token, ctx) {
    if (typeof token === 'number')
        return token;
    // User-defined guides and adjust values take precedence
    const named = ctx.guides.get(token);
    if (named !== undefined)
        return named;
    // Fixed built-in guides
    switch (token) {
        case 'w':
            return ctx.w;
        case 'h':
            return ctx.h;
        case 'l':
        case 't':
            return 0;
        case 'r':
            return ctx.w;
        case 'b':
            return ctx.h;
        case 'hc':
            return ctx.w / 2;
        case 'vc':
            return ctx.h / 2;
        case 'ss':
            return Math.min(ctx.w, ctx.h);
        case 'ls':
            return Math.max(ctx.w, ctx.h);
    }
    // Derived built-ins: wdN = w/N, hdN = h/N, ssdN = ss/N
    const derived = DERIVED_GUIDE.exec(token);
    if (derived?.[1] !== undefined && derived[2] !== undefined) {
        const divisor = parseInt(derived[2], 10);
        if (divisor > 0) {
            const base = derived[1] === 'wd' ? ctx.w : derived[1] === 'hd' ? ctx.h : Math.min(ctx.w, ctx.h);
            return base / divisor;
        }
    }
    // Angular built-ins: cdN = 21600000/N, McdN = M*21600000/N (e.g. 3cd4 = 270°)
    const angular = ANGULAR_GUIDE.exec(token);
    if (angular?.[2] !== undefined) {
        const multiplier = angular[1] ? parseInt(angular[1], 10) : 1;
        const divisor = parseInt(angular[2], 10);
        if (divisor > 0) {
            return (multiplier * ANGLE_UNITS_PER_CIRCLE) / divisor;
        }
    }
    // Numeric literal
    if (NUMERIC_LITERAL.test(token)) {
        return parseFloat(token);
    }
    throw new Error(`Unknown guide reference: ${token}`);
}
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
export function evaluateFormula(fmla, ctx) {
    const tokens = fmla.trim().split(/\s+/);
    const op = tokens[0];
    if (!op) {
        throw new Error(`Empty guide formula`);
    }
    const arg = (index) => {
        const token = tokens[index];
        if (token === undefined) {
            throw new Error(`Guide formula missing argument ${index}: ${fmla}`);
        }
        return resolveGuideValue(token, ctx);
    };
    switch (op) {
        case '*/': {
            const z = arg(3);
            return z === 0 ? 0 : (arg(1) * arg(2)) / z;
        }
        case '+-':
            return arg(1) + arg(2) - arg(3);
        case '+/': {
            const z = arg(3);
            return z === 0 ? 0 : (arg(1) + arg(2)) / z;
        }
        case '?:':
            return arg(1) > 0 ? arg(2) : arg(3);
        case 'abs':
            return Math.abs(arg(1));
        case 'at2':
            return radiansToOoxmlAngle(Math.atan2(arg(2), arg(1)));
        case 'cat2':
            return arg(1) * Math.cos(Math.atan2(arg(3), arg(2)));
        case 'sat2':
            return arg(1) * Math.sin(Math.atan2(arg(3), arg(2)));
        case 'cos':
            return arg(1) * Math.cos(ooxmlAngleToRadians(arg(2)));
        case 'sin':
            return arg(1) * Math.sin(ooxmlAngleToRadians(arg(2)));
        case 'tan':
            return arg(1) * Math.tan(ooxmlAngleToRadians(arg(2)));
        case 'max':
            return Math.max(arg(1), arg(2));
        case 'min':
            return Math.min(arg(1), arg(2));
        case 'mod': {
            const x = arg(1);
            const y = arg(2);
            const z = arg(3);
            return Math.sqrt(x * x + y * y + z * z);
        }
        case 'pin': {
            const min = arg(1);
            const value = arg(2);
            const max = arg(3);
            // Per ECMA-376: if y < x then x, else if y > z then z, else y
            return value < min ? min : value > max ? max : value;
        }
        case 'sqrt': {
            const x = arg(1);
            return x <= 0 ? 0 : Math.sqrt(x);
        }
        case 'val':
            return arg(1);
        default:
            throw new Error(`Unknown guide formula operation: ${op}`);
    }
}
/**
 * Evaluates an ordered list of guide definitions into the context.
 * Each guide may reference adjust values, built-ins, and previously
 * evaluated guides (document order matters, per ECMA-376).
 *
 * @param entries Ordered [name, formula] pairs from a gdLst
 * @param ctx Context to evaluate into (mutated)
 */
export function evaluateGuideList(entries, ctx) {
    for (const [name, fmla] of entries) {
        ctx.guides.set(name, evaluateFormula(fmla, ctx));
    }
}
//# sourceMappingURL=GuideEvaluator.js.map