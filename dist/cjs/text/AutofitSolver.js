"use strict";
/**
 * Render-time normAutofit (shrink-to-fit) solver.
 *
 * ECMA-376 §21.1.2.1.3 defines `<a:normAutofit/>` ("Normal AutoFit") as
 * "shrink text on overflow": the application scales the text down until it
 * fits the text box. PowerPoint solves this at render time and only caches
 * the solved values in the `fontScale` / `lnSpcReduction` attributes when the
 * file is saved. A normAutofit element WITHOUT stored attributes therefore
 * still means "shrink to fit" — the spec's attribute defaults (100% / 0%)
 * describe the serialized cache, not the rendered size.
 *
 * The search algorithm itself is application behavior, not specified by
 * ECMA-376. This module models the sequence observed in desktop PowerPoint:
 *
 * - fontScale steps down from 100% through
 *   {100, 92.5, 85, 77.5, 70, 65, 60, 55, 50, 45, 40, 35, 30, 25}%,
 *   which matches the values PowerPoint persists in saved files
 *   (fontScale="92500", "77500", ... are common in the wild). The floor is
 *   25%; below that PowerPoint lets the text overflow.
 * - lnSpcReduction steps {0, 10, 20}% alongside the font scale: 0% at 100%,
 *   10% for scales <= 92.5%, and 20% for scales <= 70% (observed
 *   approximation of PowerPoint's coupling; some PowerPoint versions emit
 *   other pairings, e.g. fontScale 100% with a bare lnSpcReduction, but this
 *   coupled table reproduces the common shrink ladder).
 *
 * The first (largest-scale) candidate whose measured layout height fits the
 * available text-box height wins. Height is weakly monotonic in the candidate
 * index (smaller fonts wrap to the same or fewer lines and produce the same
 * or smaller line advances), so the solver binary-searches the table instead
 * of walking it linearly: at most ~5 measured layout passes plus the initial
 * 100% guard pass.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.FIT_EPSILON_PX = exports.AUTOFIT_FLOOR_INDEX = exports.AUTOFIT_CANDIDATES = void 0;
exports.solveAutofitScale = solveAutofitScale;
/**
 * Observed PowerPoint font-scale ladder, largest first. See module docs.
 */
const FONT_SCALE_STEPS = [
    1, 0.925, 0.85, 0.775, 0.7, 0.65, 0.6, 0.55, 0.5, 0.45, 0.4, 0.35, 0.3, 0.25,
];
/**
 * Line-spacing reduction coupled to a font scale: 0% at 100%, 10% for
 * scales <= 92.5%, 20% for scales <= 70%.
 */
function lineSpaceReductionFor(fontScale) {
    if (fontScale >= 1)
        return 0;
    if (fontScale > 0.7)
        return 0.1;
    return 0.2;
}
/**
 * The full candidate table, best (100%, no reduction) first, floor
 * (25%, 20% reduction) last.
 */
exports.AUTOFIT_CANDIDATES = FONT_SCALE_STEPS.map((scale) => ({
    fontScale: scale,
    lineSpaceReduction: lineSpaceReductionFor(scale),
}));
/**
 * Index of the floor candidate (25% font scale). Used when even the smallest
 * scale overflows: PowerPoint stops shrinking at 25% and lets text overflow.
 */
exports.AUTOFIT_FLOOR_INDEX = exports.AUTOFIT_CANDIDATES.length - 1;
/**
 * Tolerance in pixels when comparing a measured layout height against the
 * available height, absorbing floating-point noise from the layout math.
 */
exports.FIT_EPSILON_PX = 0.01;
/**
 * Solves for the first (largest-scale) candidate whose layout fits.
 *
 * @param availableHeight Available text-box height in pixels (insets applied)
 * @param measure Measures a full layout pass (including re-wrapping) at the
 *   given candidate and returns the resulting total text height in pixels.
 *   Must be weakly monotonic: a later (smaller) candidate never measures
 *   taller than an earlier one. Called at most ~6 times; callers should cache
 *   per-index results so the winning layout is not recomputed.
 * @returns Index into AUTOFIT_CANDIDATES of the chosen candidate. Index 0
 *   (100%) short-circuits after a single measurement when the text already
 *   fits. When nothing fits, returns AUTOFIT_FLOOR_INDEX (which is always
 *   measured before returning).
 */
function solveAutofitScale(availableHeight, measure) {
    const fits = (index) => {
        const candidate = exports.AUTOFIT_CANDIDATES[index];
        if (!candidate)
            return false;
        return measure(candidate, index) <= availableHeight + exports.FIT_EPSILON_PX;
    };
    // Guard: skip the solver entirely when the text already fits at 100%.
    if (fits(0)) {
        return 0;
    }
    // Binary search for the smallest index (largest scale) that fits.
    // Monotonicity: if a candidate fits, every later (smaller) one does too.
    let lo = 1;
    let hi = exports.AUTOFIT_FLOOR_INDEX;
    let best = -1;
    while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (fits(mid)) {
            best = mid;
            hi = mid - 1;
        }
        else {
            lo = mid + 1;
        }
    }
    // When nothing fits the loop's final probe was the floor itself (lo can
    // only pass hi by failing at lo === hi === AUTOFIT_FLOOR_INDEX), so the
    // floor candidate has always been measured.
    return best >= 0 ? best : exports.AUTOFIT_FLOOR_INDEX;
}
