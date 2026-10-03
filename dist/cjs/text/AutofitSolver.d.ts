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
/**
 * One shrink candidate: a font-size multiplier and the fraction by which
 * spacing-derived line advances are reduced. Structurally identical to the
 * stored-value multipliers TextLayoutEngine resolves from a parsed
 * a:normAutofit, so a solved candidate can be fed through the exact same
 * layout code path as stored values.
 */
export interface AutofitCandidate {
    /** Multiplier applied to every resolved font size (1 = no scaling) */
    fontScale: number;
    /** Fraction by which spacing-derived line advances are reduced (0 = none) */
    lineSpaceReduction: number;
}
/**
 * The full candidate table, best (100%, no reduction) first, floor
 * (25%, 20% reduction) last.
 */
export declare const AUTOFIT_CANDIDATES: readonly AutofitCandidate[];
/**
 * Index of the floor candidate (25% font scale). Used when even the smallest
 * scale overflows: PowerPoint stops shrinking at 25% and lets text overflow.
 */
export declare const AUTOFIT_FLOOR_INDEX: number;
/**
 * Tolerance in pixels when comparing a measured layout height against the
 * available height, absorbing floating-point noise from the layout math.
 */
export declare const FIT_EPSILON_PX = 0.01;
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
export declare function solveAutofitScale(availableHeight: number, measure: (candidate: AutofitCandidate, index: number) => number): number;
