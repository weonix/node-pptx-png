# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2026-07-31

A ground-up fidelity and API overhaul: 30 verified rendering bugs fixed, the
rendering pipeline rebuilt across geometry, styles, text inheritance, effects,
tables, charts, SmartArt, CJK, and EMF/WMF, then a v2 API layered on top —
document handle, streaming, structured warnings, WebP/SVG/PDF output, a worker
pool, and a CLI. Test suite grew from 40 to 833 tests; the fidelity corpus
renders 192/192 slides across 4 real-world decks, verified against PowerPoint
reference exports.

### Added

#### API v2

- `openPresentation()` / `PptxDocument`: parse-once/render-many document
  handle with `slide()`, `renderAll()`, and an AsyncIterator `slides()`
  stream supporting slide selection (list or range), `AbortSignal`, and
  `onProgress`; idempotent `close()` plus `Symbol.dispose` /
  `Symbol.asyncDispose` for `using` / `await using`.
- Structured warnings channel: per-slide `warnings[]` with 13
  machine-readable codes (`missing-font`, `metafile-partial`,
  `unsupported-element`, `image-decode-failed`, ...) plus a deduplicated
  presentation-level aggregate.
- Output formats: WebP and vector SVG alongside PNG/JPEG; generalized
  `quality` option (0-1) for lossy formats.
- Whole-deck vector PDF via `doc.exportPdf()` (multi-page skia canvas;
  text and shapes stay vector).
- Worker pool: `createRenderPool()` on `node:worker_threads` with
  round-robin slide sharding, transferable buffers, and `execArgv`
  isolation.
- Size presets `'thumb' | 'preview' | 'hd' | '4k'` and a `gpu`
  (`true | false | 'auto'`) toggle.
- CLI: `npx node-pptx-png <input.pptx>` with `--format`, `--preset`,
- MCP server (`node-pptx-png-mcp`): dependency-free stdio Model Context Protocol server exposing `render_pptx`, `pptx_info`, and `export_pdf` tools for AI-agent integration
  `--slides`, `--pdf`, `--width`, `--scale`, `--quality`, `--quiet`,
  `--json` (structured JSON summary including warnings).

#### Fidelity engine (Phases 1-4)

- Full ECMA-376 preset geometry engine: all 187 preset shapes via a guide
  formula interpreter; custom geometry (`a:custGeom`) in the live render
  path; rectangle fallback replaces silent shape skips.
- Embedded font extraction (TTF/OTF/TTC, ODTTF deobfuscation, EOT
  unwrapping) and a public `fonts` API: `register`, custom `fallbacks`
  chains, `useEmbeddedFonts`; metric-compatible substitution chains
  (Calibri→Carlito, Cambria→Caladea, Arial→Liberation Sans, ...).
- Shape style references (`p:style`: `fillRef`/`lnRef`/`fontRef`), theme
  format-scheme style matrix, `p:clrMap`/`clrMapOvr` color mapping, and
  `p:bgRef` theme backgrounds.
- Full master → layout → slide text inheritance (placeholder resolution),
  `normAutofit` text scaling, `p:defaultTextStyle`, and
  master-placeholder `lstStyle` in the inheritance chain.
- Effects: outer shadows (`a:effectLst`/`a:effectRef`), inner shadow,
  glow, soft edge, reflection, and perspective/picture-fill shadows via a
  silhouette effect renderer.
- Table styles (`tableStyleId` → `tableStyles.xml`) with banding and
  first/last row/column parts; gradient, diagonal, corner, `tblBg`, and
  RTL table features.
- Connector arrowheads (`a:headEnd`/`a:tailEnd`) with line shortening.
- Charts: scatter, doughnut (including concentric), combo, bubble,
  secondary value axis, gap-width/overlap semantics.
- SmartArt rendering via the pre-rendered `dsp:` drawing part, including
  nested groups and `txXfrm`.
- CJK support: per-character wrapping with kinsoku line-break rules
  (cross-fragment) and per-script (`a:ea`/`a:cs`) font resolution.
- EMF/WMF metafile rendering (MS-EMF GDI record subset, EMF+ dual
  handling, embedded DIBs, basic WMF) — pasted Excel tables and charts
  now render.
- Picture `alphaModFix` transparency and hidden-picture handling.
- Implemented the documented-but-missing `scale` and `slideNumbers`
  render options; populated `PresentationRenderResult.errors`.
- Fidelity corpus scoreboard (`npm run corpus`): SSIM scoring against
  LibreOffice ground truth or previous runs, HTML gallery +
  `scoreboard.json`.

#### Packaging & infrastructure

- Dual ESM/CJS build with a conditional `exports` map.
- CI matrix: Node 20/22/24 on Linux (glibc + musl/Alpine, x64 + arm64),
  macOS, and Windows.
- MIT `LICENSE` file.

### Changed

- Dependencies modernized: skia-canvas 3, fast-xml-parser 5, TypeScript 6,
  vitest 4, ESLint 10 (flat config), sharp 0.35 (optional peer).
- Legacy `renderPresentation`/`renderSlide` now delegate to the
  `PptxDocument` render path (verified byte-identical output: SSIM 1.000
  across the 192-slide corpus).
- `jpegQuality` (1-100) is deprecated in favor of `quality` (0-1);
  `quality` wins when both are set.

### Fixed

- `renderSlide()` and `getSlideCount()` always failing.
- Text rendering at 2/3 of the correct size.
- Gradient angle transposition.
- Group child transform composition.
- `resolvePath()` hang on certain relationship targets.
- Text run merging dropping formatting; `a:t` whitespace loss.
- `spcPts`/`spcPct` line-spacing sentinel handling.
- ...and 20+ further verified bugs (see `CODE_REVIEW.md` in the repo
  history for the full list).

### BREAKING CHANGES

- `renderPresentation` no longer rejects on invalid input: it resolves
  with a `PresentationRenderResult` carrying a presentation-level entry in
  `result.errors`. Check `allSuccessful` / `errors` instead of `catch`.
- `allSuccessful` now refers to the *requested* slides (via
  `slideNumbers`), not all slides in the deck.
- Node.js `>= 20.9.0` is required (`engines` is enforced).
- CommonJS consumers are served through the new `exports` map
  (`dist/cjs/`); deep imports into `dist/` internals are not supported.

## [1.0.0] - 2026-01-18

### Added

- Initial release: PPTX to PNG/JPEG conversion in pure Node.js using
  skia-canvas — no LibreOffice, no headless browser.
- Shapes with preset geometries, text with styling and bullets, embedded
  images, tables, basic charts (bar/line/pie), solid/gradient/image
  backgrounds, theme colors, master/layout inheritance.
- Optional PNG optimization via `sharp` (presets `fast`, `balanced`,
  `maximum`, `web`, or custom options).

[Unreleased]: https://github.com/sdruckerfig/node-pptx-png/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/sdruckerfig/node-pptx-png/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/sdruckerfig/node-pptx-png/releases/tag/v1.0.0
