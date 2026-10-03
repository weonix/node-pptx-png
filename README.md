# node-pptx-png

Pure-Node PPTX → PNG / JPEG / WebP / SVG / PDF renderer. No LibreOffice, no headless browser, no Docker — `npm install` and render, in-process.

[![CI](https://img.shields.io/github/actions/workflow/status/sdruckerfig/node-pptx-png/ci.yml?branch=main&label=CI)](https://github.com/sdruckerfig/node-pptx-png/actions/workflows/ci.yml)
[![npm version](https://img.shields.io/npm/v/node-pptx-png)](https://www.npmjs.com/package/node-pptx-png)
[![license](https://img.shields.io/github/license/sdruckerfig/node-pptx-png)](./LICENSE)

```js
import { openPresentation } from 'node-pptx-png';
import * as fs from 'node:fs/promises';

const doc = await openPresentation('./deck.pptx');
for await (const slide of doc.slides({ preset: 'preview' })) {
  await fs.writeFile(`slide-${slide.slideNumber}.png`, slide.imageData);
}
doc.close();
```

![Title slide rendered by node-pptx-png](https://raw.githubusercontent.com/sdruckerfig/node-pptx-png/main/docs/images/slide-1.png)

_All images in this README were rendered by this library from the deck committed at [`test/fixtures/test-presentation.pptx`](test/fixtures/test-presentation.pptx) (width 960, `pngOptimization: 'web'`)._

## Why

Every existing route from PPTX to images drags in heavyweight machinery: a 400–800 MB LibreOffice install driven over a subprocess, a Docker sidecar, a JVM bridge, or a metered cloud API your documents have to leave your infrastructure for. `node-pptx-png` is a native renderer for the format itself — it parses the ECMA-376 OOXML and draws slides with [skia-canvas](https://github.com/samizdatco/skia-canvas), in your Node process.

- **~30 MB installed.** One prebuilt Skia binary plus a few small pure-JS dependencies (`jszip`, `fast-xml-parser`).
- **No external processes.** Cold start is module load — no `soffice` spawn, no browser launch, no daemon to babysit.
- **True parallelism.** A worker-thread pool renders slides across CPU cores; there is no singleton process to serialize behind.
- **Serverless-ready.** skia-canvas ships prebuilt binaries for macOS, Linux (glibc *and* musl), and Windows on x64 *and* arm64 — including dedicated AWS Lambda builds. No compiler toolchain needed on any of these platforms.
- **Your files stay yours.** Rendering happens in-process; nothing is uploaded anywhere.
- **Honest about fidelity.** Anything the renderer skips or substitutes is reported through a structured [warnings channel](#warnings-fidelity-you-can-inspect) — machine-readable codes you can gate CI on, not silent output drift.
- **MIT licensed.** No AGPL server, no per-call pricing.

Output formats: **PNG**, **JPEG**, **WebP**, vector **SVG**, and whole-deck multi-page vector **PDF** — all encoded natively by Skia, no extra dependencies.

## Rendered output

A picture-heavy content slide and a styled table, straight from the renderer:

![Picture fills, cropping, and an alpha watermark](https://raw.githubusercontent.com/sdruckerfig/node-pptx-png/main/docs/images/slide-4.png)

![Table with table-style banding, merges, and a diagonal border](https://raw.githubusercontent.com/sdruckerfig/node-pptx-png/main/docs/images/slide-5.png)

And the zero-code path — the bundled CLI:

```console
$ npx node-pptx-png deck.pptx -o out/ --format webp --preset preview --slides 1-5
slide 1 -> out/slide-1.webp (640x360)
slide 2 -> out/slide-2.webp (640x360)
slide 3 -> out/slide-3.webp (640x360)
slide 4 -> out/slide-4.webp (640x360)
slide 5 -> out/slide-5.webp (640x360)
```

## How it compares

| | Install weight | Cold start | Concurrency | Serverless | License / cost | Fidelity approach |
|---|---|---|---|---|---|---|
| **node-pptx-png** | ~30 MB (`npm install`) | Module load, in-process | Worker-thread pool across cores | Prebuilds for Lambda, glibc + musl, x64 + arm64 | MIT | Native ECMA-376 renderer on Skia; structured warnings for gaps |
| libreoffice-convert | ~400–800 MB (LibreOffice) | `soffice` subprocess spawn | `soffice` is a singleton; parallel loads contend or crash | Custom layers; 3008 MB RAM + 45 s timeout guidance on Lambda, x86 only | Free (LibreOffice is MPL) | LibreOffice Impress import filter; silent font substitution |
| Gotenberg | 424–695 MB Docker image | Container + HTTP round-trip | Scale out containers | Container platforms only | MIT | LibreOffice inside; office formats convert to **PDF only** — no PPTX→PNG in one step |
| ONLYOFFICE Document Server | ~1.8 GB server | Server boot | 20-connection cap (Community Edition) | No — long-running server | AGPL v3 or commercial | Own Office-compatible engine; strong fidelity |
| Aspose.Slides for Node.js | Java library + JVM bridge | JVM start | JVM-side threading | Impractical (JVM inside a function) | Commercial, $999+ | Proprietary engine; gold-standard fidelity |
| Cloud APIs (CloudConvert, ConvertAPI, …) | None locally | Network round-trip per call | Provider rate limits | Yes (it's SaaS) | Per-call; reaches $200–500+/mo at automation volume | Varies — and your documents leave your infrastructure |
| Puppeteer + JS viewer | Hundreds of MB (headless Chromium) | Browser launch | One browser page per render | Needs special Chromium builds | MIT viewers | Capped by whichever browser-side viewer library you drive |

Competitor figures are compiled from public documentation and issue trackers (2026); sources are cited in [`ROADMAP.md`](ROADMAP.md). To be fair in the other direction: ONLYOFFICE and Aspose currently render some complex content more faithfully than we do — their trade-offs are operational (size, license, cost), ours are listed openly in [Known limitations](#known-limitations).

## Installation

```bash
npm install node-pptx-png
```

Requires **Node.js >= 20.9**. skia-canvas downloads a prebuilt binary for macOS, Linux (glibc and musl), and Windows, on x64 and arm64 — no compiler or system packages needed on those platforms. (On anything else it falls back to building from source.)

`sharp` is an **optional** peer dependency, used only by the [PNG optimization](#png-optimization) feature:

```bash
npm install sharp   # only if you want pngOptimization
```

## Document API: parse once, render many

`openPresentation()` parses the archive, presentation part, theme, and table styles a single time and returns a reusable `PptxDocument` handle — the right tool whenever you touch a deck more than once (thumbnails plus full-size, one slide at a time, streaming).

```js
import { openPresentation } from 'node-pptx-png';

const doc = await openPresentation('./presentation.pptx'); // path or Buffer
try {
  console.log(doc.slideCount, doc.size); // size = { widthEmu, heightEmu }

  // Render one slide (1-based numbering)
  const title = await doc.slide(1, { preset: 'hd' });

  // Buffer everything, exactly like renderPresentation
  const all = await doc.renderAll({ preset: 'preview' });
} finally {
  doc.close();
}
```

`PptxDocument` is disposable: in TypeScript 5.2+ (any Node target) or plain Node.js 24+, `await using doc = await openPresentation(...)` closes it automatically.

### Streaming slides

`doc.slides()` is an async generator — process a 200-slide deck without buffering 200 PNGs, with slide selection, progress, and cancellation:

```js
const controller = new AbortController();

for await (const slide of doc.slides({
  slides: { from: 1, to: 5 }, // or a list in emit order: [3, 1, 2]
  preset: 'thumb',
  signal: controller.signal,
  onProgress: ({ done, total, slideNumber }) =>
    console.log(`${done}/${total} (slide ${slideNumber})`),
})) {
  await fs.writeFile(`slide-${slide.slideNumber}.png`, slide.imageData);
}
```

### Size presets

Instead of pixel math, pass a named `preset` (height is auto-calculated from the slide aspect ratio): `'thumb'` (256 px wide), `'preview'` (640), `'hd'` (1920), `'4k'` (3840). Explicit `width`/`height` or `scale` take precedence; the preset is then ignored.

```js
const thumb = await doc.slide(1, { preset: 'thumb' }); // 256x144 for 16:9
```

## Output formats: PNG, JPEG, WebP, SVG

`format` selects the encoding for every render call; `quality` (0–1) applies to the lossy formats (`jpeg`, `webp`). SVG output is vector: `imageData` holds the UTF-8 SVG document (mime type `image/svg+xml`). `pngOptimization` only applies to PNG output.

```js
const webp = await doc.slide(1, { format: 'webp', quality: 0.8 });
const svg = await doc.slide(1, { format: 'svg' });
await fs.writeFile('slide-1.svg', svg.imageData); // image/svg+xml bytes
```

The legacy `jpegQuality` (1–100) option still works for JPEG but is deprecated; `quality` wins when both are set.

## Whole-deck vector PDF

`exportPdf()` draws each selected slide onto one page of a shared canvas and exports a single multi-page **vector** PDF — text and shapes stay crisp at any zoom, in-process, with no PDF library involved. Page size follows the usual size options; raster options (`format`, `quality`, `pngOptimization`) are ignored. Unlike the raster methods it rejects on failure (a partial PDF is never produced).

```js
const pdf = await doc.exportPdf(); // all slides
await fs.writeFile('deck.pdf', pdf);

// Or a selection, with the usual size options
const excerpt = await doc.exportPdf({ slides: { from: 1, to: 5 }, scale: 1 });
```

## Parallel rendering with a worker pool

`createRenderPool()` spreads a deck across worker threads (`node:worker_threads`): slide numbers are sharded round-robin, each worker opens the file itself and renders its share, and the encoded images are transferred (not copied) back to the main thread, which reassembles them in order with warnings preserved. Buffer inputs are copied to each worker; file paths are cheapest. Always `close()` the pool — live workers keep the process alive.

```js
import { createRenderPool } from 'node-pptx-png';

const pool = createRenderPool({ workers: 'auto' }); // or an explicit count
try {
  const result = await pool.render('./presentation.pptx', {
    preset: 'thumb',
    slides: { from: 1, to: 50 },
  });
  // result is a PresentationRenderResult: ordered slides, warnings, errors
} finally {
  await pool.close();
}
```

Notes: the pool spawns workers lazily on the first render and reuses them; `render()` mirrors `renderPresentation` semantics (input problems resolve to a presentation-level error result rather than rejecting). The worker entry lives in the built package (`dist/pool/worker.js`); when a bundler relocates files, point the pool at it with the `workerUrl` option. The CommonJS build resolves its worker via `__dirname` (`dist/cjs/pool/worker.js`).

## CLI

The package ships a dependency-free CLI for zero-code conversion:

```bash
npx node-pptx-png deck.pptx -o out/ --format webp --preset thumb --slides 1-5
```

```
Usage: node-pptx-png <input.pptx> [options]

  -o, --outdir <dir>      Output directory (default: current directory)
  --format <fmt>          Image format: png | jpeg | webp | svg (default: png)
  --width <N>             Output width in pixels (height keeps aspect ratio)
  --scale <N>             Scale factor on the slide's native 96-DPI size
  --preset <name>         thumb (256) | preview (640) | hd (1920) | 4k (3840)
  --slides <list>         Slides to render, e.g. "1-5,8" (1-based)
  --pdf <file>            Write a whole-deck vector PDF; without --format,
                          the PDF replaces the per-slide images
  --quality <N>           Quality 0-1 for jpeg/webp (default: 0.9)
  --quiet                 Suppress progress output
  --json                  Print a JSON result summary (slides, files,
                          warnings) to stdout
```

Slide images are written as `slide-<n>.<ext>`. `--width`, `--scale`, and `--preset` are mutually exclusive. Exit codes: `0` success, `1` any slide or PDF failed, `2` invalid arguments. Progress lines go to stderr, so stdout stays clean for scripting: with `--json` it carries only the JSON summary (including structured warnings).

```bash
# Vector PDF of the first five slides
npx node-pptx-png deck.pptx --pdf deck.pdf --slides 1-5

# PDF and WebP images in one run
npx node-pptx-png deck.pptx -o out/ --format webp --pdf out/deck.pdf
```

## MCP server (AI agents)

An MCP (Model Context Protocol) stdio server ships as a second binary, so
Claude and other AI agents can rasterize decks directly:

```bash
claude mcp add pptx -- npx node-pptx-png-mcp
```

Tools: `render_pptx` (slides to PNG/WebP/SVG files), `pptx_info` (slide
count and dimensions), `export_pdf` (whole-deck vector PDF). See
[docs/mcp.md](docs/mcp.md) for Claude Desktop configuration and the full
tool reference.

## Fonts

Font fidelity is the single biggest factor in how closely a rendered slide matches PowerPoint: a substituted font with different metrics shifts every line-wrap point. The library addresses this on three levels.

### Embedded fonts

When a deck was saved with **File → Options → Save → Embed fonts in the file**, the embedded fonts (`p:embeddedFontLst` → `ppt/fonts/*.fntdata`) are extracted and registered with the renderer automatically, so text renders with the deck's own fonts. This is on by default; disable it with `fonts: { useEmbeddedFonts: false }`.

Supported embedded-font containers:

| Container                              | Support                                       |
| -------------------------------------- | --------------------------------------------- |
| Plain TTF / OTF / TTC                  | Registered directly                           |
| EOT wrapper (plain or XOR-encrypted)   | Unwrapped and registered                      |
| ODTTF-obfuscated (GUID part names)     | Deobfuscated per ECMA-376 and registered      |
| EOT with MicroType Express compression | Skipped with a warning (falls back to chains) |

Note: recent PowerPoint versions often write embedded fonts as MicroType Express compressed EOT, which cannot be decompressed without a licensed MTX implementation; those decks fall back to the substitution chains below.

### Registering your own fonts

Supply font files (or buffers) to register before rendering — the most reliable way to guarantee fidelity on servers and in containers:

```js
const result = await renderer.renderPresentation('./deck.pptx', {
  fonts: {
    register: [
      { family: 'Open Sans', source: './fonts/OpenSans-Regular.ttf' },
      { family: 'Open Sans', source: './fonts/OpenSans-Bold.ttf' },
      { family: 'Brand Font', source: brandFontBuffer },
    ],
  },
});
```

Registration is process-global: fonts stay available for subsequent renders.

### Fallback chains and metric-compatible substitutes

When a requested font is not installed, the renderer walks a fallback chain. Metric-compatible open fonts are preferred automatically **when installed**, before the generic fallbacks, because they share glyph widths with their Microsoft counterparts and preserve line breaks:

| Requested font    | Metric-compatible substitute |
| ----------------- | ---------------------------- |
| Calibri           | Carlito                      |
| Cambria           | Caladea                      |
| Arial / Helvetica | Liberation Sans              |
| Times New Roman   | Liberation Serif             |

On Linux servers and in Docker images, installing the fontconfig packages for these fonts significantly improves fidelity:

```bash
# Debian/Ubuntu
sudo apt-get install fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation

# Alpine
apk add font-carlito font-liberation
```

Custom chains can be supplied per render and replace the built-in chain for that family:

```js
const result = await renderer.renderPresentation('./deck.pptx', {
  fonts: {
    fallbacks: {
      Calibri: ['Carlito', 'Noto Sans', 'sans-serif'],
      'Futura PT': ['Futura', 'Century Gothic', 'sans-serif'],
    },
  },
});
```

### `FontOptions`

| Option             | Type                       | Default | Description                                                                             |
| ------------------ | -------------------------- | ------- | --------------------------------------------------------------------------------------- |
| `useEmbeddedFonts` | `boolean`                  | `true`  | Extract and register fonts embedded in the PPTX before rendering                        |
| `register`         | `FontRegistration[]`       | —       | Font files to register: `{ family, source (path or Buffer), weight?, style? }`          |
| `fallbacks`        | `Record<string, string[]>` | —       | Custom fallback chains, merged over the built-ins (replaces the chain for listed fonts) |

## Warnings: fidelity you can inspect

Anything the renderer skips, substitutes, or only partially draws is reported as a **structured warning** — on each `SlideRenderResult.warnings` and aggregated (deduplicated) on `PresentationRenderResult.warnings`. A clean render omits the field entirely, so `result.warnings` doubles as a CI fidelity gate.

```js
const result = await doc.renderAll({ preset: 'hd' });
for (const w of result.warnings ?? []) {
  console.warn(`[${w.code}] slide ${w.slideNumber ?? '-'}: ${w.message}`);
}
// e.g. [missing-font] slide 1: Font not available, falling back ...
```

Each `RenderWarning` carries a stable kebab-case `code`, a human-readable `message`, the 1-based `slideNumber` (absent for presentation-level warnings such as embedded-font registration), the affected `elementId` when known, and machine-readable `detail`:

| Code | Emitted when |
| --- | --- |
| `missing-font` | A requested font isn't available; a fallback chain was used |
| `embedded-font-unsupported` | An embedded font container couldn't be decoded (e.g. MicroType Express EOT) |
| `unsupported-element` | An element type the renderer doesn't draw was skipped |
| `unsupported-geometry` | A shape geometry couldn't be built; a fallback was drawn |
| `unsupported-chart-type` | A chart type with no renderer; a neutral placeholder was drawn |
| `chart-failed` | A chart failed to parse or render |
| `image-decode-failed` | An image part couldn't be decoded |
| `metafile-partial` | An EMF/WMF metafile rendered only partially (e.g. EMF+-only) |
| `diagram-missing-part` | SmartArt without its pre-rendered drawing part |
| `layout-load-failed` | A slide layout/master part failed to load |
| `media-missing` | A referenced media part is missing from the archive |
| `effect-skipped` | An effect couldn't be applied |
| `other` | Anything else |

## PNG optimization

For smaller PNG files, enable optimization (requires the optional `sharp` peer dependency):

```js
// 'web' preset: palette quantization, 60-70% smaller
const result = await renderer.renderPresentation('./presentation.pptx', {
  pngOptimization: 'web',
});

// Or lossless presets: 'fast' | 'balanced' | 'maximum'
const result2 = await renderer.renderPresentation('./presentation.pptx', {
  pngOptimization: 'balanced',
});

// Or custom options
const result3 = await renderer.renderPresentation('./presentation.pptx', {
  pngOptimization: {
    compressionLevel: 9,
    adaptiveFiltering: true,
    palette: true,
    colors: 128,
    quality: 80,
  },
});
```

| Preset       | Size Reduction | Description                                     |
| ------------ | -------------- | ----------------------------------------------- |
| `'none'`     | 0%             | No optimization (fastest)                       |
| `'fast'`     | ~1-2%          | Quick lossless compression                      |
| `'balanced'` | ~2-3%          | Lossless with adaptive filtering                |
| `'maximum'`  | ~2-3%          | Best lossless compression                       |
| `'web'`      | **60-70%**     | Palette quantization (may affect photo quality) |

## What renders

The renderer targets ECMA-376 (Office Open XML) directly. Current coverage:

| Area | Support |
| --- | --- |
| **Geometry** | All 187 ECMA-376 preset geometries (generated from `presetShapeDefinitions.xml`, adjust values and guide formulas evaluated), custom geometry (`a:custGeom`), connectors with arrowheads, grouped shapes with nested transforms |
| **Text** | Master/layout/placeholder inheritance (incl. `p:txStyles`), run styling (bold/italic/underline/strikethrough, solid and theme colors), bullet and numbered lists, alignment, line spacing, superscript/subscript, stored autofit (`a:normAutofit` fontScale / lnSpcReduction), CJK line breaking with `a:ea`/`a:cs` per-run fonts |
| **Fills & lines** | Solid fills, linear and radial gradients, picture fills, theme style matrix references (`fillRef`/`lnRef`/`fontRef` via `a:fmtScheme`), color maps (`p:clrMap`/`p:clrMapOvr`), theme background references (`p:bgRef`) |
| **Tables** | `tableStyles.xml` resolution (banding, first/last row/column, corner cells), merged cells, per-cell borders including diagonals, cell fills, in-cell text styling |
| **Charts** | Bar/column, line, pie, doughnut, area, scatter, bubble; combo charts (multiple plot types on one plot area); titles, legends, data labels. Unknown types draw a neutral placeholder plus an `unsupported-chart-type` warning |
| **SmartArt** | Rendered from the pre-baked diagram drawing part (`dsp:` shapes, nested groups) that PowerPoint writes — no layout engine required |
| **Images & metafiles** | Raster images (embedded and picture fills), EMF and WMF metafiles ([MS-EMF]/[MS-WMF] GDI record subset; EMF+ Dual files render via their embedded GDI fallback stream) |
| **Effects** | Outer shadow (including perspective shadows), inner shadow, glow, soft edges, reflection |
| **Slide plumbing** | Master/layout inheritance, placeholder resolution, backgrounds (solid/gradient/picture/`bgRef`), `mc:AlternateContent` fallback handling |

### Known limitations

Stated plainly, because the warnings channel will tell you anyway:

- **EMF+-only metafiles** (no GDI dual stream) render only the partial GDI subset (`metafile-partial` warning). Office usually writes EMF+ Dual, which renders fine.
- **MicroType Express compressed embedded fonts** (common in recent PowerPoint) cannot be decompressed; those decks fall back to substitution chains (`embedded-font-unsupported` warning).
- **No bidirectional (RTL) text shaping.**
- **Autofit honors PowerPoint's stored shrink values only** — there is no iterative shrink-to-fit, so text that was never laid out by PowerPoint (e.g. decks generated by other tools) may overflow its shape.
- **Font state is process-global** (skia-canvas `FontLibrary` has no per-render scope): concurrent renders using *different* `fonts.fallbacks` in one process can race. One configuration per process is safe, as is the worker pool.
- **Animations, transitions, and embedded audio/video** are out of scope for static rendering; the slide's static content is drawn.

## Serverless

This library's happy place: tens of MB deployed, module-load cold start, no subprocess — and skia-canvas publishes dedicated **AWS Lambda prebuilds** (x64 and arm64) plus musl builds for Alpine containers. See [`docs/serverless.md`](docs/serverless.md) for Lambda/Cloud Run deployment guides and [`docs/docker.md`](docs/docker.md) for a Docker image with metric-compatible fonts preinstalled.

## API Reference

### `openPresentation(input, options?)`

Opens a presentation once for repeated rendering.

- `input`: `Buffer | string` - File path or buffer containing PPTX data
- `options`: `{ logLevel? }` - Optional open options
- Returns: `Promise<PptxDocument>` (rejects if the input is not a readable PPTX)

### `PptxDocument`

A parsed, reusable presentation handle. All slide numbers on this API are **1-based**, consistent with the `slideNumbers` render option (the legacy `renderSlide` index is 0-based).

| Member        | Signature                                                        | Description                                                                                                                                                                               |
| ------------- | ---------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `slideCount`  | `number`                                                         | Number of slides in the deck                                                                                                                                                              |
| `size`        | `{ widthEmu: number; heightEmu: number }`                        | Native slide size in EMU (914400 EMU = 1 inch)                                                                                                                                            |
| `slide()`     | `(slideNumber, options?) => Promise<SlideRenderResult>`          | Renders one slide by 1-based number. Bad per-call input (out-of-range number, invalid `scale`/`preset`) resolves to a `success: false` result                                             |
| `slides()`    | `(options?: SlidesOptions) => AsyncGenerator<SlideRenderResult>` | Streams rendered slides in order — see `SlidesOptions` below                                                                                                                              |
| `renderAll()` | `(options?) => Promise<PresentationRenderResult>`                | Buffers all (or `slideNumbers`-selected) slides, matching `renderPresentation` semantics exactly                                                                                          |
| `exportPdf()` | `(options?: ExportPdfOptions) => Promise<Buffer>`                | Exports the selected slides as one multi-page vector PDF (`slides` selection plus the usual size options). Rejects on invalid selection or a failed slide                                 |
| `close()`     | `() => void`                                                     | Releases the archive and caches. Idempotent; render calls after `close()` throw. Also exposed as `Symbol.dispose` / `Symbol.asyncDispose`, so `using` / `await using` close automatically |

### `SlidesOptions`

`RenderOptions` plus:

| Option       | Type                                         | Description                                                                                                                                                                                             |
| ------------ | -------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `slides`     | `number[] \| { from?: number; to?: number }` | Slides to emit. List form: 1-based numbers, emitted in the given order (invalid numbers yield per-slide failure results). Range form: inclusive bounds, clamped to the deck; defaults to the whole deck |
| `signal`     | `AbortSignal`                                | Checked before each slide; on abort, iteration throws the signal's reason (an `AbortError` `DOMException` by default)                                                                                   |
| `onProgress` | `(e: { done, total, slideNumber }) => void`  | Called after each slide finishes rendering, immediately before its result is yielded                                                                                                                    |

### `createRenderPool(options?)`

Creates a worker-thread render pool.

- `options`: `RenderPoolOptions` — `workers: number | 'auto'` (default `'auto'`: `os.availableParallelism() - 1`, minimum 1), `workerUrl` to point at a relocated worker entry under bundlers
- Returns: `RenderPool` with `render(input, options?) => Promise<PresentationRenderResult>` (options accept `RenderOptions` plus the `slides` selection) and `close() => Promise<void>`

### `PptxImageRenderer`

The classic one-shot API. Each call re-opens and re-parses the input — prefer `openPresentation` when rendering the same deck more than once.

#### Constructor

```typescript
new PptxImageRenderer(options?: { logLevel?: 'debug' | 'info' | 'warn' | 'error' | 'silent' })
```

#### Methods

##### `renderPresentation(input, options?)`

Renders all slides (or specified slides) in a presentation.

- `input`: `Buffer | string` - File path or buffer containing PPTX data
- `options`: `RenderOptions` - Optional rendering options
- Returns: `Promise<PresentationRenderResult>`

```js
import { PptxImageRenderer } from 'node-pptx-png';

const renderer = new PptxImageRenderer();
const result = await renderer.renderPresentation('./presentation.pptx', {
  format: 'png',
  scale: 1.0,
  slideNumbers: [1, 2, 3], // optional: render only specific slides
});

for (const slide of result.slides) {
  if (slide.success) fs.writeFileSync(`slide-${slide.slideNumber}.png`, slide.imageData);
}
```

Buffers work anywhere a path does:

```js
const pptxBuffer = fs.readFileSync('./presentation.pptx');
const result = await renderer.renderPresentation(pptxBuffer);
```

##### `renderSlide(input, slideIndex, options?)`

Renders a single slide. Note: `slideIndex` is **0-based**, unlike the 1-based `slideNumbers` render option.

- `input`: `Buffer | string` - File path or buffer containing PPTX data
- `slideIndex`: `number` - 0-based slide index
- `options`: `RenderOptions` - Optional rendering options
- Returns: `Promise<SlideRenderResult>` (bad input such as an out-of-range index or invalid `scale` resolves to a result with `success: false`)

##### `getSlideCount(input)`

Gets the number of slides in a presentation.

- `input`: `Buffer | string` - File path or buffer containing PPTX data
- Returns: `Promise<number>`

##### `getSlideDimensions(input)`

Gets the native slide size in EMU (914400 EMU = 1 inch).

- `input`: `Buffer | string` - File path or buffer containing PPTX data
- Returns: `Promise<{ width: number; height: number }>`

### `RenderOptions`

| Option            | Type                                                 | Default  | Description                                                                                                                                                                                                                            |
| ----------------- | ---------------------------------------------------- | -------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `width`           | `number`                                             | `1920`   | Target width in pixels (height auto-calculated from the slide aspect ratio unless `height` is set)                                                                                                                                     |
| `height`          | `number`                                             | auto     | Target height in pixels                                                                                                                                                                                                                |
| `scale`           | `number`                                             | —        | Multiplier on the slide's native 96-DPI size (e.g., `1.0` renders a 16:9 slide at 1280x720). Must be positive. Only used when neither `width` nor `height` is set; precedence is `width`/`height` > `scale` > `preset` > default width |
| `preset`          | `'thumb' \| 'preview' \| 'hd' \| '4k'`               | —        | Named output width (256 / 640 / 1920 / 3840 px; height auto-calculated from the aspect ratio). Ignored when `width`, `height`, or `scale` is present                                                                                   |
| `format`          | `'png' \| 'jpeg' \| 'webp' \| 'svg'`                 | `'png'`  | Output format. SVG output is vector: `imageData` holds the UTF-8 SVG document                                                                                                                                                          |
| `quality`         | `number`                                             | `0.9`    | Encoding quality (0-1) for the lossy formats (`jpeg`, `webp`). Wins over `jpegQuality` when both are set                                                                                                                               |
| `jpegQuality`     | `number`                                             | `90`     | Deprecated JPEG-only alias (1-100); use `quality` instead                                                                                                                                                                              |
| `gpu`             | `boolean \| 'auto'`                                  | `'auto'` | GPU rendering toggle for the skia-canvas surface: `true` requests GPU (falls back to CPU when unavailable), `false` forces CPU, `'auto'` keeps the skia-canvas default. The mode used is recorded in the debug log                      |
| `slideNumbers`    | `number[]`                                           | all      | Specific slides to render (1-based), `renderPresentation`/`renderAll` only. Invalid numbers produce per-slide failure results (and `errors` entries) instead of throwing                                                                |
| `backgroundColor` | `string`                                             | —        | Override slide background color (hex, e.g., `'#FFFFFF'`)                                                                                                                                                                               |
| `logLevel`        | `'debug' \| 'info' \| 'warn' \| 'error' \| 'silent'` | `'warn'` | Logging level for diagnostic output                                                                                                                                                                                                    |
| `debugMode`       | `boolean`                                            | `false`  | Draw bounding boxes and element IDs for debugging                                                                                                                                                                                      |
| `pngOptimization` | `string \| object`                                   | `'none'` | PNG optimization preset or custom options (see [PNG optimization](#png-optimization))                                                                                                                                                  |
| `fonts`           | `FontOptions`                                        | —        | Font handling: embedded font extraction, user font registration, custom fallback chains (see [Fonts](#fonts))                                                                                                                          |

Options explicitly set to `undefined` are treated as omitted and fall back to their defaults.

### `PngOptimizationOptions`

For custom PNG optimization:

| Option              | Type      | Default | Description                               |
| ------------------- | --------- | ------- | ----------------------------------------- |
| `compressionLevel`  | `number`  | `6`     | Compression level (0-9, higher = smaller) |
| `adaptiveFiltering` | `boolean` | `true`  | Use adaptive row filtering                |
| `palette`           | `boolean` | `false` | Convert to indexed PNG (max 256 colors)   |
| `colors`            | `number`  | `256`   | Max colors for palette mode (2-256)       |
| `quality`           | `number`  | `90`    | Palette quantization quality (1-100)      |
| `dither`            | `number`  | `1.0`   | Dithering strength for palette mode (0-1) |

### `SlideRenderResult`

| Property       | Type               | Description                                                    |
| -------------- | ------------------ | -------------------------------------------------------------- |
| `slideIndex`   | `number`           | Zero-based slide index                                         |
| `slideNumber`  | `number`           | One-based slide number                                         |
| `imageData`    | `Buffer`           | Rendered image data (encoding follows `format`)                |
| `width`        | `number`           | Image width in pixels                                          |
| `height`       | `number`           | Image height in pixels                                         |
| `success`      | `boolean`          | Whether rendering succeeded                                    |
| `errorMessage` | `string?`          | Error message if failed                                        |
| `warnings`     | `RenderWarning[]?` | Structured warnings for this slide; omitted on a clean render  |

### `PresentationRenderResult`

| Property           | Type                  | Description                                                                                                                                                                                                                                            |
| ------------------ | --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `slides`           | `SlideRenderResult[]` | Results for each rendered slide (all slides, or those selected via `slideNumbers`)                                                                                                                                                                     |
| `totalSlides`      | `number`              | Total number of slides in the presentation                                                                                                                                                                                                             |
| `successfulSlides` | `number`              | Number of successfully rendered slides                                                                                                                                                                                                                 |
| `allSuccessful`    | `boolean`             | Whether all requested slides rendered successfully                                                                                                                                                                                                     |
| `errors`           | `RenderError[]?`      | Slide-level entries for failed slides and a presentation-level entry when the file itself could not be processed; omitted when no errors occurred. `renderPresentation` resolves with a presentation-level error instead of rejecting on invalid input |
| `warnings`         | `RenderWarning[]?`    | Aggregate of presentation-level plus per-slide warnings, deduplicated by code + message + slideNumber; omitted when none were emitted                                                                                                                   |

### Module formats

The package is dual ESM/CJS: `import` resolves `dist/index.js`, `require` resolves `dist/cjs/index.js`, each with its own type definitions.

## Development

```bash
npm install        # install dependencies
npm run build      # dual ESM/CJS build into dist/
npm test           # vitest suite
npm run typecheck  # tsc --noEmit (src + scripts/tests)
npm run lint       # eslint

# Render a deck with your local build
node dist/cli.js ./test/fixtures/test-presentation.pptx -o output/ --preset preview
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the fidelity-testing workflow (SSIM corpus, baselines) and PR conventions, and [SECURITY.md](SECURITY.md) for reporting vulnerabilities.

## License

[MIT](LICENSE)
