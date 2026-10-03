"use strict";
/**
 * ESM-only capture of this module's URL, used by RenderPool to locate the
 * worker entry file next to the compiled pool modules.
 *
 * This is deliberately isolated in its own module: `import.meta` is a
 * parse error in CommonJS, so the CJS transpile of this file (dist/cjs)
 * cannot be loaded. RenderPool therefore imports it lazily and only on the
 * ESM path (when `__dirname` is undefined); the CJS build resolves the
 * worker via `__dirname` and never touches this file.
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.moduleUrl = void 0;
exports.moduleUrl = 
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore -- import.meta only parses in the ESM build; the CJS
// transpile of this module is intentionally never loaded (see above)
import.meta.url;
