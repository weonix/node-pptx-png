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
export declare const moduleUrl: string;
//# sourceMappingURL=workerUrl.d.ts.map