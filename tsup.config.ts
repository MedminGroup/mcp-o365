import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['cjs'],
  target: 'node20',
  outDir: 'dist',
  clean: true,
  sourcemap: false,
  dts: false,
  // Bundle all deps — avoids node_modules pain when running via npx
  noExternal: [/.*/],
  // Leave Node built-ins as external
  external: [/^node:/],
  // Shims make CJS/ESM interop smooth
  shims: true,
  banner: {
    js: '#!/usr/bin/env node',
  },
  esbuildOptions(options) {
    // Ensure __dirname / __filename work inside CJS output
    options.define = {
      ...options.define,
    };
  },
});
