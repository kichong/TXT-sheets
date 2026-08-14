import { builtinModules } from 'node:module';
import { resolve } from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    outDir: 'dist/main',
    emptyOutDir: false,
    sourcemap: true,
    lib: { entry: resolve(__dirname, 'src/main.ts'), formats: ['cjs'], fileName: () => 'main.cjs' },
    rollupOptions: {
      external: ['electron', 'exceljs', 'jszip', 'papaparse', ...builtinModules, ...builtinModules.map((name) => `node:${name}`)],
    },
  },
});
