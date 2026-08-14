import { resolve } from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    outDir: 'dist/preload',
    emptyOutDir: false,
    sourcemap: true,
    lib: { entry: resolve(__dirname, 'src/preload.ts'), formats: ['cjs'], fileName: () => 'preload.cjs' },
    rollupOptions: { external: ['electron'] },
  },
});
