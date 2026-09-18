import { defineConfig } from 'vite';

export default defineConfig({
  base: '/VZHDO/',
  publicDir: 'public',
  build: {
    outDir: 'dist',
    emptyOutDir: true
  }
});
