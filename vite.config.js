import { defineConfig } from 'vite';
import { crx } from '@crxjs/vite-plugin';
import manifest from './manifest.json';

export default defineConfig({
  plugins: [
    crx({ manifest }),
  ],

  build: {
    outDir: 'dist',
    emptyOutDir: true,
    // Disable sourcemaps in production builds to avoid leaking source structure.
    // Switch to 'inline' during development as needed.
    sourcemap: false,
    minify: 'oxc',
  },

  // Suppress Vite's console clearing so build output is always visible
  clearScreen: false,
});
