import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: { port: 5173, open: true },
  css: {
    // Force an explicit PostCSS config to avoid inheriting Tailwind from parent/global configs
    postcss: {
      plugins: [],
    },
  },
});


