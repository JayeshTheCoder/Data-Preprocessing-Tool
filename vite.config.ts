// vite.config.ts

import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // Rule for file uploads
      '/upload': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      // Rules for all the cleaning processes
      '/clean_sales': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      '/clean_oe': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      '/clean_wc': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      '/clean_pex': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      // Rule for downloads
      '/download': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      // Existing rule for inference
      '/inference': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      '/run_pipeline': {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
      '/run_pipeline/bulk':  {
        target: 'http://localhost:8080',
        changeOrigin: true,
      },
    }
  }
})