import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],

  // Worker config: use ES modules so workers can use import statements
  worker: {
    format: 'es',
  },

  // Ensure xlsx is bundled properly in workers
  optimizeDeps: {
    exclude: [],
    include: ['xlsx'],
  },

  build: {
    // Increase chunk size warning limit (xlsx is bundled into the parser worker)
    chunkSizeWarningLimit: 2000,
  },
})
