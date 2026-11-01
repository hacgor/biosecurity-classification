import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/biosecurity-classification/',
  server: {
    port: 5173,
    strictPort: true,
  },
})
