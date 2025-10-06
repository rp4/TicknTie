import { defineConfig } from 'vite'

export default defineConfig({
  server: {
    port: 5173,
    open: true
  },
  build: {
    outDir: 'dist',
    sourcemap: false, // Disable sourcemaps for production
    minify: 'esbuild', // Use esbuild for faster minification
    rollupOptions: {
      output: {
        // Optimize chunk size
        chunkFileNames: 'assets/[name]-[hash].js',
        entryFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash].[ext]'
      }
    },
    // Optimize dependencies
    reportCompressedSize: false, // Faster builds
    chunkSizeWarningLimit: 2000 // Increase warning limit
  }
})