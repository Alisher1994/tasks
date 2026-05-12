import { defineConfig } from "vite";

const BACKEND_PORT = Number(process.env.BACKEND_PORT) || 3000;

export default defineConfig({
  root: ".",
  publicDir: "public",
  build: {
    outDir: "dist",
    emptyOutDir: true,
    target: "es2020",
    sourcemap: true
  },
  server: {
    port: 5173,
    strictPort: false,
    proxy: {
      "/api": {
        target: `http://localhost:${BACKEND_PORT}`,
        changeOrigin: true
      },
      "/health": {
        target: `http://localhost:${BACKEND_PORT}`,
        changeOrigin: true
      }
    }
  }
});
