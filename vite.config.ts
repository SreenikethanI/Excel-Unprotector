import { defineConfig } from "vite";
import preact from "@preact/preset-vite";

// https://vite.dev/config/
export default defineConfig({
  plugins: [preact()],
  base: "/Excel-Unprotector/",
  build: {
    sourcemap: true,
  },
  server: {
    host: true,
    open: true,
  },
  preview: {
    host: true,
    open: true,
  },
});
