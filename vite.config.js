import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // Relative asset paths so the built dist/ works both at a web root (Netlify)
  // and when loaded via file:// from the packaged Electron app on Windows.
  base: "./",
});
