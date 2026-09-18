import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "fs";
import path from "path";

const certFile = path.resolve(__dirname, "backend/localhost+1.pem");
const keyFile  = path.resolve(__dirname, "backend/localhost+1-key.pem");
const hasCerts = fs.existsSync(certFile) && fs.existsSync(keyFile);

export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    port: 5173,
    https: hasCerts
      ? { cert: fs.readFileSync(certFile), key: fs.readFileSync(keyFile) }
      : false,
    proxy: {
      "/api": {
        target: hasCerts ? "https://localhost:3001" : "http://localhost:3001",
        changeOrigin: true,
        secure: false,
      },
    },
  },
});