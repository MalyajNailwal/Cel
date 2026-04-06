import { defineConfig, Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import fs from 'fs';

function inlineAssets(): Plugin {
  return {
    name: 'inline-assets',
    transformIndexHtml(html, ctx) {
      return html;
    },
  };
}

export default defineConfig({
  plugins: [react(), inlineAssets()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'src/taskpane.html'),
        commands: path.resolve(__dirname, 'src/commands.html'),
      },
    },
    outDir: 'dist',
  },
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync('./key.pem'),
      cert: fs.readFileSync('./cert.pem'),
    },
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
    proxy: {
      '/api': {
        target: 'http://localhost:8000',
        changeOrigin: true,
        secure: false,
      },
    },
  },
});
