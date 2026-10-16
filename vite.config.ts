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

const devHttps = fs.existsSync('./key.pem') && fs.existsSync('./cert.pem')
  ? { key: fs.readFileSync('./key.pem'), cert: fs.readFileSync('./cert.pem') }
  : undefined;

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
    https: devHttps,
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
