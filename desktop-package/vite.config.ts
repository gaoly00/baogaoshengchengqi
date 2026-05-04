import { defineConfig } from 'vite';

export default defineConfig({
  base: './', // 使用相对路径，适用于 Electron 本地文件
  server: {
    port: 5000,
    host: '0.0.0.0',
    allowedHosts: true,
  },
  build: {
    assetsDir: 'assets',
    outDir: 'dist',
  },
});
