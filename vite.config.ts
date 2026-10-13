import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      // ExcelJS's main entry pulls in Node built-ins; the dist bundle is
      // self-contained and browser-safe.
      exceljs: 'exceljs/dist/exceljs.min.js',
    },
  },
  optimizeDeps: {
    include: ['exceljs'],
  },
});
