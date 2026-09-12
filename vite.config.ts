import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';
import { resolve } from 'node:path';

export default defineConfig({
  plugins: [react()],
  server: {
    host: '127.0.0.1',
    port: 5173,
    strictPort: false,
    fs: {
      deny: [
        '.env',
        '.env.*',
        '**/.git/**',
        '**/.qa/**',
        resolve('tmp').replaceAll('\\', '/') + '/**',
        '**/.npm-cache/**',
        '**/*.docx',
        '**/*.doc',
        '**/*.pdf',
      ],
    },
  },
  test: { environment: 'jsdom', include: ['tests/**/*.test.ts'] },
});
