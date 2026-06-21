/// <reference types="vitest" />
import path from 'path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    return {
      server: {
        port: 3000,
        host: '0.0.0.0',
      },
      plugins: [react()],
      test: {
        globals: true,
        environment: 'jsdom',
        setupFiles: './tests/setup.ts',
        exclude: ['node_modules', 'dist', '.idea', '.git', '.cache', 'tests/e2e/**'],
        coverage: {
          provider: 'v8',
          reporter: ['text', 'json', 'html'],
        },
      },
      define: {
        'process.env.API_KEY': JSON.stringify(env.GEMINI_API_KEY),
        'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY)
      },
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
          ...(mode === 'test' ? {
            '../services/browserTestRunner': path.resolve(__dirname, 'tests/vitestAdapter.ts'),
            './services/browserTestRunner': path.resolve(__dirname, 'tests/vitestAdapter.ts'),
            '/services/browserTestRunner': path.resolve(__dirname, 'tests/vitestAdapter.ts'),
          } : {})
        }
      }
    };
});
