/// <reference types="vitest" />
import path from 'path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';

import { cloudflare } from "@cloudflare/vite-plugin";

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    return {
      server: {
        port: 3000,
        host: '0.0.0.0',
      },
      plugins: [
        react(),
        ...(process.env.VITEST || process.env.PLAYWRIGHT_TEST ? [] : [cloudflare()])
      ],
      test: {
        globals: true,
        environment: 'jsdom',
        setupFiles: './tests/setup.ts',
        // `.claude/worktrees` holds agent worktrees, which are full checkouts living
        // inside the repo. Without excluding them the suite collects every test twice
        // and fails on the Playwright specs, which vitest cannot run.
        exclude: [
          'node_modules', 'dist', '.idea', '.git', '.cache',
          '**/tests/e2e/**', '.claude/worktrees/**'
        ],
        pool: 'forks',
        fileParallelism: false,
        maxWorkers: 1,
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