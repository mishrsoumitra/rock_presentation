import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['packages/*/test/**/*.test.ts', 'apps/*/test/**/*.test.ts', 'tools/**/*.test.ts'],
    // your own decks (corpus/local) take minutes each: only with `corepack pnpm test:decks`
    exclude: [...configDefaults.exclude, ...(process.env.TEST_DECKS ? [] : ['tools/render/decks.test.ts'])],
    environment: 'node',
    testTimeout: 60_000,
  },
});
