import { defineConfig } from 'tsup'

export default defineConfig({
    entry: ['src/index.ts'],
    splitting: false,
    sourcemap: true,
    clean: true,
    bundle: true,
    minify: false,
    target: 'es2024',
    noExternal: [/(.*)/],
    format: ['esm', 'cjs'],
})