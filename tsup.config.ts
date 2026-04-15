import { defineConfig } from "tsup";

export default defineConfig({
    entry: ["src/index.ts"],
    format: ["cjs"],
    target: "node18",
    sourcemap: false,
    clean: true,
    bundle: true,
    minify: false,
    outDir: "dist",
    banner: {
        js: "#!/usr/bin/env node"
    }
});