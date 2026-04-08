import { defineConfig } from 'vite';

export default defineConfig({
    build: {
        lib: {
            entry: './src/extension.ts',
            formats: ['es'], // VS Code usually expects CommonJS
            fileName: 'extension',
        },
        rollupOptions: {
            // Mark 'vscode' and Node built-ins as external so Vite doesn't try to bundle them
            external: ['vscode', 'fs', 'path', 'os', 'child_process'],
            output: {
                entryFileNames: '[name].js',
            },
        },
        outDir: 'dist',
        emptyOutDir: true,
        sourcemap: false, // Essential for debugging
    },
});
