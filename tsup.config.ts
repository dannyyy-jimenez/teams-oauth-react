/// <reference types="tsup" />
import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["cjs", "esm"],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  treeshake: true,
  minify: false,
  external: [
    "react",
    "react-dom",
    "next",
    "@azure/msal-browser",
    "@azure/msal-react",
    "@microsoft/microsoft-graph-client",
  ],
  esbuildOptions(options: { banner?: { js?: string } }) {
    options.banner = {
      js: '"use client";',
    };
  },
});
