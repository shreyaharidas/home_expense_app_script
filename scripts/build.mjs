import { build } from "esbuild";
import { cp, mkdir, rm } from "node:fs/promises";

const outputDirectory = "dist";

await rm(outputDirectory, { recursive: true, force: true });
await mkdir(outputDirectory, { recursive: true });

await build({
  entryPoints: ["src/Code.ts"],
  bundle: true,
  format: "iife",
  globalName: "HomeExpenseApp",
  target: "es2019",
  outfile: `${outputDirectory}/Code.js`,
  footer: {
    js: [
      "function onCalculate() { return HomeExpenseApp.onCalculate.apply(null, arguments); }",
      "function handleEdit(e) { return HomeExpenseApp.handleEdit(e); }",
    ].join("\n"),
  },
});

await cp("src/appsscript.json", `${outputDirectory}/appsscript.json`);
