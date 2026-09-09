import { createRequire } from "node:module";

// NODE_PATH is provided by the caller so the bundled CommonJS package can be
// reused without installing dependencies in the staged Typst project.
const require = createRequire(import.meta.url);
const sharp = require("sharp");

const [input, output, widthArg] = process.argv.slice(2);
if (!input || !output) {
  throw new Error("usage: svg_to_png.mjs INPUT.svg OUTPUT.png [WIDTH_PX]");
}
const width = Number.parseInt(widthArg ?? "1800", 10);
await sharp(input, { density: 240 })
  .resize({ width, withoutEnlargement: true })
  .png({ compressionLevel: 9 })
  .toFile(output);
