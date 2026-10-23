// Usage: npx tsx scripts/render.ts <file.xlsx> <outDir> — builds the XLSX and PDF like the site does.
import fs from "node:fs";
import path from "node:path";
// eslint-disable-next-line @typescript-eslint/no-require-imports
const pdfmake = require("pdfmake");
import { readFileGrid } from "../src/lib/readFile";
import { analyze } from "../src/lib/cashbook";
import { buildXlsx } from "../src/lib/buildXlsx";
import { buildPdfDefinition, PDF_FONTS, PDF_FONT_FILES } from "../src/lib/buildPdf";
import { fileBase } from "../src/lib/layout";
import { kyivNow } from "../src/lib/format";

async function main() {
  const [file, out, fix] = process.argv.slice(2);
  const buf = fs.readFileSync(file);
  const grid = await readFileGrid(file, buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  let res = analyze(grid, path.basename(file), { fixed: new Set(), kept: new Set() });
  console.log("issues:", res.issues.map((i) => `${i.id} | ${i.where} | ${i.text} | ${i.fix}`));
  const fixed = new Set<string>();
  while (fix && res.issues.length) {
    res.issues.forEach((i) => fixed.add(i.id));
    res = analyze(grid, path.basename(file), { fixed, kept: new Set() });
    console.log("after fix:", res.issues.map((i) => `${i.id} | ${i.text}`));
  }
  console.log("stats", res.stats, "fixed", res.fixed.map((i) => i.id));
  const logo = fs.readFileSync("public/logo.png").toString("base64");
  const createdAt = kyivNow();
  fs.mkdirSync(out, { recursive: true });
  const base = path.join(out, fileBase(res.doc));
  fs.writeFileSync(base + ".xlsx", Buffer.from(await buildXlsx(res.doc, { logoBase64: "data:image/png;base64," + logo, createdAt })));
  for (const f of PDF_FONT_FILES) pdfmake.virtualfs.writeFileSync(f, fs.readFileSync(path.join("public/fonts", f)));
  pdfmake.addFonts(PDF_FONTS);
  pdfmake.setUrlAccessPolicy(() => false);
  const def = buildPdfDefinition(res.doc, { logoDataUrl: "data:image/png;base64," + logo, createdAt });
  fs.writeFileSync(base + ".pdf", await pdfmake.createPdf(def).getBuffer());
  console.log("written", base);
}
main();
