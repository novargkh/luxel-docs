// Usage: npx tsx scripts/check.ts <file.xlsx> — prints the model and issues.
import ExcelJS from "exceljs";
import { mainSheet, readGrid } from "../src/lib/grid";
import { analyze, totals } from "../src/lib/cashbook";

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(process.argv[2]);
  const grid = readGrid(mainSheet(wb));
  const { doc, issues } = analyze(grid, process.argv[2], { fixed: new Set(), kept: new Set() });
  if (doc.kind === "cashbook") {
    console.log(doc.title, doc.date, doc.rate, "rateCol", doc.hasRateCol, "eq", doc.hasEqCol);
    for (const b of doc.blocks)
      console.log(b.title, b.opening, b.closing, b.incomes.length, b.expenses.length, totals(b), b.balanced);
    console.log(doc.blocks[1]?.incomes.slice(0, 3), doc.blocks[1]?.expenses.slice(-2));
  } else console.log(doc.title, doc.header, doc.rows.length);
  console.log("issues", issues);
}
main();
