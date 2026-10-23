import { key, type Cell, type Grid } from "./grid";

export const ACCEPTED = /\.(xlsx|xls)$/i;

/** Reads the main sheet of an .xlsx (ExcelJS) or old .xls (SheetJS) file. */
export async function readFileGrid(name: string, data: ArrayBuffer): Promise<Grid> {
  if (/\.xls$/i.test(name)) return readXls(data);
  const ExcelJS = (await import("exceljs")).default;
  const { mainSheet, readGrid } = await import("./grid");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  return readGrid(mainSheet(wb));
}

async function readXls(data: ArrayBuffer): Promise<Grid> {
  const XLSX = await import("xlsx");
  const wb = XLSX.read(new Uint8Array(data), { type: "array", cellFormula: true, cellHTML: false, cellText: false });
  let best: Grid | null = null;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const cells = new Map<string, Cell>();
    let maxRow = 0;
    let maxCol = 0;
    for (const ref of Object.keys(ws)) {
      if (ref.startsWith("!")) continue;
      const c = ws[ref] as { t?: string; v?: unknown; f?: string };
      const { r, c: col } = XLSX.utils.decode_cell(ref);
      const v =
        c.t === "n" && typeof c.v === "number" && Number.isFinite(c.v) ? c.v
        : c.t === "s" && typeof c.v === "string" ? c.v
        : c.t === "b" ? (c.v ? "TRUE" : "FALSE")
        : null;
      const f = c.f ? String(c.f) : undefined;
      if (v == null && !f) continue;
      if (typeof v === "string" && !v.trim() && !f) continue;
      cells.set(key(r + 1, col + 1), f ? { v, f } : { v });
      maxRow = Math.max(maxRow, r + 1);
      maxCol = Math.max(maxCol, col + 1);
    }
    if (!best || cells.size > best.cells.size) best = { sheetName, maxRow, maxCol, cells };
  }
  if (!best) throw new Error("У файлі немає аркушів");
  return best;
}
