import type { Workbook, Worksheet, CellValue } from "exceljs";

/** One source cell: its value (cached result for formulas) and formula. */
export type Cell = {
  v: number | string | null;
  f?: string;
};

/** A sheet read into a plain map, 1-based row/col like Excel. */
export type Grid = {
  sheetName: string;
  maxRow: number;
  maxCol: number;
  cells: Map<string, Cell>;
};

export const key = (r: number, c: number) => `${r}:${c}`;

export function colLetter(c: number): string {
  let s = "";
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export const addr = (r: number, c: number) => `${colLetter(c)}${r}`;

function plain(v: CellValue | undefined): number | string | null {
  if (v == null) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    if ("richText" in v) return v.richText.map((t) => t.text).join("");
    if ("formula" in v || "sharedFormula" in v)
      return plain((v as { result?: CellValue }).result);
    if ("text" in v) return String((v as { text: unknown }).text);
    if ("error" in v) return null;
  }
  return null;
}

/** Picks the sheet that has the most filled cells (files have one sheet, but be safe). */
export function mainSheet(wb: Workbook): Worksheet {
  let best: Worksheet | undefined;
  let bestCount = -1;
  wb.eachSheet((ws) => {
    let n = 0;
    ws.eachRow((row) => row.eachCell(() => n++));
    if (n > bestCount) {
      best = ws;
      bestCount = n;
    }
  });
  if (!best) throw new Error("У файлі немає аркушів");
  return best;
}

export function readGrid(ws: Worksheet): Grid {
  const cells = new Map<string, Cell>();
  let maxRow = 0;
  let maxCol = 0;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      // Merged cells repeat the master's value; keep it only once.
      if (cell.isMerged && cell.master && cell.master.address !== cell.address) return;
      const v = plain(cell.value);
      const f = cell.formula ? String(cell.formula) : undefined;
      if (v == null && !f) return;
      if (typeof v === "string" && !v.trim() && !f) return;
      cells.set(key(r, c), f ? { v, f } : { v });
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    });
  });
  return { sheetName: ws.name, maxRow, maxCol, cells };
}
