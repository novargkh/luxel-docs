import ExcelJS from "exceljs";
import { docDateShort, docDateText, type CashBook, type Doc, type Table } from "./cashbook";
import { columns, lines, type ColKey } from "./layout";

export type BrandAssets = { createdAt: string };

const RED = "FFE92620";
const INK = "FF1A1A1A";
const MUTED = "FF7A7A7A";
const LINE = "FFE4E4E4";
const SOFT = "FFF6F6F6";
const PINK = "FFFDEDEC";
const FONT = "Calibri";
const NUM = '#,##0.00;-#,##0.00;"0,00"';

const font = (o: Partial<ExcelJS.Font> = {}): Partial<ExcelJS.Font> => ({ name: FONT, size: 10, color: { argb: INK }, ...o });
const fill = (argb: string): ExcelJS.Fill => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const thin = (argb = LINE): Partial<ExcelJS.Border> => ({ style: "thin", color: { argb } });

function setupSheet(wb: ExcelJS.Workbook, doc: Doc, brand: BrandAssets, widths: number[]) {
  const ws = wb.addWorksheet(doc.kind === "cashbook" ? "Касовий звіт" : doc.title.slice(0, 31), {
    views: [{ showGridLines: false }],
    pageSetup: {
      paperSize: 9,
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.25, footer: 0.3 },
    },
    headerFooter: {
      oddFooter:
        `&L&"${FONT},Regular"&8&K7A7A7A${doc.title}${docDateShort(doc) ? " · " + docDateShort(doc) : ""}` +
        `&C&"${FONT},Regular"&8&K7A7A7AСторінка &P з &N` +
        `&R&"${FONT},Regular"&8&K7A7A7AСформовано: ${brand.createdAt}`,
    },
  });
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  wb.creator = "LUXEL";
  wb.created = new Date();

  // Header: title, date, creation stamp (no logo in documents).
  ws.getRow(1).height = 24;
  ws.getRow(2).height = 16;
  const last = widths.length;
  const titleCol = 1;
  ws.getCell(1, titleCol).value = doc.title;
  ws.getCell(1, titleCol).font = font({ size: 16, bold: true });
  ws.getCell(1, titleCol).alignment = { vertical: "middle" };
  const sub = [docDateText(doc), doc.kind === "cashbook" && doc.rate ? `курс ${String(doc.rate).replace(".", ",")}` : ""]
    .filter(Boolean)
    .join("  ·  ");
  ws.getCell(2, titleCol).value = sub;
  ws.getCell(2, titleCol).font = font({ size: 10, color: { argb: MUTED } });
  ws.getCell(1, last).value = `Сформовано: ${brand.createdAt}`;
  ws.getCell(1, last).font = font({ size: 8, color: { argb: MUTED } });
  ws.getCell(1, last).alignment = { horizontal: "right", vertical: "middle" };
  for (let c = 1; c <= last; c++) ws.getCell(3, c).border = { bottom: { style: "medium", color: { argb: RED } } };
  ws.getRow(3).height = 6;
  return ws;
}

function headerCell(cell: ExcelJS.Cell, value: string, align: "left" | "right" | "center" = "left") {
  cell.value = value;
  cell.font = font({ bold: true, size: 9 });
  cell.fill = fill(SOFT);
  cell.alignment = { horizontal: align, vertical: "middle", wrapText: true };
  cell.border = { bottom: thin("FFBDBDBD") };
}

export async function buildXlsx(doc: Doc, brand: BrandAssets): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  if (doc.kind === "cashbook") cashbookSheet(wb, doc, brand);
  else tableSheet(wb, doc, brand);
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

function cashbookSheet(wb: ExcelJS.Workbook, doc: CashBook, brand: BrandAssets) {
  const cols = columns(doc);
  const ws = setupSheet(wb, doc, brand, cols.map((c) => c.xlsxWidth));
  const ci = (k: ColKey) => cols.findIndex((c) => c.key === k) + 1;
  const L = (k: ColKey) => ws.getColumn(ci(k)).letter;
  const last = cols.length;
  let r = 5;

  // Table header (two rows: group + currency).
  const h1 = r;
  const h2 = r + 1;
  for (const [i, c] of cols.entries()) {
    if (c.group) {
      headerCell(ws.getCell(h1, i + 1), c.group, "center");
      headerCell(ws.getCell(h2, i + 1), c.title, "right");
    } else {
      headerCell(ws.getCell(h1, i + 1), c.title, c.numeric ? "right" : "left");
      headerCell(ws.getCell(h2, i + 1), "");
      ws.mergeCells(h1, i + 1, h2, i + 1);
    }
  }
  ws.mergeCells(h1, ci("inUah"), h1, ci("inUsd"));
  ws.mergeCells(h1, ci("outUah"), h1, ci("outUsd"));
  ws.getRow(h1).height = 18;
  ws.pageSetup.printTitlesRow = `${h1}:${h2}`;
  ws.views = [{ state: "frozen", ySplit: h2, showGridLines: false }];
  r = h2 + 1;

  const rowsOf = new Map<object, { open?: number; first?: number; last?: number; turnover?: number; closing?: number }>();
  const slot = (b: object) => rowsOf.get(b) ?? (rowsOf.set(b, {}), rowsOf.get(b)!);
  const amountKeys = ["inUah", "inUsd", "outUah", "outUsd"] as const;

  const styleRow = (row: number, opts: { bold?: boolean; bg?: string; top?: string; size?: number; color?: string }) => {
    for (let c = 1; c <= last; c++) {
      const cell = ws.getCell(row, c);
      cell.font = font({ bold: opts.bold, size: opts.size ?? 9, color: { argb: opts.color ?? INK } });
      if (opts.bg) cell.fill = fill(opts.bg);
      cell.border = { bottom: thin(), ...(opts.top ? { top: { style: "thin", color: { argb: opts.top } } } : {}) };
      const col = cols[c - 1];
      cell.alignment = { vertical: "middle", horizontal: col.numeric ? "right" : "left", wrapText: col.key === "party" };
      if (col.numeric) cell.numFmt = col.key === "rate" ? "0.00##" : NUM;
    }
  };

  for (const line of lines(doc)) {
    const row = ws.getRow(r);
    if (line.type === "block") {
      ws.getCell(r, 1).value = line.label.toUpperCase();
      ws.mergeCells(r, 1, r, last);
      ws.getCell(r, 1).font = font({ bold: true, size: 9, color: { argb: RED } });
      ws.getCell(r, 1).alignment = { vertical: "bottom" };
      row.height = 22;
    } else if (line.type === "sub") {
      ws.getCell(r, ci("account")).value = line.label;
      styleRow(r, { bold: true, color: MUTED, size: 8 });
    } else if (line.type === "opening") {
      const b = line.block;
      ws.getCell(r, ci("account")).value = line.label;
      ws.getCell(r, ci("inUah")).value = b.opening?.uah ?? null;
      ws.getCell(r, ci("inUsd")).value = b.opening?.usd ?? null;
      styleRow(r, { bold: true, bg: SOFT });
      slot(b).open = r;
    } else if (line.type === "entry") {
      const e = line.entry;
      ws.getCell(r, ci("no")).value = e.no ? Number(e.no) || e.no : null;
      ws.getCell(r, ci("account")).value = e.account || null;
      ws.getCell(r, ci("party")).value = e.party || null;
      for (const k of amountKeys) {
        const v = e[k];
        const f = e.constFormulas[k];
        ws.getCell(r, ci(k)).value = v == null ? null : f ? { formula: f, result: v } : v;
      }
      if (ci("rate")) ws.getCell(r, ci("rate")).value = e.rate;
      if (ci("eq") && e.eq != null) {
        const derived = e.inUah != null && e.rate && Math.abs(e.inUah / e.rate - e.eq) < 0.005;
        ws.getCell(r, ci("eq")).value = derived && ci("rate")
          ? { formula: `${L("inUah")}${r}/${L("rate")}${r}`, result: e.eq }
          : e.eq;
      }
      styleRow(r, {});
      ws.getCell(r, ci("no")).alignment = { horizontal: "left", vertical: "middle" };
      ws.getCell(r, ci("no")).font = font({ size: 9, color: { argb: MUTED } });
      const s = slot(line.block);
      s.first ??= r;
      s.last = r;
    } else if (line.type === "turnover") {
      const s = slot(line.block);
      ws.getCell(r, ci("account")).value = line.label;
      for (const k of amountKeys) {
        const sum = [...line.block.incomes, ...line.block.expenses].reduce((a, e) => a + (e[k] ?? 0), 0);
        ws.getCell(r, ci(k)).value = { formula: `SUM(${L(k)}${s.first}:${L(k)}${s.last})`, result: sum };
      }
      if (ci("eq")) {
        const sum = [...line.block.incomes, ...line.block.expenses].reduce((a, e) => a + (e.eq ?? 0), 0);
        ws.getCell(r, ci("eq")).value = { formula: `SUM(${L("eq")}${s.first}:${L("eq")}${s.last})`, result: sum };
      }
      styleRow(r, { bold: true, top: "FFBDBDBD" });
      s.turnover = r;
    } else if (line.type === "closing") {
      const b = line.block;
      const s = slot(b);
      ws.getCell(r, ci("account")).value = line.label;
      const pairs = [
        ["inUah", "outUah", b.closing?.uah ?? null],
        ["inUsd", "outUsd", b.closing?.usd ?? null],
      ] as const;
      for (const [inK, outK, v] of pairs) {
        const live = b.balanced && s.open && s.turnover;
        ws.getCell(r, ci(inK)).value = live
          ? { formula: `${L(inK)}${s.open}+${L(inK)}${s.turnover}-${L(outK)}${s.turnover}`, result: v ?? 0 }
          : v;
      }
      if (ci("eq") && b.closing?.eq != null) ws.getCell(r, ci("eq")).value = b.closing.eq;
      styleRow(r, { bold: true, bg: PINK, top: RED, size: 10 });
      s.closing = r;
    }
    r++;
  }

  // Summary after the table: one column pair per block.
  r += 1;
  ws.getCell(r, 1).value = "ПІДСУМОК";
  ws.getCell(r, 1).font = font({ bold: true, size: 9, color: { argb: RED } });
  r++;
  const pairs: [ColKey, ColKey][] = [["inUah", "inUsd"], ["outUah", "outUsd"]];
  if (ci("rate") && ci("eq")) pairs.push(["rate", "eq"]);
  for (let start = 0; start < doc.blocks.length; start += pairs.length) {
    const group = doc.blocks.slice(start, start + pairs.length);
    const hr = r;
    headerCell(ws.getCell(hr, 1), "");
    headerCell(ws.getCell(hr, 2), "");
    headerCell(ws.getCell(hr, 3), "");
    ws.mergeCells(hr, 1, hr + 1, 3);
    group.forEach((b, gi) => {
      const [a, c] = pairs[gi];
      headerCell(ws.getCell(hr, ci(a)), b.title, "center");
      headerCell(ws.getCell(hr, ci(c)), "");
      ws.mergeCells(hr, ci(a), hr, ci(c));
      headerCell(ws.getCell(hr + 1, ci(a)), "грн", "right");
      headerCell(ws.getCell(hr + 1, ci(c)), "$", "right");
    });
    r = hr + 2;
    const labels = ["Залишок на початок дня", "Надходження", "Видаток", "Залишок на кінець дня"];
    labels.forEach((label, li) => {
      ws.getCell(r, 1).value = label;
      ws.mergeCells(r, 1, r, 3);
      group.forEach((b, gi) => {
        const s = rowsOf.get(b) ?? {};
        const [a, c] = pairs[gi];
        const src: [string | undefined, string | undefined] =
          li === 0 ? [s.open ? `${L("inUah")}${s.open}` : undefined, s.open ? `${L("inUsd")}${s.open}` : undefined]
          : li === 1 ? [s.turnover ? `${L("inUah")}${s.turnover}` : undefined, s.turnover ? `${L("inUsd")}${s.turnover}` : undefined]
          : li === 2 ? [s.turnover ? `${L("outUah")}${s.turnover}` : undefined, s.turnover ? `${L("outUsd")}${s.turnover}` : undefined]
          : [s.closing ? `${L("inUah")}${s.closing}` : undefined, s.closing ? `${L("inUsd")}${s.closing}` : undefined];
        const ref = (addr: string | undefined) => {
          if (!addr) return 0;
          const v = ws.getCell(addr).value as ExcelJS.CellValue;
          const result = typeof v === "number" ? v : v && typeof v === "object" && "result" in v ? (v.result as number) : 0;
          return { formula: addr, result };
        };
        ws.getCell(r, ci(a)).value = ref(src[0]);
        ws.getCell(r, ci(c)).value = ref(src[1]);
      });
      const isEdge = li === 0 || li === 3;
      for (let col = 1; col <= last; col++) {
        const cell = ws.getCell(r, col);
        if (col > 3 && !group.some((_, gi) => pairs[gi].includes(cols[col - 1].key))) continue;
        cell.font = font({ size: 9, bold: isEdge });
        cell.border = { bottom: thin() };
        if (li === 3) cell.fill = fill(PINK);
        if (col > 3) {
          cell.numFmt = NUM;
          cell.alignment = { horizontal: "right" };
        }
      }
      r++;
    });
    r++;
  }
}

function tableSheet(wb: ExcelJS.Workbook, doc: Table, brand: BrandAssets) {
  const n = doc.header?.length ?? doc.rows[0]?.length ?? 1;
  const widths = [...Array(n).keys()].map((i) => {
    const longest = Math.max(
      doc.header?.[i]?.length ?? 0,
      ...doc.rows.slice(0, 300).map((r) => (r[i] == null ? 0 : String(r[i]).length)),
    );
    return Math.min(45, Math.max(8, longest + 2));
  });
  const ws = setupSheet(wb, doc, brand, widths);
  let r = 5;
  if (doc.header) {
    doc.header.forEach((h, i) => headerCell(ws.getCell(r, i + 1), h, doc.numericCols[i] ? "right" : "left"));
    ws.pageSetup.printTitlesRow = `${r}:${r}`;
    ws.views = [{ state: "frozen", ySplit: r, showGridLines: false }];
    r++;
  }
  for (const row of doc.rows) {
    row.forEach((v, i) => {
      const cell = ws.getCell(r, i + 1);
      cell.value = v;
      cell.font = font({ size: 9 });
      cell.border = { bottom: thin() };
      if (typeof v === "number") {
        cell.numFmt = Number.isInteger(v) ? "#,##0" : NUM;
        cell.alignment = { horizontal: "right" };
      }
    });
    r++;
  }
}
