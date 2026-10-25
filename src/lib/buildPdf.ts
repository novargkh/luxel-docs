import type { Content, TDocumentDefinitions, TableCell } from "pdfmake/interfaces";
import { docDateShort, docDateText, type CashBook, type Doc, type Table } from "./cashbook";
import { money } from "./format";
import { columns, lines, summary } from "./layout";

export type PdfBrand = { createdAt: string };

const RED = "#E92620";
const INK = "#1A1A1A";
const MUTED = "#7A7A7A";
const LINE = "#E4E4E4";
const SOFT = "#F6F6F6";
const PINK = "#FDEDEC";

export const PDF_FONTS = {
  Manrope: {
    normal: "Manrope_400Regular.ttf",
    bold: "Manrope_700Bold.ttf",
    italics: "Manrope_400Regular.ttf",
    bolditalics: "Manrope_700Bold.ttf",
  },
  ManropeSemi: {
    normal: "Manrope_600SemiBold.ttf",
    bold: "Manrope_800ExtraBold.ttf",
    italics: "Manrope_600SemiBold.ttf",
    bolditalics: "Manrope_800ExtraBold.ttf",
  },
};
export const PDF_FONT_FILES = [
  "Manrope_400Regular.ttf",
  "Manrope_600SemiBold.ttf",
  "Manrope_700Bold.ttf",
  "Manrope_800ExtraBold.ttf",
];

const tableLayout = {
  hLineWidth: (i: number, node: { table: { headerRows?: number } }) => (i === (node.table.headerRows ?? 0) ? 0.8 : 0.4),
  vLineWidth: () => 0,
  hLineColor: (i: number, node: { table: { headerRows?: number } }) => (i === (node.table.headerRows ?? 0) ? "#BDBDBD" : LINE),
  paddingLeft: () => 4,
  paddingRight: () => 4,
  paddingTop: () => 2.5,
  paddingBottom: () => 2.5,
};

function frame(doc: Doc, brand: PdfBrand, body: Content[]): TDocumentDefinitions {
  const sub = [docDateText(doc), doc.kind === "cashbook" && doc.rate ? `курс ${String(doc.rate).replace(".", ",")}` : ""]
    .filter(Boolean)
    .join("   ·   ");
  return {
    pageSize: "A4",
    pageOrientation: "portrait",
    pageMargins: [32, 50, 32, 44],
    info: { title: `LUXEL — ${doc.title}${docDateShort(doc) ? " " + docDateShort(doc) : ""}`, author: "LUXEL", creator: "LUXEL" },
    defaultStyle: { font: "Manrope", fontSize: 8, color: INK, lineHeight: 1.15 },
    header: () => ({
      margin: [32, 24, 32, 0],
      text: `${doc.title}${docDateShort(doc) ? "  ·  " + docDateShort(doc) : ""}`,
      alignment: "right",
      fontSize: 7.5,
      color: MUTED,
    }),
    footer: (page: number, pages: number) => ({
      margin: [32, 14, 32, 0],
      stack: [
        { canvas: [{ type: "line", x1: 0, y1: 0, x2: 531, y2: 0, lineWidth: 0.4, lineColor: LINE }] },
        {
          margin: [0, 6, 0, 0],
          columns: [
            { text: `Сформовано: ${brand.createdAt}`, fontSize: 7, color: MUTED },
            { text: `Сторінка ${page} з ${pages}`, alignment: "right", fontSize: 7, color: MUTED },
          ],
        },
      ],
    }),
    content: [
      { text: doc.title, font: "ManropeSemi", bold: true, fontSize: 20, margin: [0, 2, 0, 2] },
      { text: sub, fontSize: 9, color: MUTED, margin: [0, 0, 0, 8] },
      { canvas: [{ type: "rect", x: 0, y: 0, w: 28, h: 2.5, color: RED }], margin: [0, 0, 0, 14] },
      ...body,
    ],
  };
}

export function buildPdfDefinition(doc: Doc, brand: PdfBrand): TDocumentDefinitions {
  return doc.kind === "cashbook" ? cashbookPdf(doc, brand) : tablePdf(doc, brand);
}

function summaryCards(doc: CashBook): Content[] {
  return doc.blocks.map((b) => {
    const s = summary(b);
    const card = (label: string, uah: number | null, usd: number | null, accent: boolean): Content => ({
      stack: [
        { text: label.toUpperCase(), fontSize: 6.5, color: accent ? RED : MUTED, characterSpacing: 0.4, margin: [0, 0, 0, 4] },
        { text: uah == null ? "—" : `${money(uah)} ₴`, font: "ManropeSemi", bold: accent, fontSize: 10.5 },
        { text: usd == null ? "" : `${money(usd)} $`, fontSize: 8.5, color: accent ? INK : MUTED, margin: [0, 1, 0, 0] },
      ],
    });
    return {
      unbreakable: true,
      margin: [0, 0, 0, 12],
      stack: [
        { text: b.title, font: "ManropeSemi", fontSize: 9, margin: [0, 0, 0, 6] },
        {
          table: {
            widths: ["*", "*", "*", "*"],
            body: [
              s.map((x, i) => ({
                ...(card(x.label, x.uah, x.usd, i === 3) as object),
                fillColor: i === 3 ? PINK : SOFT,
                margin: [8, 7, 8, 7],
              })) as TableCell[],
            ],
          },
          layout: { hLineWidth: () => 0, vLineWidth: (i: number) => (i > 0 && i < 4 ? 3 : 0), vLineColor: () => "#FFFFFF", paddingLeft: () => 0, paddingRight: () => 0, paddingTop: () => 0, paddingBottom: () => 0 },
        },
      ],
    };
  });
}

function cashbookPdf(doc: CashBook, brand: PdfBrand): TDocumentDefinitions {
  const cols = columns(doc);
  const n = cols.length;
  const hdr = (text: string, alignment: "left" | "right" | "center" = "left", extra: object = {}): TableCell => ({
    text, alignment, bold: true, fontSize: 7, color: INK, fillColor: SOFT, ...extra,
  });
  const g1: TableCell[] = [];
  const g2: TableCell[] = [];
  for (let i = 0; i < n; i++) {
    const c = cols[i];
    if (c.group) {
      if (cols[i - 1]?.group !== c.group) g1.push(hdr(c.group, "center", { colSpan: 2 }));
      else g1.push({ text: "" });
      g2.push(hdr(c.title, "right"));
    } else {
      g1.push(hdr(c.title, c.numeric ? "right" : "left", { rowSpan: 2 }));
      g2.push({ text: "" });
    }
  }
  const body: TableCell[][] = [g1, g2];
  const fmt = (v: number | null | undefined) => (v == null ? "" : money(v));
  const span = (text: string, style: object, fill?: string): TableCell[] => [
    { text, colSpan: n, ...style, ...(fill ? { fillColor: fill } : {}) },
    ...Array(n - 1).fill({ text: "" }),
  ];
  const row = (cells: Record<string, string>, style: object = {}, fill?: string): TableCell[] =>
    cols.map((c) => ({
      text: cells[c.key] ?? "",
      alignment: c.numeric ? "right" : "left",
      noWrap: c.numeric,
      ...style,
      ...(c.key === "no" ? { color: MUTED } : {}),
      ...(fill ? { fillColor: fill } : {}),
    }));

  for (const line of lines(doc)) {
    if (line.type === "block") {
      body.push(span(line.label.toUpperCase(), { bold: true, color: RED, fontSize: 7.5, characterSpacing: 0.6, margin: [0, 8, 0, 1] }));
    } else if (line.type === "sub") {
      body.push(span(line.label, { bold: true, color: MUTED, fontSize: 7, margin: [0, 3, 0, 0] }));
    } else if (line.type === "opening") {
      const o = line.block.opening!;
      body.push(row({ account: line.label, inUah: fmt(o.uah), inUsd: fmt(o.usd) }, { bold: true }, SOFT));
    } else if (line.type === "entry") {
      const e = line.entry;
      body.push(row({
        no: e.no, account: e.account, party: e.party,
        inUah: fmt(e.inUah), inUsd: fmt(e.inUsd), outUah: fmt(e.outUah), outUsd: fmt(e.outUsd),
        rate: e.rate == null ? "" : String(e.rate).replace(".", ","), eq: fmt(e.eq),
      }));
    } else if (line.type === "turnover") {
      const all = [...line.block.incomes, ...line.block.expenses];
      const s = (k: "inUah" | "inUsd" | "outUah" | "outUsd" | "eq") => fmt(all.reduce((a, e) => a + (e[k] ?? 0), 0));
      body.push(row({ account: line.label, inUah: s("inUah"), inUsd: s("inUsd"), outUah: s("outUah"), outUsd: s("outUsd"), eq: s("eq") }, { bold: true }));
    } else if (line.type === "closing") {
      const c = line.block.closing!;
      body.push(row({ account: line.label, inUah: fmt(c.uah), inUsd: fmt(c.usd), eq: fmt(c.eq) }, { bold: true }, PINK));
    }
  }

  // Account labels like "Залишок на початок дня" span the number column too.
  for (const r of body.slice(2)) {
    const acc = r[1] as { text?: string; bold?: boolean };
    if (acc?.bold && !(r[0] as { text?: string }).text && !(r[0] as { colSpan?: number }).colSpan) {
      r[0] = { ...(acc as object), colSpan: 3 } as TableCell;
      r[1] = { text: "" };
      r[2] = { text: "" };
    }
  }

  return frame(doc, brand, [
    ...summaryCards(doc),
    {
      table: { headerRows: 2, dontBreakRows: true, widths: cols.map((c) => c.pdfWidth), body },
      layout: tableLayout,
      margin: [0, 4, 0, 0],
    },
  ]);
}

function tablePdf(doc: Table, brand: PdfBrand): TDocumentDefinitions {
  const n = doc.header?.length ?? doc.rows[0]?.length ?? 1;
  const body: TableCell[][] = [];
  if (doc.header) body.push(doc.header.map((h, i) => ({ text: h, bold: true, fontSize: 7, fillColor: SOFT, alignment: doc.numericCols[i] ? "right" : "left" })));
  for (const r of doc.rows)
    body.push([...Array(n).keys()].map((i) => {
      const v = r[i];
      return typeof v === "number"
        ? { text: Number.isInteger(v) ? money(v, 0) : money(v), alignment: "right", noWrap: true }
        : { text: v == null ? "" : String(v) };
    }));
  if (!body.length) body.push([{ text: "Порожній аркуш" }]);
  const widths = [...Array(n).keys()].map((i) => (doc.numericCols[i] ? "auto" : "*"));
  return {
    ...frame(doc, brand, [
      { table: { headerRows: doc.header ? 1 : 0, dontBreakRows: true, widths, body }, layout: tableLayout },
    ]),
    pageOrientation: n > 7 ? "landscape" : "portrait",
  };
}
