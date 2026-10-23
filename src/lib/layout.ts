import { totals, type Block, type CashBook, type Entry } from "./cashbook";

/** Table columns of the cash book document, shared by the XLSX and PDF builders. */
export type ColKey = "no" | "account" | "party" | "inUah" | "inUsd" | "outUah" | "outUsd" | "rate" | "eq";

export type Col = { key: ColKey; group?: string; title: string; numeric: boolean; xlsxWidth: number; pdfWidth: number | "*" };

export function columns(doc: CashBook): Col[] {
  const cols: Col[] = [
    { key: "no", title: "№", numeric: false, xlsxWidth: 6, pdfWidth: 16 },
    { key: "account", title: "Каса / рахунок", numeric: false, xlsxWidth: 20, pdfWidth: 62 },
    { key: "party", title: "Від кого отримано чи кому видано", numeric: false, xlsxWidth: 36, pdfWidth: "*" },
    { key: "inUah", group: "Надходження", title: "грн", numeric: true, xlsxWidth: 15, pdfWidth: 60 },
    { key: "inUsd", group: "Надходження", title: "$", numeric: true, xlsxWidth: 13, pdfWidth: 52 },
    { key: "outUah", group: "Видаток", title: "грн", numeric: true, xlsxWidth: 15, pdfWidth: 60 },
    { key: "outUsd", group: "Видаток", title: "$", numeric: true, xlsxWidth: 13, pdfWidth: 52 },
  ];
  if (doc.hasRateCol) cols.push({ key: "rate", title: "Курс", numeric: true, xlsxWidth: 8, pdfWidth: 26 });
  if (doc.hasEqCol) cols.push({ key: "eq", title: "Екв., $", numeric: true, xlsxWidth: 13, pdfWidth: 52 });
  return cols;
}

export type Line =
  | { type: "block"; block: Block; label: string }
  | { type: "opening"; block: Block; label: string }
  | { type: "sub"; label: string }
  | { type: "entry"; entry: Entry; block: Block }
  | { type: "turnover"; block: Block; label: string }
  | { type: "closing"; block: Block; label: string };

export function lines(doc: CashBook): Line[] {
  const out: Line[] = [];
  for (const b of doc.blocks) {
    out.push({ type: "block", block: b, label: b.title });
    if (b.opening) out.push({ type: "opening", block: b, label: "Залишок на початок дня" });
    if (b.incomes.length) {
      out.push({ type: "sub", label: "Надходження" });
      for (const e of b.incomes) out.push({ type: "entry", entry: e, block: b });
    }
    if (b.expenses.length) {
      out.push({ type: "sub", label: "Видаток" });
      for (const e of b.expenses) out.push({ type: "entry", entry: e, block: b });
    }
    if (b.incomes.length || b.expenses.length) out.push({ type: "turnover", block: b, label: "Оборот за день" });
    if (b.closing) out.push({ type: "closing", block: b, label: "Залишок на кінець дня" });
  }
  return out;
}

/** Summary per block: opening, incomes, expenses, closing in грн and $. */
export function summary(b: Block) {
  const t = totals(b);
  return [
    { label: "Залишок на початок дня", uah: b.opening?.uah ?? null, usd: b.opening?.usd ?? null },
    { label: "Надходження", uah: t.inUah, usd: t.inUsd },
    { label: "Видаток", uah: t.outUah, usd: t.outUsd },
    { label: "Залишок на кінець дня", uah: b.closing?.uah ?? null, usd: b.closing?.usd ?? null },
  ];
}

const LAT: Record<string, string> = {
  а: "a", б: "b", в: "v", г: "h", ґ: "g", д: "d", е: "e", є: "ie", ж: "zh", з: "z", и: "y", і: "i", ї: "i", й: "i",
  к: "k", л: "l", м: "m", н: "n", о: "o", п: "p", р: "r", с: "s", т: "t", у: "u", ф: "f", х: "kh", ц: "ts", ч: "ch",
  ш: "sh", щ: "shch", ь: "", ю: "iu", я: "ia", ы: "y", э: "e", ё: "io", ъ: "",
};

/** Latin file name: some browsers drop Cyrillic download names. */
export function fileBase(doc: { title: string; date: { d: number; m: number; y: number } | null }) {
  const p = (n: number) => String(n).padStart(2, "0");
  const d = doc.date ? `_${p(doc.date.d)}.${p(doc.date.m + 1)}.${doc.date.y}` : "";
  const title = [...doc.title.trim()]
    .map((ch) => {
      const low = ch.toLowerCase();
      const t = LAT[low];
      if (t == null) return ch;
      return ch !== low && t ? t[0].toUpperCase() + t.slice(1) : t;
    })
    .join("")
    .replace(/[^A-Za-z0-9.-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return `LUXEL_${title || "dokument"}${d}`;
}
