import { addr, key, type Grid } from "./grid";
import { Evaluator, Unsupported } from "./formula";
import { longDate, money, parseDate, round2, shortDate, tidy, type DateParts } from "./format";

export type Amounts = { inUah: number | null; inUsd: number | null; outUah: number | null; outUsd: number | null };

export type Entry = Amounts & {
  row: number;
  no: string;
  account: string;
  party: string;
  rate: number | null;
  eq: number | null;
  /** Source formulas without cell references (e.g. "499100/45.3"), kept in the XLSX. */
  constFormulas: Partial<Record<keyof Amounts, string>>;
};

export type Balance = { row: number; uah: number | null; usd: number | null; eq: number | null };

export type Block = {
  title: string;
  opening: Balance | null;
  incomes: Entry[];
  expenses: Entry[];
  closing: Balance | null;
  /** True when closing = opening + incomes − expenses (both currencies). */
  balanced: boolean;
};

export type CashBook = {
  kind: "cashbook";
  title: string;
  date: DateParts | null;
  rate: number | null;
  hasRateCol: boolean;
  hasEqCol: boolean;
  blocks: Block[];
};

export type Table = {
  kind: "table";
  title: string;
  date: DateParts | null;
  header: string[] | null;
  rows: (string | number | null)[][];
  numericCols: boolean[];
};

export type Doc = CashBook | Table;

export type IssueKind = "formula" | "balance" | "conversion";

export type Issue = {
  id: string;
  kind: IssueKind;
  /** Short place, e.g. "Рядок 152, грн". */
  where: string;
  text: string;
  /** What "Виправити" will write. */
  fix: string;
  /** Cell overrides the fix applies (absent for formula issues: recompute). */
  set?: { r: number; c: number; v: number }[];
  /** Formula cell to keep cached when the user keeps it as is. */
  formulaCell?: { r: number; c: number };
};

export type Choices = { fixed: Set<string>; kept: Set<string> };

const TOL = 0.011;

type Cols = {
  inUah: number; inUsd: number | null; outUah: number; outUsd: number | null;
  rate: number | null; eq: number | null; no: number | null;
  labelCols: number[]; accountCol: number | null; partyCol: number | null;
  firstDataRow: number;
};

function text(v: unknown): string {
  return typeof v === "string" ? v : v == null ? "" : String(v);
}

function findLayout(grid: Grid): Cols | null {
  const get = (r: number, c: number) => grid.cells.get(key(r, c))?.v ?? null;
  for (let r = 1; r <= Math.min(grid.maxRow, 20); r++) {
    let inCol = 0;
    let outCol = 0;
    let rateCol: number | null = null;
    for (let c = 1; c <= grid.maxCol; c++) {
      const t = text(get(r, c)).toLowerCase();
      if (!inCol && /надходж|приход/.test(t)) inCol = c;
      if (!outCol && /видат|расход|витрат/.test(t)) outCol = c;
      if (/^\s*курс\s*$/.test(t)) rateCol = c;
    }
    if (!inCol || !outCol || outCol <= inCol) continue;

    // Currency sub-header: the next rows with "грн".
    let sub = 0;
    for (let rr = r + 1; rr <= r + 3; rr++) {
      for (let c = inCol; c <= grid.maxCol; c++) if (/грн|uah/i.test(text(get(rr, c)))) sub = rr;
      if (sub) break;
    }
    const isUah = (c: number) => sub > 0 && /грн|uah/i.test(text(get(sub, c)));
    const isUsd = (c: number) => sub > 0 && /дол|\$|usd/i.test(text(get(sub, c)));
    const pick = (from: number, to: number, f: (c: number) => boolean) => {
      for (let c = from; c < to; c++) if (f(c)) return c;
      return null;
    };
    const end = rateCol ?? grid.maxCol + 1;
    const inUah = pick(inCol, outCol, isUah) ?? inCol;
    const inUsd = pick(inCol, outCol, isUsd) ?? (sub ? null : null);
    const outUah = pick(outCol, end, isUah) ?? outCol;
    const outUsd = pick(outCol, end, isUsd);
    const firstDataRow = (sub || r) + 1;

    // Number column: left of the amounts, mostly small integers.
    const counts = new Map<number, number>();
    const textCounts = new Map<number, number>();
    for (let rr = firstDataRow; rr <= grid.maxRow; rr++)
      for (let c = 1; c < inUah; c++) {
        const v = get(rr, c);
        if (typeof v === "number" && Number.isInteger(v) && v > 0 && v < 10000) counts.set(c, (counts.get(c) ?? 0) + 1);
        if (typeof v === "string") textCounts.set(c, (textCounts.get(c) ?? 0) + 1);
      }
    let no: number | null = null;
    for (const [c, n] of counts) if (n >= 3 && (no == null || n > (counts.get(no) ?? 0))) no = c;
    const labelCols = [...Array(inUah - 1).keys()].map((i) => i + 1).filter((c) => c !== no);
    const textColsBy = (pred: (c: number) => boolean) =>
      [...textCounts.entries()].filter(([c]) => pred(c)).sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
    const accountCol = no != null ? textColsBy((c) => c < no) : null;
    const partyCol = no != null ? textColsBy((c) => c > no) : textColsBy(() => true);

    // Equivalent column: right of the rate, has numbers but no header.
    let eq: number | null = null;
    if (rateCol) {
      for (let c = rateCol + 1; c <= grid.maxCol; c++) {
        let n = 0;
        for (let rr = firstDataRow; rr <= grid.maxRow; rr++) if (typeof get(rr, c) === "number" || grid.cells.get(key(rr, c))?.f) n++;
        if (n >= 2) {
          eq = c;
          break;
        }
      }
    }
    return { inUah, inUsd, outUah, outUsd, rate: rateCol, eq, no, labelCols, accountCol, partyCol, firstDataRow };
  }
  return null;
}

function headerInfo(grid: Grid, upto: number, fileName: string) {
  let date: DateParts | null = null;
  let rate: number | null = null;
  for (let r = 1; r < upto; r++)
    for (let c = 1; c <= grid.maxCol; c++) {
      const t = text(grid.cells.get(key(r, c))?.v);
      if (!date) date = parseDate(t);
      const m = t.match(/курс\s*([\d]+(?:[.,]\d+)?)/i);
      if (m && rate == null) rate = parseFloat(m[1].replace(",", "."));
    }
  if (!date) date = parseDate(fileName);
  return { date, rate };
}

const CONV = /(\d[\d\s]*(?:[.,]\d+)?)\s*\/\s*(\d+(?:[.,]\d+)?)/;
const toNum = (s: string) => parseFloat(s.replace(/\s/g, "").replace(",", "."));

/**
 * Reads the cash book, checks the arithmetic and returns the document model
 * plus the open issues. `choices` are the user's answers to earlier issues.
 */
export type Stats = { formulas: number; conversions: number; balances: number };

export type Analysis = {
  doc: Doc;
  /** Open issues the user still has to decide on. */
  issues: Issue[];
  /** Issues the user chose to keep as they are. */
  kept: Issue[];
  /** Issues already fixed. */
  fixed: Issue[];
  /** How many things were checked, for the "no errors" message. */
  stats: Stats;
};

/**
 * Reads the cash book, checks the arithmetic and returns the document model
 * plus the open issues. `choices` are the user's answers to earlier issues.
 */
export function analyze(grid: Grid, fileName: string, choices: Choices): Analysis {
  const ev = new Evaluator(grid);
  const formulaCount = [...grid.cells.values()].filter((c) => c.f).length;
  const layout = findLayout(grid);
  if (!layout) {
    const all = checkFormulas(grid, ev);
    for (const i of all) if (choices.kept.has(i.id) && i.formulaCell) ev.keepCached.add(key(i.formulaCell.r, i.formulaCell.c));
    ev.reset();
    return {
      doc: genericTable(grid, ev, fileName),
      issues: all.filter((i) => !choices.kept.has(i.id) && !choices.fixed.has(i.id)),
      kept: all.filter((i) => choices.kept.has(i.id)),
      fixed: all.filter((i) => choices.fixed.has(i.id)),
      stats: { formulas: formulaCount, conversions: 0, balances: 0 },
    };
  }

  // Every issue met on the way, in its original wording (before any fix).
  const seen = new Map<string, Issue>();
  // Fixes can change sums that other checks rely on, so settle them in rounds.
  for (let round = 0; round < 5; round++) {
    ev.reset();
    const formulaIssues = checkFormulas(grid, ev);
    for (const i of formulaIssues) if (choices.kept.has(i.id) && i.formulaCell) ev.keepCached.add(key(i.formulaCell.r, i.formulaCell.c));
    ev.reset();
    const m = buildCashBook(grid, ev, layout, fileName);
    const all = [...formulaIssues, ...checkConversions(m, layout).issues, ...checkBalances(m, layout).issues];
    for (const i of all) if (!seen.has(i.id)) seen.set(i.id, i);
    let changed = false;
    for (const i of all) {
      if (!choices.fixed.has(i.id) || !i.set) continue;
      for (const s of i.set) {
        const k = key(s.r, s.c);
        if (ev.overrides.get(k) !== s.v) {
          ev.overrides.set(k, s.v);
          changed = true;
        }
      }
    }
    if (!changed) break;
  }
  ev.reset();
  const model = buildCashBook(grid, ev, layout, fileName);
  const conv = checkConversions(model, layout);
  const bal = checkBalances(model, layout);
  const open = (xs: Issue[]) => xs.filter((i) => !choices.kept.has(i.id) && !choices.fixed.has(i.id));
  const first = open([...checkFormulas(grid, ev), ...conv.issues]);
  // Balances depend on the rows above, so ask about them once those are settled.
  const issues = first.length ? first : open(bal.issues);
  return {
    doc: model,
    issues,
    kept: [...seen.values()].filter((i) => choices.kept.has(i.id)),
    fixed: [...seen.values()].filter((i) => choices.fixed.has(i.id)),
    stats: { formulas: formulaCount, conversions: conv.checked, balances: bal.checked },
  };
}

function checkFormulas(grid: Grid, ev: Evaluator): Issue[] {
  const out: Issue[] = [];
  for (const [k, cell] of grid.cells) {
    if (!cell.f || typeof cell.v !== "number") continue;
    const [r, c] = k.split(":").map(Number);
    let calc: number | string;
    try {
      calc = ev.evaluate(cell.f);
    } catch (e) {
      if (e instanceof Unsupported) continue;
      throw e;
    }
    if (typeof calc !== "number" || Math.abs(calc - cell.v) <= 0.005) continue;
    out.push({
      id: `formula:${addr(r, c)}`,
      kind: "formula",
      where: `Клітинка ${addr(r, c)}`,
      text: `Формула =${cell.f.replace(/^=/, "")} дає ${money(calc)}, а у файлі збережено ${money(cell.v)} (файл не перерахований).`,
      fix: `Записати ${money(calc)}`,
      formulaCell: { r, c },
    });
  }
  return out.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
}

function buildCashBook(grid: Grid, ev: Evaluator, L: Cols, fileName: string): CashBook {
  const val = (r: number, c: number | null) => (c == null ? null : ev.value(r, c));
  const num = (r: number, c: number | null) => {
    const v = val(r, c);
    return typeof v === "number" ? v : null;
  };
  const { date, rate } = headerInfo(grid, L.firstDataRow, fileName);
  const blocks: Block[] = [];
  let block: Block | null = null;
  let section: "in" | "out" = "in";

  const newBlock = (title: string) => {
    block = { title, opening: null, incomes: [], expenses: [], closing: null, balanced: true };
    blocks.push(block);
    section = "in";
    return block;
  };
  const amountsOf = (r: number): Amounts => ({
    inUah: num(r, L.inUah), inUsd: num(r, L.inUsd), outUah: num(r, L.outUah), outUsd: num(r, L.outUsd),
  });
  const balanceOf = (r: number): Balance => {
    const a = amountsOf(r);
    return { row: r, uah: a.inUah ?? a.outUah, usd: a.inUsd ?? a.outUsd, eq: num(r, L.eq) };
  };

  for (let r = L.firstDataRow; r <= grid.maxRow; r++) {
    const label = L.labelCols.map((c) => text(val(r, c))).join(" ").replace(/\s+/g, " ").trim();
    const low = label.toLowerCase();
    const a = amountsOf(r);
    const hasAmount = a.inUah != null || a.inUsd != null || a.outUah != null || a.outUsd != null;
    const noVal = L.no != null ? val(r, L.no) : null;
    const hasNo = typeof noVal === "number";

    if (!hasNo && /залишок на поч|остаток на нач|доп\.?\s*остат|дод\.?\s*залиш/.test(low) && hasAmount) {
      const title = /доп|дод/.test(low) ? "Додатковий залишок" : blocks.length ? `Каса ${blocks.length + 1}` : "Основна каса";
      newBlock(title).opening = balanceOf(r);
      continue;
    }
    if (!hasNo && /разом|итого|всего|усього/.test(low)) continue; // own totals are rebuilt
    if (!hasNo && /расход|видат|витрат/.test(low) && !hasAmount) {
      section = "out";
      continue;
    }
    if (!hasNo && !hasAmount) continue;
    const b: Block = block ?? newBlock(blocks.length ? `Каса ${blocks.length + 1}` : "Основна каса");
    if (!hasNo && /залишок на кін|остаток на кон|^остаток$|^залишок$/.test(low)) {
      b.closing = balanceOf(r);
      block = null;
      continue;
    }

    const constFormulas: Entry["constFormulas"] = {};
    for (const [field, c] of Object.entries({ inUah: L.inUah, inUsd: L.inUsd, outUah: L.outUah, outUsd: L.outUsd })) {
      const f = c == null ? undefined : grid.cells.get(key(r, c))?.f;
      if (f && !/[A-Z]+\$?\d/i.test(f.replace(/^=/, "").replace(/E[+-]?\d/gi, "")) && !ev.overrides.has(key(r, c!)))
        constFormulas[field as keyof Amounts] = f.replace(/^=/, "");
    }
    const entry: Entry = {
      row: r,
      no: hasNo ? String(noVal) : "",
      account: tidy(text(val(r, L.accountCol))),
      party: tidy(text(val(r, L.partyCol)) || (L.accountCol == null ? label : "")),
      ...a,
      rate: num(r, L.rate),
      eq: num(r, L.eq),
      constFormulas,
    };
    const isOut = a.outUah != null || a.outUsd != null;
    const isIn = a.inUah != null || a.inUsd != null;
    if (isOut && !isIn) b.expenses.push(entry);
    else if (isIn && !isOut) b.incomes.push(entry);
    else (section === "out" ? b.expenses : b.incomes).push(entry);
  }

  for (const b of blocks) {
    const t = totals(b);
    b.balanced =
      !b.opening || !b.closing ||
      (Math.abs((b.opening.uah ?? 0) + t.inUah - t.outUah - (b.closing.uah ?? 0)) <= TOL &&
        Math.abs((b.opening.usd ?? 0) + t.inUsd - t.outUsd - (b.closing.usd ?? 0)) <= TOL);
  }

  return {
    kind: "cashbook",
    title: "Касовий звіт",
    date,
    rate,
    hasRateCol: L.rate != null && blocks.some((b) => [...b.incomes, ...b.expenses].some((e) => e.rate != null)),
    hasEqCol: L.eq != null,
    blocks,
  };
}

export function totals(b: Block) {
  const s = (xs: Entry[], f: keyof Amounts) => round2(xs.reduce((a, e) => a + (e[f] ?? 0), 0));
  const all = [...b.incomes, ...b.expenses];
  return {
    inUah: s(all, "inUah"), inUsd: s(all, "inUsd"),
    outUah: s(all, "outUah"), outUsd: s(all, "outUsd"),
    eq: round2(all.reduce((a, e) => a + (e.eq ?? 0), 0)),
  };
}

function checkBalances(doc: CashBook, L: Cols): { issues: Issue[]; checked: number } {
  const out: Issue[] = [];
  let checked = 0;
  for (const b of doc.blocks) {
    if (!b.opening || !b.closing) continue;
    const t = totals(b);
    const cur = [
      { name: "грн", sign: "₴", open: b.opening.uah, close: b.closing.uah, inn: t.inUah, outt: t.outUah, col: L.inUah },
      { name: "$", sign: "$", open: b.opening.usd, close: b.closing.usd, inn: t.inUsd, outt: t.outUsd, col: L.inUsd },
    ];
    for (const x of cur) {
      if (x.col == null || (x.close == null && x.open == null)) continue;
      checked++;
      const expected = round2((x.open ?? 0) + x.inn - x.outt);
      if (Math.abs(expected - (x.close ?? 0)) <= TOL) continue;
      out.push({
        id: `balance:${addr(b.closing.row, x.col)}`,
        kind: "balance",
        where: `${b.title}, залишок на кінець, ${x.name}`,
        text:
          `${money(x.open ?? 0)} + надходження ${money(x.inn)} − видаток ${money(x.outt)} = ${money(expected)} ${x.sign}, ` +
          `а у файлі ${money(x.close ?? 0)} ${x.sign} (різниця ${money(round2((x.close ?? 0) - expected))}).`,
        fix: `Записати ${money(expected)} ${x.sign}`,
        set: [{ r: b.closing.row, c: x.col, v: expected }],
      });
    }
  }
  return { issues: out, checked };
}

function checkConversions(doc: CashBook, L: Cols): { issues: Issue[]; checked: number } {
  const out: Issue[] = [];
  let checked = 0;
  for (const b of doc.blocks) {
    for (const list of [b.incomes, b.expenses]) {
      list.forEach((e, i) => {
        const m = e.party.match(CONV);
        if (!m) return;
        const field = e.inUsd != null ? "inUsd" : e.outUsd != null ? "outUsd" : null;
        if (!field) return;
        const col = field === "inUsd" ? L.inUsd : L.outUsd;
        if (col == null) return;
        const a = toNum(m[1]);
        const d = toNum(m[2]);
        if (!d) return;
        const expected = a / d;
        const next = list[i + 1];
        const fee = next && /^%/.test(next.party.trim()) && next[field] != null ? next[field]! : 0;
        const value = e[field]!;
        checked++;
        if (Math.abs(value + fee - expected) <= TOL) return;
        const right = round2(expected - fee);
        out.push({
          id: `conversion:${addr(e.row, col)}`,
          kind: "conversion",
          where: `Рядок ${e.row}: ${e.party}`,
          text:
            `${money(a, 0)} / ${String(d).replace(".", ",")} = ${money(expected)} $, а записано ${money(value)} $` +
            (fee ? ` + ${money(fee)} $ комісії = ${money(value + fee)} $.` : "."),
          fix: `Записати ${money(right)} $`,
          set: [{ r: e.row, c: col, v: right }],
        });
      });
    }
  }
  return { issues: out, checked };
}

function genericTable(grid: Grid, ev: Evaluator, fileName: string): Table {
  const cols: number[] = [];
  for (let c = 1; c <= grid.maxCol; c++)
    for (let r = 1; r <= grid.maxRow; r++)
      if (grid.cells.has(key(r, c))) {
        cols.push(c);
        break;
      }
  const rows: (string | number | null)[][] = [];
  for (let r = 1; r <= grid.maxRow; r++) {
    const row = cols.map((c) => {
      const v = ev.value(r, c);
      return typeof v === "string" ? tidy(v) || null : v;
    });
    if (row.some((v) => v != null && v !== "")) rows.push(row);
  }
  let header: string[] | null = null;
  const hi = rows.findIndex((r) => r.filter((v) => typeof v === "string").length >= 2);
  if (hi >= 0 && hi < 5) {
    header = rows[hi].map((v) => (v == null ? "" : String(v)));
    rows.splice(0, hi + 1);
  }
  const numericCols = cols.map((_, i) => rows.filter((r) => typeof r[i] === "number").length > rows.filter((r) => typeof r[i] === "string").length);
  const { date } = headerInfo(grid, 1, fileName);
  return { kind: "table", title: grid.sheetName || "Документ", date: date ?? parseDate(fileName), header, rows, numericCols };
}

export const docDateText = (d: Doc) => (d.date ? longDate(d.date) : "");
export const docDateShort = (d: Doc) => (d.date ? shortDate(d.date) : "");
