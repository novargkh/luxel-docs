import { colNumber, key, type Grid } from "./grid";

/** Thrown for formulas we can't evaluate (other sheets, unknown functions). */
export class Unsupported extends Error {}

type Tok =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "ref"; r: number; c: number }
  | { t: "range"; r1: number; c1: number; r2: number; c2: number }
  | { t: "fn"; v: string }
  | { t: "op"; v: string };

const REF = /^\$?([A-Z]{1,3})\$?(\d+)/i;

function tokenize(src: string): Tok[] {
  const s = src.replace(/^=/, "");
  const out: Tok[] = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === " ") {
      i++;
      continue;
    }
    if (s[i] === "'" || s.slice(i).match(/^[A-Z0-9_]+!/i)) throw new Unsupported("sheet ref");
    const num = s.slice(i).match(/^\d+(\.\d+)?(E[+-]?\d+)?/i);
    if (num && !s.slice(i).match(REF)) {
      out.push({ t: "num", v: parseFloat(num[0]) });
      i += num[0].length;
      continue;
    }
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      if (end < 0) throw new Unsupported("string");
      out.push({ t: "str", v: s.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    const fn = s.slice(i).match(/^([A-Z][A-Z0-9.]*)\s*\(/i);
    if (fn) {
      out.push({ t: "fn", v: fn[1].toUpperCase() });
      i += fn[0].length;
      continue;
    }
    const ref = s.slice(i).match(REF);
    if (ref) {
      const r1 = parseInt(ref[2], 10);
      const c1 = colNumber(ref[1]);
      i += ref[0].length;
      if (s[i] === ":") {
        const ref2 = s.slice(i + 1).match(REF);
        if (!ref2) throw new Unsupported("range");
        i += 1 + ref2[0].length;
        const r2 = parseInt(ref2[2], 10);
        const c2 = colNumber(ref2[1]);
        out.push({
          t: "range",
          r1: Math.min(r1, r2),
          c1: Math.min(c1, c2),
          r2: Math.max(r1, r2),
          c2: Math.max(c1, c2),
        });
      } else {
        out.push({ t: "ref", r: r1, c: c1 });
      }
      continue;
    }
    const op = s.slice(i).match(/^(<>|<=|>=|[-+*/^&(),;<>=%])/);
    if (op) {
      out.push({ t: "op", v: op[1] === ";" ? "," : op[1] });
      i += op[1].length;
      continue;
    }
    throw new Unsupported(`token at ${i}`);
  }
  return out;
}

type Val = number | string | Val[];

/**
 * Evaluates formulas of a grid. `value(r,c)` returns the current value of a
 * cell, recomputing formula cells (memoised), so dependent formulas see
 * recomputed inputs rather than Excel's cached results.
 */
export class Evaluator {
  private memo = new Map<string, number | string | null>();
  private stack = new Set<string>();
  /** Formula cells whose cached value must be kept as is (user's choice). */
  keepCached = new Set<string>();
  /** Values set by the user's fixes; they win over the file. */
  overrides = new Map<string, number>();

  constructor(private grid: Grid) {}

  reset() {
    this.memo.clear();
  }

  value(r: number, c: number): number | string | null {
    const k = key(r, c);
    if (this.overrides.has(k)) return this.overrides.get(k)!;
    const cell = this.grid.cells.get(k);
    if (!cell) return null;
    if (!cell.f || this.keepCached.has(k)) return cell.v;
    if (this.memo.has(k)) return this.memo.get(k)!;
    if (this.stack.has(k)) throw new Unsupported("cycle");
    this.stack.add(k);
    try {
      const v = this.evaluate(cell.f);
      this.memo.set(k, v);
      return v;
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
      this.memo.set(k, cell.v);
      return cell.v;
    } finally {
      this.stack.delete(k);
    }
  }

  /** Evaluates a formula; throws Unsupported if it can't. */
  evaluate(formula: string): number | string {
    const toks = tokenize(formula);
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v: string) => peek()?.t === "op" && (peek() as { v: string }).v === v;

    const num = (v: Val): number => {
      if (Array.isArray(v)) throw new Unsupported("range in scalar");
      if (typeof v === "number") return v;
      if (v === "") return 0;
      const n = Number(String(v).replace(",", "."));
      if (Number.isNaN(n)) throw new Unsupported("text in arithmetic");
      return n;
    };

    const primary = (): Val => {
      const t = toks[p++];
      if (!t) throw new Unsupported("eof");
      if (t.t === "num") return t.v;
      if (t.t === "str") return t.v;
      if (t.t === "ref") return this.value(t.r, t.c) ?? 0;
      if (t.t === "range") {
        const vals: Val[] = [];
        for (let r = t.r1; r <= t.r2; r++)
          for (let c = t.c1; c <= t.c2; c++) {
            const v = this.value(r, c);
            if (v != null) vals.push(v);
          }
        return vals;
      }
      if (t.t === "fn") {
        const args: Val[] = [];
        if (!isOp(")")) {
          args.push(compare());
          while (isOp(",")) {
            p++;
            args.push(compare());
          }
        }
        if (!isOp(")")) throw new Unsupported("paren");
        p++;
        return call(t.v, args);
      }
      if (t.v === "(") {
        const v = compare();
        if (!isOp(")")) throw new Unsupported("paren");
        p++;
        return v;
      }
      if (t.v === "-") return -num(power());
      if (t.v === "+") return num(power());
      throw new Unsupported("unexpected");
    };

    const postfix = (): Val => {
      let v = primary();
      while (isOp("%")) {
        p++;
        v = num(v) / 100;
      }
      return v;
    };
    const power = (): Val => {
      let v = postfix();
      while (isOp("^")) {
        p++;
        v = Math.pow(num(v), num(postfix()));
      }
      return v;
    };
    const term = (): Val => {
      let v = power();
      while (isOp("*") || isOp("/")) {
        const op = (toks[p++] as { v: string }).v;
        const rhs = num(power());
        if (op === "/" && rhs === 0) throw new Unsupported("div0");
        v = op === "*" ? num(v) * rhs : num(v) / rhs;
      }
      return v;
    };
    const sum = (): Val => {
      let v = term();
      while (isOp("+") || isOp("-")) {
        const op = (toks[p++] as { v: string }).v;
        const rhs = num(term());
        v = op === "+" ? num(v) + rhs : num(v) - rhs;
      }
      return v;
    };
    const concat = (): Val => {
      let v = sum();
      while (isOp("&")) {
        p++;
        v = String(v) + String(sum());
      }
      return v;
    };
    const compare = (): Val => {
      const v = concat();
      for (const op of ["=", "<>", "<", ">", "<=", ">="]) {
        if (isOp(op)) {
          p++;
          const a = num(v);
          const b = num(concat());
          const res =
            op === "=" ? a === b : op === "<>" ? a !== b : op === "<" ? a < b : op === ">" ? a > b : op === "<=" ? a <= b : a >= b;
          return res ? 1 : 0;
        }
      }
      return v;
    };

    const flat = (args: Val[]): number[] =>
      args.flatMap((a) => (Array.isArray(a) ? a : [a])).filter((x): x is number => typeof x === "number");

    const call = (fn: string, args: Val[]): Val => {
      switch (fn) {
        case "SUM":
          return flat(args).reduce((a, b) => a + b, 0);
        case "MIN":
          return Math.min(...flat(args));
        case "MAX":
          return Math.max(...flat(args));
        case "AVERAGE": {
          const xs = flat(args);
          if (!xs.length) throw new Unsupported("avg empty");
          return xs.reduce((a, b) => a + b, 0) / xs.length;
        }
        case "ABS":
          return Math.abs(num(args[0]));
        case "ROUND":
        case "ROUNDUP":
        case "ROUNDDOWN": {
          const d = args[1] == null ? 0 : num(args[1]);
          const m = Math.pow(10, d);
          const x = num(args[0]) * m;
          const r = fn === "ROUND" ? Math.round(x) : fn === "ROUNDUP" ? Math.sign(x) * Math.ceil(Math.abs(x)) : Math.trunc(x);
          return r / m;
        }
        case "IF":
          return num(args[0]) ? (args[1] ?? 1) : (args[2] ?? 0);
        default:
          throw new Unsupported(fn);
      }
    };

    const v = compare();
    if (p !== toks.length) throw new Unsupported("trailing");
    if (Array.isArray(v)) throw new Unsupported("range result");
    return v;
  }
}

/** Cells referenced by a formula (only plain refs and small ranges). */
export function references(formula: string): { r: number; c: number }[] {
  try {
    const out: { r: number; c: number }[] = [];
    for (const t of tokenize(formula)) {
      if (t.t === "ref") out.push({ r: t.r, c: t.c });
      if (t.t === "range") for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) out.push({ r, c });
    }
    return out;
  } catch {
    return [];
  }
}
