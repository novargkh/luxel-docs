const MONTHS_GEN = [
  "січня", "лютого", "березня", "квітня", "травня", "червня",
  "липня", "серпня", "вересня", "жовтня", "листопада", "грудня",
];

// Month stems in Ukrainian and Russian, as written in the source files.
const MONTH_STEMS: [RegExp, number][] = [
  [/^(січ|янв)/, 0], [/^(лют|фев)/, 1], [/^(бер|мар)/, 2], [/^(кві|апр)/, 3],
  [/^(тра|мая|май)/, 4], [/^(чер|июн)/, 5], [/^(лип|июл)/, 6], [/^(сер|авг)/, 7],
  [/^(вер|сен)/, 8], [/^(жов|окт)/, 9], [/^(лис|ноя)/, 10], [/^(гру|дек)/, 11],
];

export type DateParts = { d: number; m: number; y: number };

/** "19 октября  2026 р", "19.10.2026", "19-10-26" → parts. */
export function parseDate(text: string): DateParts | null {
  const t = text.toLowerCase().replace(/\s+/g, " ");
  const word = t.match(/(\d{1,2})\s+([a-zа-яіїєґ]+)\s+(\d{4})/i);
  if (word) {
    const stem = MONTH_STEMS.find(([re]) => re.test(word[2]));
    if (stem) return { d: +word[1], m: stem[1], y: +word[3] };
  }
  const num = t.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/);
  if (num) {
    const y = num[3].length === 2 ? 2000 + +num[3] : +num[3];
    const m = +num[2] - 1;
    if (m >= 0 && m < 12 && +num[1] >= 1 && +num[1] <= 31) return { d: +num[1], m, y };
  }
  return null;
}

export const longDate = (p: DateParts) => `${p.d} ${MONTHS_GEN[p.m]} ${p.y}`;
const pad = (n: number) => String(n).padStart(2, "0");
export const shortDate = (p: DateParts) => `${pad(p.d)}.${pad(p.m + 1)}.${p.y}`;

/** Current Kyiv time as "23.10.2026 14:05". */
export function kyivNow(now = new Date()): string {
  const parts = new Intl.DateTimeFormat("uk-UA", {
    timeZone: "Europe/Kyiv",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(now);
  const get = (t: string) => parts.find((p) => p.type === t)?.value ?? "";
  return `${get("day")}.${get("month")}.${get("year")} ${get("hour")}:${get("minute")}`;
}

/** 1384845.69 → "1 384 845,69". Keep the cell from wrapping (Manrope has no usable no-break space). */
export function money(n: number | null | undefined, digits = 2): string {
  if (n == null || !Number.isFinite(n)) return "";
  const fixed = Math.abs(n).toFixed(digits);
  const [int, frac] = fixed.split(".");
  const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return (n < 0 && Number(fixed) !== 0 ? "−" : "") + grouped + (frac ? "," + frac : "");
}

export const round2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;

/** Tidies a text cell: collapses spaces and capitalises the first letter. */
export function tidy(s: string): string {
  const t = s.replace(/\s+/g, " ").trim();
  return t ? t[0].toLocaleUpperCase("uk") + t.slice(1) : t;
}
