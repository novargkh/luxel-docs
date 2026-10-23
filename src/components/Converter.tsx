"use client";

import { useCallback, useRef, useState } from "react";
import type { Grid } from "@/lib/grid";
import { analyze, docDateText, type Analysis, type Choices, type Issue } from "@/lib/cashbook";
import { kyivNow, money } from "@/lib/format";
import { fileBase, summary } from "@/lib/layout";

type Loaded = { name: string; grid: Grid; createdAt: string };

const MAX_SIZE = 15 * 1024 * 1024;

async function asBase64(url: string): Promise<string> {
  const buf = await (await fetch(url)).arrayBuffer();
  let bin = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

function save(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

export function Converter() {
  const [loaded, setLoaded] = useState<Loaded | null>(null);
  const [choices, setChoices] = useState<Choices>({ fixed: new Set(), kept: new Set() });
  const [result, setResult] = useState<Analysis | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [drag, setDrag] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  const reset = () => {
    setLoaded(null);
    setResult(null);
    setError(null);
    setChoices({ fixed: new Set(), kept: new Set() });
    if (input.current) input.current.value = "";
  };

  const open = useCallback(async (file: File) => {
    setError(null);
    setResult(null);
    if (/\.xls$/i.test(file.name)) {
      setError("Це старий формат .xls. Відкрийте файл в Excel і збережіть як «Книга Excel (.xlsx)».");
      return;
    }
    if (!/\.xlsx$/i.test(file.name)) {
      setError("Потрібен файл Excel у форматі .xlsx");
      return;
    }
    if (file.size > MAX_SIZE) {
      setError("Файл завеликий (понад 15 МБ)");
      return;
    }
    setBusy("Читаємо файл…");
    try {
      const ExcelJS = (await import("exceljs")).default;
      const { mainSheet, readGrid } = await import("@/lib/grid");
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await file.arrayBuffer());
      const grid = readGrid(mainSheet(wb));
      if (!grid.cells.size) throw new Error("Аркуш порожній");
      const fresh: Choices = { fixed: new Set(), kept: new Set() };
      setLoaded({ name: file.name, grid, createdAt: kyivNow() });
      setChoices(fresh);
      setResult(analyze(grid, file.name, fresh));
    } catch (e) {
      console.error(e);
      setError(`Не вдалося прочитати файл: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setBusy(null);
    }
  }, []);

  const decide = (issues: Issue[], how: "fix" | "keep") => {
    if (!loaded) return;
    const next: Choices = { fixed: new Set(choices.fixed), kept: new Set(choices.kept) };
    for (const i of issues) (how === "fix" ? next.fixed : next.kept).add(i.id);
    setChoices(next);
    setResult(analyze(loaded.grid, loaded.name, next));
  };

  const download = async (kind: "xlsx" | "pdf") => {
    if (!result || !loaded) return;
    setBusy(kind === "pdf" ? "Готуємо PDF…" : "Готуємо XLSX…");
    try {
      const logo = await asBase64("/logo.png");
      const name = fileBase(result.doc);
      if (kind === "xlsx") {
        const { buildXlsx } = await import("@/lib/buildXlsx");
        const buf = await buildXlsx(result.doc, { logoBase64: `data:image/png;base64,${logo}`, createdAt: loaded.createdAt });
        save(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${name}.xlsx`);
      } else {
        const [{ buildPdfDefinition, PDF_FONTS, PDF_FONT_FILES }, pdfMod] = await Promise.all([
          import("@/lib/buildPdf"),
          import("pdfmake/build/pdfmake"),
        ]);
        const pdfMake = (pdfMod as unknown as { default: typeof pdfMod }).default ?? pdfMod;
        const vfs: Record<string, string> = {};
        await Promise.all(PDF_FONT_FILES.map(async (f) => (vfs[f] = await asBase64(`/fonts/${f}`))));
        const pm = pdfMake as unknown as {
          addVirtualFileSystem(v: Record<string, string>): void;
          addFonts(f: object): void;
          createPdf(d: object): { getBlob(): Promise<Blob> };
        };
        pm.addVirtualFileSystem(vfs);
        pm.addFonts(PDF_FONTS);
        const def = buildPdfDefinition(result.doc, { logoDataUrl: `data:image/png;base64,${logo}`, createdAt: loaded.createdAt });
        save(await pm.createPdf(def).getBlob(), `${name}.pdf`);
      }
    } catch (e) {
      console.error(e);
      setError(`Не вдалося створити файл: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setBusy(null);
    }
  };

  const ready = result && result.issues.length === 0;

  return (
    <div className="space-y-5">
      {!loaded && (
        <button
          type="button"
          onClick={() => input.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setDrag(true);
          }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDrag(false);
            const f = e.dataTransfer.files?.[0];
            if (f) open(f);
          }}
          className={`flex w-full flex-col items-center justify-center rounded-2xl border-2 border-dashed px-6 py-14 text-center transition ${
            drag ? "border-brand bg-brand-soft" : "border-line bg-soft hover:border-ink/30"
          }`}
        >
          <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" className="text-brand" aria-hidden>
            <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
            <path d="M14 3v5h5M12 18v-6M9 15l3-3 3 3" />
          </svg>
          <span className="mt-4 text-base font-semibold">Перетягніть файл Excel сюди</span>
          <span className="mt-1 text-sm text-muted">або натисніть, щоб вибрати (.xlsx)</span>
        </button>
      )}
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) open(f);
        }}
      />

      {busy && <p className="text-sm text-muted">{busy}</p>}
      {error && <p className="rounded-xl bg-brand-soft px-4 py-3 text-sm text-brand">{error}</p>}

      {loaded && result && (
        <div className="space-y-5">
          <div className="flex items-start justify-between gap-4 rounded-2xl border border-line px-5 py-4">
            <div className="min-w-0">
              <div className="truncate text-sm font-semibold">{loaded.name}</div>
              <div className="mt-0.5 text-sm text-muted">
                {result.doc.title}
                {docDateText(result.doc) ? ` · ${docDateText(result.doc)}` : ""}
              </div>
            </div>
            <button onClick={reset} className="shrink-0 text-sm font-medium text-muted underline-offset-4 hover:text-ink hover:underline">
              Інший файл
            </button>
          </div>

          <CheckReport result={result} onDecide={decide} />

          {result.doc.kind === "cashbook" && <SummaryPreview result={result} />}

          <div>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <button
                disabled={!ready || !!busy}
                onClick={() => download("xlsx")}
                className="rounded-xl bg-ink px-5 py-3.5 text-sm font-semibold text-white transition hover:bg-black disabled:cursor-not-allowed disabled:opacity-35"
              >
                Завантажити XLSX
              </button>
              <button
                disabled={!ready || !!busy}
                onClick={() => download("pdf")}
                className="rounded-xl bg-brand px-5 py-3.5 text-sm font-semibold text-white transition hover:brightness-95 disabled:cursor-not-allowed disabled:opacity-35"
              >
                Завантажити PDF
              </button>
            </div>
            {!ready && <p className="mt-2 text-center text-xs text-muted">Спочатку оберіть, що робити з помилками вище</p>}
          </div>
        </div>
      )}
    </div>
  );
}

function CheckReport({ result, onDecide }: { result: Analysis; onDecide: (i: Issue[], how: "fix" | "keep") => void }) {
  const { issues, kept, fixed, stats } = result;
  const checked = [
    stats.formulas ? `формул: ${stats.formulas}` : "",
    stats.conversions ? `конвертацій: ${stats.conversions}` : "",
    stats.balances ? `залишків: ${stats.balances}` : "",
  ].filter(Boolean);

  if (!issues.length) {
    const clean = !kept.length && !fixed.length;
    return (
      <div className="rounded-2xl bg-ok-soft px-5 py-4">
        <div className="flex items-center gap-2 font-semibold text-ok">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" aria-hidden>
            <path d="M20 6 9 17l-5-5" />
          </svg>
          {clean ? "Математичних помилок у підрахунках немає" : "Усі помилки опрацьовано"}
        </div>
        {checked.length > 0 && <p className="mt-1 text-sm text-ink/70">Перевірено {checked.join(", ")}.</p>}
        {fixed.length > 0 && <p className="mt-1 text-sm text-ink/70">Виправлено: {fixed.length}.</p>}
        {kept.length > 0 && <p className="mt-1 text-sm text-ink/70">Залишено як у файлі: {kept.length}.</p>}
      </div>
    );
  }

  return (
    <div className="rounded-2xl border border-brand/30 bg-brand-soft/50 px-5 py-4">
      <div className="font-semibold text-brand">
        Знайдено помилки в підрахунках: {issues.length}
      </div>
      <p className="mt-0.5 text-sm text-ink/70">Виправити їх у документі чи залишити як у файлі?</p>
      <ul className="mt-4 space-y-3">
        {issues.map((i) => (
          <li key={i.id} className="rounded-xl bg-white px-4 py-3 shadow-[0_0_0_1px_var(--color-line)]">
            <div className="text-xs font-semibold tracking-wide text-muted uppercase">{i.where}</div>
            <div className="mt-1 text-sm">{i.text}</div>
            <div className="mt-3 flex flex-wrap gap-2">
              <button onClick={() => onDecide([i], "fix")} className="rounded-lg bg-ink px-3 py-1.5 text-xs font-semibold text-white hover:bg-black">
                Виправити: {i.fix.replace(/^Записати /, "")}
              </button>
              <button onClick={() => onDecide([i], "keep")} className="rounded-lg border border-line bg-white px-3 py-1.5 text-xs font-semibold hover:border-ink/40">
                Залишити як є
              </button>
            </div>
          </li>
        ))}
      </ul>
      {issues.length > 1 && (
        <div className="mt-4 flex flex-wrap gap-2">
          <button onClick={() => onDecide(issues, "fix")} className="text-sm font-semibold text-ink underline underline-offset-4">
            Виправити все
          </button>
          <span className="text-muted">·</span>
          <button onClick={() => onDecide(issues, "keep")} className="text-sm font-semibold text-ink underline underline-offset-4">
            Залишити все як є
          </button>
        </div>
      )}
    </div>
  );
}

function SummaryPreview({ result }: { result: Analysis }) {
  if (result.doc.kind !== "cashbook") return null;
  return (
    <div className="space-y-4">
      {result.doc.blocks.map((b, bi) => (
        <div key={bi}>
          <div className="mb-2 text-sm font-semibold">{b.title}</div>
          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
            {summary(b).map((s, i) => (
              <div key={i} className={`rounded-xl px-3.5 py-3 ${i === 3 ? "bg-brand-soft" : "bg-soft"}`}>
                <div className={`text-[10px] font-semibold tracking-wide uppercase ${i === 3 ? "text-brand" : "text-muted"}`}>{s.label}</div>
                <div className="tabular mt-1.5 text-sm font-semibold">{s.uah == null ? "—" : `${money(s.uah)} ₴`}</div>
                <div className="tabular text-xs text-muted">{s.usd == null ? "" : `${money(s.usd)} $`}</div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
