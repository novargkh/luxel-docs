/* eslint-disable @next/next/no-img-element */
"use client";

import { useEffect } from "react";
import { docDateShort, docDateText, totals, type CashBook, type Doc, type Table } from "@/lib/cashbook";
import { money } from "@/lib/format";
import { columns, lines, summary, type ColKey } from "@/lib/layout";

type Props = {
  doc: Doc;
  createdAt: string;
  canDownload: boolean;
  busy: boolean;
  onDownload: (kind: "xlsx" | "pdf") => void;
  onClose: () => void;
};

/** Full-screen HTML view of the document, styled like the PDF. */
export function DocPreview({ doc, createdAt, canDownload, busy, onDownload, onClose }: Props) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    document.addEventListener("keydown", onKey);
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.removeEventListener("keydown", onKey);
      document.body.style.overflow = overflow;
    };
  }, [onClose]);

  const sub = [docDateText(doc), doc.kind === "cashbook" && doc.rate ? `курс ${String(doc.rate).replace(".", ",")}` : ""]
    .filter(Boolean)
    .join("  ·  ");

  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-[#ececec]" role="dialog" aria-modal="true" aria-label="Перегляд документа">
      <div className="flex items-center gap-2 border-b border-line bg-white px-4 py-2.5">
        <button onClick={onClose} className="rounded-lg px-2.5 py-1.5 text-sm font-medium text-muted hover:bg-soft hover:text-ink">
          ← Назад
        </button>
        <div className="min-w-0 flex-1 truncate text-sm font-semibold">
          {doc.title}
          {docDateShort(doc) ? ` · ${docDateShort(doc)}` : ""}
        </div>
        <button
          disabled={!canDownload || busy}
          onClick={() => onDownload("xlsx")}
          className="rounded-lg bg-ink px-3 py-1.5 text-xs font-semibold text-white hover:bg-black disabled:opacity-35"
        >
          XLSX
        </button>
        <button
          disabled={!canDownload || busy}
          onClick={() => onDownload("pdf")}
          className="rounded-lg bg-brand px-3 py-1.5 text-xs font-semibold text-white hover:brightness-95 disabled:opacity-35"
        >
          PDF
        </button>
      </div>
      {!canDownload && (
        <div className="bg-brand-soft px-4 py-1.5 text-center text-xs text-brand">
          Є невирішені помилки в підрахунках — завантаження відкриється після рішення по них
        </div>
      )}

      <div className="flex-1 overflow-auto px-3 py-6 sm:px-6">
        <article className="mx-auto w-full max-w-[900px] bg-white px-5 py-6 text-ink shadow-[0_1px_3px_rgba(0,0,0,0.12)] sm:px-10 sm:py-9">
          <header className="flex items-center justify-between gap-4">
            <img src="/logo.png" alt="LUXEL" className="h-[18px] w-auto" />
            <span className="text-[11px] text-muted">
              {doc.title}
              {docDateShort(doc) ? ` · ${docDateShort(doc)}` : ""}
            </span>
          </header>

          <h1 className="mt-8 text-[28px] font-extrabold tracking-tight">{doc.title}</h1>
          {sub && <p className="mt-0.5 text-sm text-muted">{sub}</p>}
          <div className="mt-3 mb-7 h-[3px] w-10 bg-brand" />

          {doc.kind === "cashbook" ? <CashBookView doc={doc} /> : <TableView doc={doc} />}

          <footer className="mt-10 flex items-center justify-between gap-3 border-t border-line pt-3 text-[11px] text-muted">
            <span>Сформовано: {createdAt}</span>
            <span className="font-extrabold tracking-[0.2em] text-brand">LUXEL</span>
          </footer>
        </article>
      </div>
    </div>
  );
}

const fmt = (v: number | null | undefined) => (v == null ? "" : money(v));

function CashBookView({ doc }: { doc: CashBook }) {
  const cols = columns(doc);
  const n = cols.length;
  const num = (k: ColKey) => cols.find((c) => c.key === k)?.numeric;
  const cell = (k: ColKey, v: string, extra = "") => (
    <td key={k} className={`px-2 py-1.5 ${num(k) ? "tabular text-right" : ""} ${k === "no" ? "text-muted" : ""} ${extra}`}>
      {v}
    </td>
  );
  const labelRow = (key: string, label: string, vals: Partial<Record<ColKey, string>>, cls: string) => (
    <tr key={key} className={`border-b border-line font-bold ${cls}`}>
      <td colSpan={3} className="px-2 py-1.5">{label}</td>
      {cols.slice(3).map((c) => cell(c.key, vals[c.key] ?? ""))}
    </tr>
  );

  return (
    <>
      {doc.blocks.map((b, bi) => (
        <section key={bi} className="mb-5">
          <div className="mb-2 text-[13px] font-semibold">{b.title}</div>
          <div className="grid grid-cols-2 gap-1 sm:grid-cols-4">
            {summary(b).map((s, i) => (
              <div key={i} className={`px-3 py-2.5 ${i === 3 ? "bg-brand-soft" : "bg-soft"}`}>
                <div className={`text-[9px] tracking-wide uppercase ${i === 3 ? "text-brand" : "text-muted"}`}>{s.label}</div>
                <div className={`tabular mt-1 text-[14px] ${i === 3 ? "font-bold" : "font-semibold"}`}>{s.uah == null ? "—" : `${money(s.uah)} ₴`}</div>
                <div className={`tabular text-[11px] ${i === 3 ? "" : "text-muted"}`}>{s.usd == null ? "" : `${money(s.usd)} $`}</div>
              </div>
            ))}
          </div>
        </section>
      ))}

      <div className="mt-6 overflow-x-auto">
        <table className="w-full min-w-[760px] border-collapse text-[11.5px]">
          <thead className="bg-soft text-[10.5px] font-bold">
            <tr>
              {cols.map((c, i) =>
                c.group ? (
                  cols[i - 1]?.group === c.group ? null : (
                    <th key={c.key} colSpan={2} className="px-2 pt-2 pb-1 text-center">{c.group}</th>
                  )
                ) : (
                  <th key={c.key} rowSpan={2} className={`px-2 py-2 align-top ${c.numeric ? "text-right" : "text-left"}`}>{c.title}</th>
                ),
              )}
            </tr>
            <tr className="border-b border-[#bdbdbd]">
              {cols.filter((c) => c.group).map((c) => (
                <th key={c.key} className="px-2 pb-2 text-right">{c.title}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {lines(doc).map((line, li) => {
              if (line.type === "block")
                return (
                  <tr key={li}>
                    <td colSpan={n} className="px-2 pt-5 pb-1.5 text-[10.5px] font-bold tracking-[0.08em] text-brand uppercase">{line.label}</td>
                  </tr>
                );
              if (line.type === "sub")
                return (
                  <tr key={li} className="border-b border-line">
                    <td colSpan={n} className="px-2 pt-2.5 pb-1 text-[10.5px] font-bold text-muted">{line.label}</td>
                  </tr>
                );
              if (line.type === "opening") {
                const o = line.block.opening!;
                return labelRow(`o${li}`, line.label, { inUah: fmt(o.uah), inUsd: fmt(o.usd) }, "bg-soft");
              }
              if (line.type === "turnover") {
                const t = totals(line.block);
                return labelRow(`t${li}`, line.label, {
                  inUah: fmt(t.inUah), inUsd: fmt(t.inUsd), outUah: fmt(t.outUah), outUsd: fmt(t.outUsd), eq: fmt(t.eq),
                }, "border-t border-t-[#bdbdbd]");
              }
              if (line.type === "closing") {
                const c = line.block.closing!;
                return labelRow(`c${li}`, line.label, { inUah: fmt(c.uah), inUsd: fmt(c.usd), eq: fmt(c.eq) }, "bg-brand-soft");
              }
              const e = line.entry;
              const vals: Record<ColKey, string> = {
                no: e.no, account: e.account, party: e.party,
                inUah: fmt(e.inUah), inUsd: fmt(e.inUsd), outUah: fmt(e.outUah), outUsd: fmt(e.outUsd),
                rate: e.rate == null ? "" : String(e.rate).replace(".", ","), eq: fmt(e.eq),
              };
              return (
                <tr key={li} className="border-b border-line hover:bg-soft/60">
                  {cols.map((c) => cell(c.key, vals[c.key]))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </>
  );
}

function TableView({ doc }: { doc: Table }) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-[11.5px]">
        {doc.header && (
          <thead className="bg-soft text-[10.5px] font-bold">
            <tr className="border-b border-[#bdbdbd]">
              {doc.header.map((h, i) => (
                <th key={i} className={`px-2 py-2 ${doc.numericCols[i] ? "text-right" : "text-left"}`}>{h}</th>
              ))}
            </tr>
          </thead>
        )}
        <tbody>
          {doc.rows.map((r, ri) => (
            <tr key={ri} className="border-b border-line">
              {r.map((v, i) => (
                <td key={i} className={`px-2 py-1.5 ${typeof v === "number" ? "tabular text-right" : ""}`}>
                  {typeof v === "number" ? (Number.isInteger(v) ? money(v, 0) : money(v)) : (v ?? "")}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
