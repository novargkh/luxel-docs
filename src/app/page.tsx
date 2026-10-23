import { Brand } from "@/components/Brand";
import { Converter } from "@/components/Converter";
import { logout } from "./login/actions";

export default function Home() {
  return (
    <main className="mx-auto flex min-h-dvh max-w-2xl flex-col px-4 py-8 sm:py-12">
      <header className="flex items-center justify-between">
        <Brand />
        <form action={logout}>
          <button className="text-sm text-muted hover:text-ink">Вийти</button>
        </form>
      </header>

      <h1 className="mt-10 text-3xl font-bold tracking-tight sm:text-4xl">Перетворювач у LUXEL-документ</h1>
      <p className="mt-2 text-muted">
        Завантажте касовий звіт з Excel. Система перевірить підрахунки та підготує фірмові XLSX і PDF.
      </p>
      <div className="mt-4 mb-8 h-[3px] w-10 rounded-full bg-brand" />

      <Converter />

      <footer className="mt-auto pt-12">
        <div className="flex gap-3 rounded-xl bg-soft px-4 py-3 text-xs leading-relaxed text-muted">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="mt-0.5 shrink-0" aria-hidden>
            <rect x="4" y="11" width="16" height="10" rx="2" />
            <path d="M8 11V7a4 4 0 0 1 8 0v4" />
          </svg>
          <p>
            Системі заборонено зберігати фінансові документи. Файл обробляється лише у вашому браузері й нікуди не
            надсилається; копії не зберігаються ні на сервері, ні в хмарі. Після закриття сторінки повторно завантажити
            готові файли неможливо, тож збережіть їх одразу.
          </p>
        </div>
      </footer>
    </main>
  );
}
