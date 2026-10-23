"use client";

import { useActionState } from "react";
import { login, type LoginState } from "./actions";

export function LoginForm() {
  const [state, action, pending] = useActionState<LoginState, FormData>(login, {});
  return (
    <form action={action} className="space-y-3">
      <label className="block text-sm font-medium" htmlFor="password">Пароль</label>
      <input
        id="password"
        name="password"
        type="password"
        autoComplete="current-password"
        required
        autoFocus
        className="w-full rounded-lg border border-line px-3.5 py-2.5 text-base outline-none focus:border-ink"
      />
      {state.error && <p className="text-sm text-brand">{state.error}</p>}
      <button
        disabled={pending}
        className="w-full rounded-lg bg-ink px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-black disabled:opacity-60"
      >
        {pending ? "Перевіряємо…" : "Увійти"}
      </button>
    </form>
  );
}
