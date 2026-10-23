"use server";

import { cookies } from "next/headers";
import { redirect } from "next/navigation";
import { COOKIE, sessionToken, sitePassword } from "@/lib/auth";

export type LoginState = { error?: string };

export async function login(_: LoginState, form: FormData): Promise<LoginState> {
  const password = sitePassword();
  if (!password) return { error: "Пароль сайту ще не налаштований (ACCESS_PASSWORD у Vercel)." };
  const given = String(form.get("password") ?? "").trim();
  if (given !== password) {
    // Slow down guessing.
    await new Promise((r) => setTimeout(r, 1200));
    return { error: "Невірний пароль" };
  }
  (await cookies()).set(COOKIE, await sessionToken(password), {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax",
    path: "/",
    maxAge: 60 * 60 * 24 * 30,
  });
  redirect("/");
}

export async function logout() {
  (await cookies()).delete(COOKIE);
  redirect("/login");
}
