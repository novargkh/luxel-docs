import { LoginForm } from "./LoginForm";

export default function LoginPage() {
  return (
    <main className="mx-auto flex min-h-dvh max-w-sm flex-col justify-center px-4 py-10">
      <h1 className="text-2xl font-bold tracking-tight">Перетворювач у LUXEL-документ</h1>
      <p className="mt-1 mb-6 text-sm text-muted">Вхід лише для працівників LUXEL</p>
      <LoginForm />
    </main>
  );
}
