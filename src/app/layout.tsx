import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Перетворювач у LUXEL-документ",
  description: "Касовий звіт з Excel у фірмовий XLSX і PDF",
  robots: { index: false, follow: false },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="uk">
      <body className="min-h-dvh">{children}</body>
    </html>
  );
}
