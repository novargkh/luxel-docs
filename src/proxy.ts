import { NextResponse, type NextRequest } from "next/server";
import { COOKIE, sessionToken, sitePassword } from "@/lib/auth";

export async function proxy(req: NextRequest) {
  const password = sitePassword();
  const cookie = req.cookies.get(COOKIE)?.value;
  if (password && cookie && cookie === (await sessionToken(password))) return NextResponse.next();
  return NextResponse.redirect(new URL("/login", req.nextUrl.origin));
}

export const config = {
  // Everything except the login page and the public assets it needs.
  matcher: ["/((?!login|_next/static|_next/image|fonts/|logo.png|favicon.ico|icon.png).*)"],
};
