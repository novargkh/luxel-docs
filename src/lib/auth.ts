export const COOKIE = "luxel_docs";

/**
 * Cookie value for a signed-in browser: HMAC of a fixed label keyed by the
 * site password, so changing ACCESS_PASSWORD signs everyone out.
 */
export async function sessionToken(password: string): Promise<string> {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey("raw", enc.encode(password), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode("luxel-docs-session-v1"));
  return Array.from(new Uint8Array(sig), (b) => b.toString(16).padStart(2, "0")).join("");
}

export function sitePassword(): string | null {
  const p = process.env.ACCESS_PASSWORD?.trim();
  return p ? p : null;
}
