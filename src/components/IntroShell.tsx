"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { Logo } from "./Logo";

const ASSEMBLE_MS = 1900;
const FLY_MS = 750;

/**
 * After sign-in the logo assembles on a white screen, then flies into its
 * place in the page header ([data-logo-target]) while the page fades in.
 */
export function IntroShell({ intro, children }: { intro: boolean; children: React.ReactNode }) {
  const [phase, setPhase] = useState<"assemble" | "fly" | "done">(intro ? "assemble" : "done");
  const logoRef = useRef<HTMLDivElement>(null);
  const bgRef = useRef<HTMLDivElement>(null);
  const started = useRef(false);

  const fly = useCallback(() => {
    if (started.current) return;
    started.current = true;
    const logo = logoRef.current;
    const target = document.querySelector<SVGElement>("[data-logo-target]");
    if (!logo || !target) {
      setPhase("done");
      return;
    }
    setPhase("fly");
    const from = logo.getBoundingClientRect();
    const to = target.getBoundingClientRect();
    const scale = to.width / from.width;
    const dx = to.left - from.left;
    const dy = to.top - from.top;
    const easing = "cubic-bezier(0.65, 0, 0.25, 1)";
    logo.animate(
      [{ transform: "translate(0, 0) scale(1)" }, { transform: `translate(${dx}px, ${dy}px) scale(${scale})` }],
      { duration: FLY_MS, easing, fill: "forwards" },
    );
    bgRef.current?.animate([{ opacity: 1 }, { opacity: 0 }], { duration: FLY_MS * 0.8, delay: FLY_MS * 0.2, easing, fill: "forwards" });
    setTimeout(() => setPhase("done"), FLY_MS + 30);
  }, []);

  useEffect(() => {
    if (!intro) return;
    document.cookie = "luxel_intro=; Max-Age=0; path=/";
    if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
      setPhase("done");
      return;
    }
    const t = setTimeout(fly, ASSEMBLE_MS);
    return () => clearTimeout(t);
  }, [intro, fly]);

  return (
    <div className={phase === "assemble" ? "intro-hold" : phase === "fly" ? "intro-reveal" : undefined}>
      {children}
      {phase !== "done" && (
        <div className="fixed inset-0 z-50" onClick={fly} aria-hidden>
          <div ref={bgRef} className="absolute inset-0 bg-white" />
          <div className="absolute inset-0 flex items-center justify-center px-8">
            <div ref={logoRef} className="w-[min(72vw,440px)] origin-top-left">
              <Logo animated className="block h-auto w-full" />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
