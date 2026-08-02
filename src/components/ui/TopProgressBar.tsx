"use client";

import { usePathname, useSearchParams } from "next/navigation";
import { Suspense, useEffect, useRef, useState } from "react";

import styles from "./TopProgressBar.module.css";

type Phase = "idle" | "loading" | "done";

/**
 * App-wide navigation progress bar. It starts the instant an internal link is
 * clicked (capture-phase listener, before Next.js begins the route transition)
 * and completes when the new route commits (detected via a pathname change).
 * This gives immediate "something is happening" feedback so navigation never
 * feels frozen while the server renders a `force-dynamic` page.
 *
 * State transitions happen in the click handler and during render (React's
 * "adjust state when a prop changes" pattern) — never synchronously inside an
 * effect — so there are no cascading-render lint violations.
 */
export function TopProgressBar() {
  return (
    // useSearchParams requires a Suspense boundary; fallback is nothing since
    // the bar is a transient overlay.
    <Suspense fallback={null}>
      <TopProgressBarInner />
    </Suspense>
  );
}

function TopProgressBarInner() {
  const pathname = usePathname();
  const searchParams = useSearchParams();
  // Key changes whenever the committed URL changes (path OR query).
  const locationKey = `${pathname}?${searchParams.toString()}`;
  const [phase, setPhase] = useState<Phase>("idle");
  const [seenKey, setSeenKey] = useState(locationKey);
  const safetyRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Complete the bar as soon as the committed URL changes (navigation done).
  if (locationKey !== seenKey) {
    setSeenKey(locationKey);
    setPhase((current) => (current === "loading" ? "done" : "idle"));
  }

  useEffect(() => {
    function clearSafety() {
      if (safetyRef.current) {
        clearTimeout(safetyRef.current);
        safetyRef.current = null;
      }
    }

    function onClick(event: MouseEvent) {
      if (
        event.defaultPrevented ||
        event.button !== 0 ||
        event.metaKey ||
        event.ctrlKey ||
        event.shiftKey ||
        event.altKey
      ) {
        return;
      }

      const target = event.target as HTMLElement | null;
      const anchor = target?.closest("a");
      if (!anchor) return;

      const href = anchor.getAttribute("href");
      if (!href) return;
      if (anchor.target && anchor.target !== "_self") return;
      if (anchor.hasAttribute("download")) return;

      let url: URL;
      try {
        url = new URL(anchor.href, window.location.href);
      } catch {
        return;
      }
      if (url.origin !== window.location.origin) return;
      // Same page (hash-only or identical URL) — no route change, no bar.
      if (
        url.pathname === window.location.pathname &&
        url.search === window.location.search
      ) {
        return;
      }

      setPhase("loading");
      // Safety net: if navigation stalls/errors, finish after 15s.
      clearSafety();
      safetyRef.current = setTimeout(() => {
        setPhase((current) => (current === "loading" ? "done" : current));
      }, 15000);
    }

    document.addEventListener("click", onClick, true);
    return () => {
      document.removeEventListener("click", onClick, true);
      clearSafety();
    };
  }, []);

  const className =
    phase === "loading"
      ? `${styles.bar} ${styles.loading}`
      : phase === "done"
        ? `${styles.bar} ${styles.done}`
        : styles.bar;

  return (
    <div
      className={className}
      aria-hidden
      onTransitionEnd={(event) => {
        if (event.propertyName === "opacity" && phase === "done") {
          setPhase("idle");
        }
      }}
    />
  );
}
