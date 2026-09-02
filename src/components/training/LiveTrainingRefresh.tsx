"use client";

import { useEffect, useRef } from "react";
import { useRouter } from "next/navigation";

const MIN_REFRESH_MS = 2000;

/**
 * Refetch RSC data when the tab becomes visible again so Workforce profile,
 * Candidate profile, and Training summary pick up register/matrix writes
 * without a manual refresh.
 */
export function LiveTrainingRefresh(): null {
  const router = useRouter();
  const lastRefreshAt = useRef(0);

  useEffect(() => {
    function refreshIfStale() {
      const now = Date.now();
      if (now - lastRefreshAt.current < MIN_REFRESH_MS) return;
      lastRefreshAt.current = now;
      router.refresh();
    }

    function onVisible() {
      if (document.visibilityState === "visible") refreshIfStale();
    }

    window.addEventListener("focus", refreshIfStale);
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      window.removeEventListener("focus", refreshIfStale);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, [router]);

  return null;
}
