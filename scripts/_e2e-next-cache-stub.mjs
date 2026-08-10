/**
 * Test-harness implementation of "next/cache" for the live E2E driver.
 *
 * Next's real unstable_cache needs a request-scoped incremental cache that does
 * not exist outside a Next server. Rather than disabling caching (a pure
 * passthrough made every SharePoint read hit Graph and turned a 50-row matrix
 * commit into hundreds of full list reads), this mirrors what production does:
 *
 *   - entries keyed by `keyParts` + call arguments
 *   - entries expire after `revalidate` seconds (the app passes 45)
 *   - the in-flight promise is cached, so concurrent callers share one read
 *     (same dedupe behaviour as production)
 *   - revalidateTag(tag) drops every entry carrying that tag — how the app
 *     invalidates a list after a write
 *
 * So E2E timing is representative of production, and a post-write read still
 * sees fresh data because the write path calls revalidateSharePointList().
 *
 * Production code is untouched; the swap happens in scripts/_e2e-hook.mjs.
 */

/** key -> { value: Promise, expiresAt: number, tags: string[] } */
const store = new Map();

const DEFAULT_REVALIDATE_SECONDS = 45;

export function unstable_cache(fn, keyParts = [], options = {}) {
  const baseKey = JSON.stringify(keyParts ?? []);
  const tags = Array.isArray(options.tags) ? options.tags : [];
  const ttlMs =
    (typeof options.revalidate === "number"
      ? options.revalidate
      : DEFAULT_REVALIDATE_SECONDS) * 1000;

  return async (...args) => {
    const key = `${baseKey}|${JSON.stringify(args)}`;
    const now = Date.now();
    const hit = store.get(key);
    if (hit && hit.expiresAt > now) {
      return hit.value;
    }

    const value = fn(...args);
    store.set(key, { value, expiresAt: now + ttlMs, tags });
    try {
      return await value;
    } catch (error) {
      // A failed read must never be served to the next caller.
      store.delete(key);
      throw error;
    }
  };
}

export function revalidateTag(tag) {
  for (const [key, entry] of store) {
    if (entry.tags.includes(tag)) store.delete(key);
  }
}

export function revalidatePath() {}
export function unstable_noStore() {}
export function unstable_expirePath() {}
export function unstable_expireTag(tag) {
  revalidateTag(tag);
}

/** Harness diagnostics only. */
export function __cacheSize() {
  return store.size;
}
