import "server-only";

import { revalidateTag, unstable_cache } from "next/cache";
import { after } from "next/server";

import type { SharePointListKey } from "@/lib/schema/sharepointSchema";

/** Default SharePoint read TTL (seconds). Short enough for portal freshness. */
export const SHAREPOINT_CACHE_REVALIDATE_SECONDS = 45;

export function sharePointListTag(listKey: SharePointListKey | string): string {
  return `sp:${listKey}`;
}

/**
 * Cache an expensive SharePoint/Graph read across requests.
 * Key parts must uniquely identify the query (list, filter, id, etc.).
 */
export function cachedSharePointRead<T>(
  keyParts: string[],
  tags: string[],
  fn: () => Promise<T>,
  revalidateSeconds: number = SHAREPOINT_CACHE_REVALIDATE_SECONDS,
): Promise<T> {
  return unstable_cache(fn, keyParts, {
    tags,
    revalidate: revalidateSeconds,
  })();
}

/** Mark cached SharePoint list data stale after a write.
 * Use expire:0 so the next read cannot keep serving pre-delete rows
 * (profile "max" is stale-while-revalidate and caused deleted Permissions
 * to vanish briefly then reappear on refresh).
 */
export function revalidateSharePointList(
  listKey: SharePointListKey | string,
): void {
  const tag = sharePointListTag(listKey);
  try {
    revalidateTag(tag, { expire: 0 });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    // Next.js forbids revalidateTag during RSC render / cached functions.
    // Defer until after the response so incidental writes (e.g. audit logs
    // from a page) cannot crash the request.
    if (!/during render|cached functions/i.test(message)) {
      throw error;
    }
    try {
      after(() => {
        revalidateTag(tag, { expire: 0 });
      });
    } catch {
      console.warn(
        `[sharePoint] skipped revalidateTag(${tag}) outside a request`,
      );
    }
  }
}

export function revalidateSharePointLists(
  listKeys: Array<SharePointListKey | string>,
): void {
  for (const listKey of listKeys) {
    revalidateSharePointList(listKey);
  }
}
