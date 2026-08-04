import "server-only";

import { revalidateTag, unstable_cache } from "next/cache";

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
  revalidateTag(sharePointListTag(listKey), { expire: 0 });
}

export function revalidateSharePointLists(
  listKeys: Array<SharePointListKey | string>,
): void {
  for (const listKey of listKeys) {
    revalidateSharePointList(listKey);
  }
}
