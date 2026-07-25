import {
  SHAREPOINT_LISTS,
  SHAREPOINT_SITE,
  type SharePointListKey,
} from "@/lib/schema/sharepointSchema";

function required(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

/**
 * Runtime SharePoint config. List/field names always come from the schema.
 * Only Graph IDs are supplied via environment variables.
 */
export function getSharePointConfig() {
  return {
    site: SHAREPOINT_SITE,
    siteId: getSharePointSiteId(),
    lists: SHAREPOINT_LISTS,
  };
}

/** Raw SHAREPOINT_SITE_ID from env (path form or composite GUID). */
export function getSharePointSiteId(): string {
  return required("SHAREPOINT_SITE_ID");
}

/**
 * Graph path root for this site.
 * Path-based IDs (`host:/sites/Name`) must use a trailing colon before
 * child resources (`/lists`, `/drive`), e.g.
 * `/sites/host:/sites/Name:/lists/{listId}/items`
 * Without it, Graph treats the request as a site and fails expand=fields.
 * @see https://learn.microsoft.com/en-us/graph/api/resources/sharepoint
 */
export function getSharePointSiteApiRoot(): string {
  const siteId = getSharePointSiteId().replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    const withTransition = siteId.endsWith(":") ? siteId : `${siteId}:`;
    return `/sites/${withTransition}`;
  }
  return `/sites/${siteId}`;
}

export function getSharePointListId(listKey: SharePointListKey): string {
  const envVar = SHAREPOINT_LISTS[listKey].listIdEnvVar;
  return required(envVar);
}

export function getSharePointListConfig<K extends SharePointListKey>(listKey: K) {
  return {
    ...SHAREPOINT_LISTS[listKey],
    listId: getSharePointListId(listKey),
  };
}
