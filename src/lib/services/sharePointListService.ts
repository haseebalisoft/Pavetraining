import "server-only";

import { ResponseType } from "@microsoft/microsoft-graph-client";

import {
  cachedSharePointRead,
  revalidateSharePointList,
  sharePointListTag,
} from "@/lib/cache/sharePointCache";
import {
  getSharePointListId,
  getSharePointSiteApiRoot,
} from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";
import {
  getSharePointFields,
  type SharePointListKey,
} from "@/lib/schema/sharepointSchema";

export type SharePointFields = Record<string, unknown>;

export interface SharePointListItem {
  id: string;
  fields: SharePointFields;
  /** Graph item created timestamp when available (used for document uploaded date). */
  createdDateTime?: string | null;
  /** Graph item last-modified timestamp when available. */
  lastModifiedDateTime?: string | null;
}

function escapeODataString(value: string): string {
  // OData string literals use ASCII single quotes; double them to escape.
  // Also normalise typographic apostrophes so they cannot confuse parsers.
  return value
    .replace(/[\u2018\u2019\u02BC]/g, "'")
    .replace(/'/g, "''");
}

/**
 * Microsoft Graph JS SDK concatenates $filter values without URL-encoding.
 * Unencoded `&` (e.g. "Harbour & Hill…") splits the query string and yields
 * errors like: unterminated string literal in `fields/CompanyName eq 'Harbour`.
 */
export function encodeODataFilter(filter: string): string {
  return encodeURIComponent(filter);
}

type ListQueryOptions = {
  filter?: string;
  top?: number;
  selectFields?: string[];
};

/**
 * Uncached Graph list query. Prefer getListItemsByKey for app reads.
 */
async function fetchListItemsUncached(
  listId: string,
  options?: ListQueryOptions,
): Promise<SharePointListItem[]> {
  const siteRoot = getSharePointSiteApiRoot();
  const client = getGraphClient();
  const pageSize = options?.top && options.top > 0 ? Math.min(options.top, 200) : 200;
  const hardCap = options?.top && options.top > 200 ? options.top : undefined;

  const items: SharePointListItem[] = [];
  let request = client
    .api(`${siteRoot}/lists/${listId}/items`)
    .expand(
      options?.selectFields?.length
        ? `fields($select=${options.selectFields.join(",")})`
        : "fields",
    )
    .top(pageSize)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

  if (options?.filter) {
    // GraphRequest.filter() does not encode; encode here so & and other
    // reserved URL characters inside literal values stay inside $filter.
    request = request.filter(encodeODataFilter(options.filter));
  }

  let response = (await request.get()) as {
    value?: Array<{
      id: string;
      createdDateTime?: string;
      lastModifiedDateTime?: string;
      fields?: SharePointFields;
    }>;
    "@odata.nextLink"?: string;
  };

  for (;;) {
    for (const item of response.value ?? []) {
      items.push({
        id: String(item.id),
        fields: item.fields ?? {},
        createdDateTime: item.createdDateTime ?? null,
        lastModifiedDateTime: item.lastModifiedDateTime ?? null,
      });
      if (hardCap && items.length >= hardCap) {
        return items.slice(0, hardCap);
      }
    }

    const nextLink = response["@odata.nextLink"];
    if (!nextLink) break;

    response = (await client
      .api(nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, ""))
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .get()) as typeof response;
  }

  return items;
}

async function fetchListItemByIdUncached(
  listId: string,
  itemId: string,
): Promise<SharePointListItem | null> {
  const siteRoot = getSharePointSiteApiRoot();
  const client = getGraphClient();

  try {
    const item = (await client
      .api(`${siteRoot}/lists/${listId}/items/${itemId}`)
      .expand("fields")
      .get()) as {
      id: string;
      createdDateTime?: string;
      lastModifiedDateTime?: string;
      fields?: SharePointFields;
    };

    return {
      id: String(item.id),
      fields: item.fields ?? {},
      createdDateTime: item.createdDateTime ?? null,
      lastModifiedDateTime: item.lastModifiedDateTime ?? null,
    };
  } catch (error: unknown) {
    const status =
      typeof error === "object" &&
      error !== null &&
      "statusCode" in error &&
      typeof (error as { statusCode?: unknown }).statusCode === "number"
        ? (error as { statusCode: number }).statusCode
        : undefined;

    if (status === 404) {
      return null;
    }

    throw error;
  }
}

/**
 * Low-level SharePoint list access via Microsoft Graph (uncached by list id).
 * Prefer getListItemsByKey so list-tag caching applies.
 */
export async function getListItems(
  listId: string,
  options?: ListQueryOptions,
): Promise<SharePointListItem[]> {
  return fetchListItemsUncached(listId, options);
}

export async function getListItemsByKey(
  listKey: SharePointListKey,
  options?: ListQueryOptions,
): Promise<SharePointListItem[]> {
  const listId = getSharePointListId(listKey);
  const filter = options?.filter ?? "";
  const top = String(options?.top ?? "");
  const select = options?.selectFields?.join(",") ?? "";

  return cachedSharePointRead(
    ["sp-list-items", listKey, listId, filter, top, select],
    [sharePointListTag(listKey)],
    () => fetchListItemsUncached(listId, options),
  );
}

export async function getListItemById(
  listId: string,
  itemId: string,
): Promise<SharePointListItem | null> {
  return fetchListItemByIdUncached(listId, itemId);
}

export async function getListItemByKey(
  listKey: SharePointListKey,
  itemId: string,
): Promise<SharePointListItem | null> {
  const listId = getSharePointListId(listKey);
  const trimmedId = itemId.trim();
  if (!trimmedId) return null;

  return cachedSharePointRead(
    ["sp-list-item", listKey, listId, trimmedId],
    [sharePointListTag(listKey)],
    () => fetchListItemByIdUncached(listId, trimmedId),
  );
}

export function buildFieldEqualsFilter(
  fieldName: string,
  value: string,
): string {
  return `fields/${fieldName} eq '${escapeODataString(value)}'`;
}

/** Numeric LookupId equality — preferred over display-name text filters. */
export function buildFieldLookupIdEqualsFilter(
  lookupIdFieldName: string,
  lookupId: string | number,
): string {
  const id = Number(lookupId);
  if (!Number.isFinite(id)) {
    throw new Error(`Invalid lookup id for filter field ${lookupIdFieldName}.`);
  }
  return `fields/${lookupIdFieldName} eq ${id}`;
}

export function buildSchemaFieldEqualsFilter(
  listKey: SharePointListKey,
  fieldKey: string,
  value: string,
): string {
  const fields = getSharePointFields(listKey) as Record<string, string>;
  const internalName = fields[fieldKey];
  if (!internalName) {
    throw new Error(
      `Unknown schema field "${fieldKey}" for list "${listKey}".`,
    );
  }
  return buildFieldEqualsFilter(internalName, value);
}

export function asBoolean(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    return value.toLowerCase() === "true" || value === "1" || value === "Yes";
  }
  if (typeof value === "number") return value === 1;
  return false;
}

export function asString(value: unknown): string | undefined {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (typeof value === "number") return String(value);
  return undefined;
}

export function asNullableString(value: unknown): string | null {
  return asString(value) ?? null;
}

/** Resolves plain strings or SharePoint lookup objects to display text. */
export function asLookupOrString(value: unknown): string | null {
  const direct = asString(value);
  if (direct) {
    return direct;
  }

  if (Array.isArray(value)) {
    for (const entry of value) {
      const resolved = asLookupOrString(entry);
      if (resolved) return resolved;
    }
    return null;
  }

  if (value && typeof value === "object") {
    const record = value as {
      LookupValue?: unknown;
      lookupValue?: unknown;
      DisplayName?: unknown;
      Title?: unknown;
    };
    return (
      asString(record.LookupValue) ??
      asString(record.lookupValue) ??
      asString(record.DisplayName) ??
      asString(record.Title) ??
      null
    );
  }

  return null;
}

/**
 * Graph often returns only `{Field}LookupId` without LookupValue.
 * Reads nested LookupId or the sibling `{fieldInternalName}LookupId` field.
 */
export function extractLookupId(
  fields: SharePointFields | Record<string, unknown>,
  fieldInternalName: string,
): string | null {
  const direct = fields[fieldInternalName];
  if (direct && typeof direct === "object" && "LookupId" in direct) {
    const id = (direct as { LookupId?: unknown }).LookupId;
    if (typeof id === "number" || typeof id === "string") {
      const text = String(id).trim();
      if (text) return text;
    }
  }

  const lookupIdField = fields[`${fieldInternalName}LookupId`];
  if (typeof lookupIdField === "number" || typeof lookupIdField === "string") {
    const text = String(lookupIdField).trim();
    return text || null;
  }

  return null;
}

/**
 * Downloads file binary for a SharePoint list item that has an associated drive item.
 */
export async function getListItemFileContent(
  listKey: SharePointListKey,
  itemId: string,
): Promise<{ content: ArrayBuffer; contentType: string; fileName: string | null } | null> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId(listKey);
  const client = getGraphClient();

  try {
    const driveItem = (await client
      .api(`${siteRoot}/lists/${listId}/items/${itemId}/driveItem`)
      .get()) as { name?: string; file?: { mimeType?: string } };

    const content = (await client
      .api(`${siteRoot}/lists/${listId}/items/${itemId}/driveItem/content`)
      .responseType(ResponseType.ARRAYBUFFER)
      .get()) as ArrayBuffer;

    return {
      content,
      contentType: driveItem.file?.mimeType ?? "application/octet-stream",
      fileName: asNullableString(driveItem.name),
    };
  } catch (error: unknown) {
    const status =
      typeof error === "object" &&
      error !== null &&
      "statusCode" in error &&
      typeof (error as { statusCode?: unknown }).statusCode === "number"
        ? (error as { statusCode: number }).statusCode
        : undefined;

    if (status === 404) {
      return null;
    }

    throw error;
  }
}

/**
 * Creates a SharePoint list item. `fields` must use SharePoint internal names.
 */
export async function createListItemByKey(
  listKey: SharePointListKey,
  fields: SharePointFields,
): Promise<SharePointListItem> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId(listKey);
  const client = getGraphClient();

  const created = (await client
    .api(`${siteRoot}/lists/${listId}/items`)
    .post({ fields })) as {
    id: string;
    createdDateTime?: string;
    fields?: SharePointFields;
  };

  revalidateSharePointList(listKey);

  return {
    id: String(created.id),
    fields: created.fields ?? fields,
    createdDateTime: created.createdDateTime ?? null,
  };
}

/**
 * Updates fields on an existing SharePoint list item.
 */
export async function updateListItemFieldsByKey(
  listKey: SharePointListKey,
  itemId: string,
  fields: SharePointFields,
): Promise<SharePointListItem> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId(listKey);
  const client = getGraphClient();

  await client
    .api(`${siteRoot}/lists/${listId}/items/${itemId}/fields`)
    .patch(fields);

  // Bypass read cache so the caller gets post-write fields immediately.
  const refreshed = await fetchListItemByIdUncached(listId, itemId);
  revalidateSharePointList(listKey);

  if (!refreshed) {
    throw new Error(`Updated item ${itemId} could not be reloaded.`);
  }
  return refreshed;
}

/**
 * Deletes a SharePoint list item via Graph.
 */
export async function deleteListItemByKey(
  listKey: SharePointListKey,
  itemId: string,
): Promise<void> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId(listKey);
  const client = getGraphClient();

  await client.api(`${siteRoot}/lists/${listId}/items/${itemId}`).delete();
  revalidateSharePointList(listKey);
}

/**
 * Returns internal column names for a SharePoint list (Graph columns API).
 * Used for schema health checks (e.g. EventCompany present on Events).
 */
export async function getListColumnNames(
  listKey: SharePointListKey,
): Promise<string[]> {
  const siteRoot = getSharePointSiteApiRoot();
  const listId = getSharePointListId(listKey);

  return cachedSharePointRead(
    ["sp-list-columns", listKey, listId],
    [sharePointListTag(listKey)],
    async () => {
      const client = getGraphClient();
      const response = (await client
        .api(`${siteRoot}/lists/${listId}/columns`)
        .top(200)
        .get()) as { value?: Array<{ name?: string }> };

      return (response.value ?? [])
        .map((column) => column.name?.trim())
        .filter((name): name is string => Boolean(name));
    },
    300,
  );
}

export async function listHasColumn(
  listKey: SharePointListKey,
  columnName: string,
): Promise<boolean> {
  const columns = await getListColumnNames(listKey);
  const target = columnName.trim().toLowerCase();
  return columns.some((name) => name.toLowerCase() === target);
}

/** Maps schema field keys to SharePoint internal names for write payloads. */
export function toSharePointFields(
  listKey: SharePointListKey,
  values: Record<string, unknown>,
): SharePointFields {
  const schema = getSharePointFields(listKey) as Record<string, string>;
  const payload: SharePointFields = {};

  for (const [key, value] of Object.entries(values)) {
    if (value === undefined) {
      continue;
    }
    const internal = schema[key];
    if (!internal || internal === "ID") {
      continue;
    }
    payload[internal] = value;
  }

  return payload;
}
