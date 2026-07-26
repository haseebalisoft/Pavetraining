import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import {
  SHAREPOINT_LISTS,
  type SharePointListKey,
} from "../schema/sharepointSchema";
import type { ListRow } from "../types/models";

export interface SpListClient {
  spHttpClient: SPHttpClient;
  webUrl: string;
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Build `/_api/web/lists/...` root. Prefer GUID when set (titles with "/"
 * like "Offers / Promotions" 404 if used raw in getbytitle).
 */
function listApiRoot(
  client: SpListClient,
  listKey: SharePointListKey
): string {
  const list = SHAREPOINT_LISTS[listKey] as {
    listName: string;
    listId?: string;
  };
  if (list.listId) {
    return (
      client.webUrl + "/_api/web/lists(guid'" + list.listId + "')"
    );
  }
  // Encode so spaces and "/" don't break the path
  return (
    client.webUrl +
    "/_api/web/lists/getbytitle('" +
    encodeURIComponent(escapeOData(list.listName)) +
    "')"
  );
}

function restHeaders(
  extra?: Record<string, string>
): Record<string, string> {
  const headers: Record<string, string> = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
    "odata-version": "",
  };
  if (extra) {
    for (const key in extra) {
      if (Object.prototype.hasOwnProperty.call(extra, key)) {
        headers[key] = extra[key];
      }
    }
  }
  return headers;
}

function writeHeaders(
  extra?: Record<string, string>
): Record<string, string> {
  // Use plain application/json for writes — SPO can throw
  // JsonReaderException when Content-Type includes odata=nometadata.
  const headers: Record<string, string> = {
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json",
    "odata-version": "",
  };
  if (extra) {
    for (const key in extra) {
      if (Object.prototype.hasOwnProperty.call(extra, key)) {
        headers[key] = extra[key];
      }
    }
  }
  return headers;
}

/**
 * Fields that must not be written as plain strings via REST (complex types).
 */
const NON_WRITABLE_FIELDS: { [name: string]: boolean } = {
  CompanyLogo: true,
  Attachments: true,
  FileLeafRef: true,
  FileRef: true,
  FileDirRef: true,
  AuthorId: true,
  EditorId: true,
  ContentTypeId: true,
  ContentType: true,
  GUID: true,
  UniqueId: true,
  Created: true,
  Modified: true,
  ID: true,
  Id: true,
  CompanyLookupId: true,
};

/**
 * Drop empty / complex values and ensure Title is present for list creates.
 */
export function sanitizeWriteFields(
  fields: Record<string, unknown>,
  options?: { forCreate?: boolean }
): Record<string, unknown> {
  const clean: Record<string, unknown> = {};
  for (const key in fields) {
    if (!Object.prototype.hasOwnProperty.call(fields, key)) continue;
    // CompanyLogo is writable only as a stringified thumbnail JSON blob
    if (key === "CompanyLogo") {
      const v = fields[key];
      if (typeof v === "string" && v.trim().charAt(0) === "{") {
        clean[key] = v;
      }
      continue;
    }
    if (NON_WRITABLE_FIELDS[key]) continue;
    if (key.indexOf(".") >= 0) continue;
    if (key.indexOf("@odata.") === 0) continue;
    if (key === "__metadata") continue;

    const value = fields[key];
    if (value === undefined) continue;
    if (value === null) {
      if (!options || !options.forCreate) {
        clean[key] = null;
      }
      continue;
    }
    if (typeof value === "string" && value.trim() === "") continue;
    if (typeof value === "object" && !Array.isArray(value)) {
      // Skip nested objects (lookups/thumbnails) — use *Id fields instead
      continue;
    }
    clean[key] = value;
  }

  if (
    options &&
    options.forCreate &&
    (clean.Title === undefined || clean.Title === null || clean.Title === "")
  ) {
    const name = clean.CompanyName || clean.CandidateName || clean.UserEmail;
    if (typeof name === "string" && name.trim()) {
      clean.Title = name.trim();
    }
  }

  return clean;
}

function mapItem(item: Record<string, unknown>, fallbackId?: string): ListRow {
  const id = String(
    item.Id != null
      ? item.Id
      : item.ID != null
        ? item.ID
        : fallbackId != null
          ? fallbackId
          : ""
  );
  const fields: Record<string, unknown> = {};
  for (const key in item) {
    if (
      Object.prototype.hasOwnProperty.call(item, key) &&
      key !== "@odata.etag"
    ) {
      fields[key] = item[key];
    }
  }
  return { id: id, fields: fields };
}

async function getJson(
  client: SpListClient,
  url: string
): Promise<Record<string, unknown>> {
  const response: SPHttpClientResponse = await client.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: restHeaders() }
  );
  if (!response.ok) {
    const text = await response.text();
    throw new Error("SharePoint REST " + response.status + ": " + text);
  }
  return (await response.json()) as Record<string, unknown>;
}

/**
 * Reads list items via SharePoint REST (current user context).
 * Follows nextLink pages so Admin gets full list data, not only the first page.
 */
export async function getListItems(
  client: SpListClient,
  listKey: SharePointListKey,
  options?: {
    filter?: string;
    top?: number;
    select?: string[];
    expand?: string[];
    /** Soft cap across pages (default 5000). */
    maxItems?: number;
  }
): Promise<ListRow[]> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const top = options && options.top ? options.top : 5000;
  const maxItems =
    options && options.maxItems ? options.maxItems : 20000;
  const params: string[] = ["$top=" + top];
  if (options && options.filter) {
    params.push("$filter=" + encodeURIComponent(options.filter));
  }
  if (options && options.select && options.select.length > 0) {
    params.push(
      "$select=" + options.select.map(encodeURIComponent).join(",")
    );
  }
  if (options && options.expand && options.expand.length > 0) {
    params.push(
      "$expand=" + options.expand.map(encodeURIComponent).join(",")
    );
  }

  let url = listApiRoot(client, listKey) + "/items?" + params.join("&");

  const rows: ListRow[] = [];
  while (url && rows.length < maxItems) {
    let json: Record<string, unknown>;
    try {
      json = await getJson(client, url);
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      throw new Error('Failed to read "' + listName + '": ' + message);
    }

    const items = (json.value as Array<Record<string, unknown>>) || [];
    for (let i = 0; i < items.length; i++) {
      rows.push(mapItem(items[i]));
      if (rows.length >= maxItems) {
        break;
      }
    }

    const next =
      (json["@odata.nextLink"] as string | undefined) ||
      (json["odata.nextLink"] as string | undefined);
    url = next || "";
  }

  return rows;
}

export async function getListItem(
  client: SpListClient,
  listKey: SharePointListKey,
  itemId: string
): Promise<ListRow | null> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const url =
    listApiRoot(client, listKey) +
    "/items(" +
    encodeURIComponent(itemId) +
    ")";

  const response = await client.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    { headers: restHeaders() }
  );

  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      "Failed to read item " +
        itemId +
        ' from "' +
        listName +
        '" (' +
        response.status +
        "): " +
        text
    );
  }

  const item = (await response.json()) as Record<string, unknown>;
  return mapItem(item, itemId);
}

export async function createListItem(
  client: SpListClient,
  listKey: SharePointListKey,
  fields: Record<string, unknown>
): Promise<ListRow> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const clean = sanitizeWriteFields(fields, { forCreate: true });
  if (Object.keys(clean).length === 0) {
    throw new Error(
      'Nothing to create in "' +
        listName +
        '". Fill at least Title or Company Name / Candidate Name.'
    );
  }

  const bodyText = JSON.stringify(clean);
  if (!bodyText || bodyText.charAt(0) !== "{") {
    throw new Error("Invalid create payload (not a JSON object).");
  }

  const url = listApiRoot(client, listKey) + "/items";

  const response = await client.spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: writeHeaders(),
      body: bodyText,
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      'Failed to create item in "' +
        listName +
        '" (' +
        response.status +
        "): " +
        text
    );
  }

  const item = (await response.json()) as Record<string, unknown>;
  return {
    id: String(item.Id != null ? item.Id : item.ID != null ? item.ID : ""),
    fields: clean,
  };
}

export async function updateListItem(
  client: SpListClient,
  listKey: SharePointListKey,
  itemId: string,
  fields: Record<string, unknown>
): Promise<void> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const clean = sanitizeWriteFields(fields, { forCreate: false });
  if (Object.keys(clean).length === 0) {
    throw new Error("Nothing to update — no writable field values provided.");
  }

  const bodyText = JSON.stringify(clean);
  if (!bodyText || bodyText.charAt(0) !== "{") {
    throw new Error("Invalid update payload (not a JSON object).");
  }

  const url =
    listApiRoot(client, listKey) +
    "/items(" +
    encodeURIComponent(itemId) +
    ")";

  const response = await client.spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: writeHeaders({
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
      }),
      body: bodyText,
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      "Failed to update item " +
        itemId +
        ' in "' +
        listName +
        '" (' +
        response.status +
        "): " +
        text
    );
  }
}

/**
 * Deletes a list item via SharePoint REST DELETE.
 */
export async function deleteListItem(
  client: SpListClient,
  listKey: SharePointListKey,
  itemId: string
): Promise<void> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const url =
    listApiRoot(client, listKey) +
    "/items(" +
    encodeURIComponent(itemId) +
    ")";

  const response = await client.spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: restHeaders({
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
      }),
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      "Failed to delete item " +
        itemId +
        ' in "' +
        listName +
        '" (' +
        response.status +
        "): " +
        text
    );
  }
}

export function asString(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "string") return value.trim() || null;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (Array.isArray(value)) {
    const parts: string[] = [];
    for (let i = 0; i < value.length; i++) {
      const part = asString(value[i]);
      if (part) parts.push(part);
    }
    return parts.length ? parts.join(", ") : null;
  }
  if (typeof value === "object" && value !== null) {
    const record = value as Record<string, unknown>;
    if (typeof record.Title === "string") return record.Title.trim() || null;
    if (typeof record.LookupValue === "string") {
      return record.LookupValue.trim() || null;
    }
    if (typeof record.CompanyName === "string") {
      return record.CompanyName.trim() || null;
    }
    if (typeof record.Email === "string") return record.Email.trim() || null;
  }
  return null;
}

export function asBoolean(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  if (typeof value === "string") {
    const n = value.trim().toLowerCase();
    return n === "true" || n === "yes" || n === "1";
  }
  return false;
}

export function fieldEqualsFilter(
  internalName: string,
  value: string
): string {
  return internalName + " eq '" + escapeOData(value) + "'";
}

/**
 * Turns SharePoint loginName / UPN into a plain email when possible.
 */
export function normalizeSharePointUserEmail(raw: string): string {
  let value = (raw || "").trim();
  if (!value) return "";

  const membership = value.match(/\|membership\|([^|]+)$/i);
  if (membership && membership[1]) {
    value = membership[1];
  } else {
    const pipe = value.lastIndexOf("|");
    if (pipe >= 0 && value.indexOf("@") > pipe) {
      value = value.substring(pipe + 1);
    }
  }

  const hash = value.lastIndexOf("#");
  if (hash >= 0 && value.indexOf("@") > hash) {
    value = value.substring(hash + 1);
  }

  return value.trim().toLowerCase();
}
