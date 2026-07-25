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
 * Reads list items via SharePoint REST (current user context).
 * No Next.js / external API — direct list access inside SharePoint.
 */
export async function getListItems(
  client: SpListClient,
  listKey: SharePointListKey,
  options?: { filter?: string; top?: number; select?: string[] }
): Promise<ListRow[]> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const top = options?.top ?? 500;
  const params: string[] = [`$top=${top}`];
  if (options?.filter) {
    params.push(`$filter=${encodeURIComponent(options.filter)}`);
  }
  if (options?.select && options.select.length > 0) {
    params.push(`$select=${options.select.map(encodeURIComponent).join(",")}`);
  }

  const url = `${client.webUrl}/_api/web/lists/getbytitle('${escapeOData(
    listName
  )}')/items?${params.join("&")}`;

  const response: SPHttpClientResponse = await client.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
      },
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      `Failed to read "${listName}" (${response.status}): ${text}`
    );
  }

  const json = (await response.json()) as {
    value?: Array<Record<string, unknown>>;
  };
  const items = json.value ?? [];
  return items.map((item) => {
    const id = String(item.Id ?? item.ID ?? "");
    const fields: Record<string, unknown> = { ...item };
    delete fields["@odata.etag"];
    return { id, fields };
  });
}

export function asString(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "string") return value.trim() || null;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (typeof value === "object" && value !== null) {
    const record = value as Record<string, unknown>;
    if (typeof record.Title === "string") return record.Title.trim() || null;
    if (typeof record.LookupValue === "string") {
      return record.LookupValue.trim() || null;
    }
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
  return `${internalName} eq '${escapeOData(value)}'`;
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}
