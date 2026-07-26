import { SPHttpClient } from "@microsoft/sp-http";

import {
  SHAREPOINT_LISTS,
  type SharePointListKey,
} from "../schema/sharepointSchema";
import {
  type SpListClient,
  updateListItem,
} from "./sharePointListService";

export interface ThumbnailMeta {
  type: "thumbnail";
  fileName: string;
  fieldName: string;
  serverUrl: string;
  serverRelativeUrl: string;
  id?: string;
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}

function siteOrigin(webUrl: string): string {
  try {
    return new URL(webUrl).origin;
  } catch {
    const i = webUrl.indexOf("/", webUrl.indexOf("//") + 2);
    return i > 0 ? webUrl.substring(0, i) : webUrl;
  }
}

/**
 * Parse SharePoint Image/Thumbnail column value (object or JSON string).
 */
export function parseThumbnailField(
  value: unknown
): { serverRelativeUrl?: string; fileName?: string; id?: string } | null {
  if (value === null || value === undefined || value === "") return null;

  let parsed: unknown = value;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (trimmed.charAt(0) === "{") {
      try {
        parsed = JSON.parse(trimmed);
      } catch {
        return { serverRelativeUrl: trimmed };
      }
    } else if (trimmed.charAt(0) === "/") {
      return { serverRelativeUrl: trimmed };
    } else {
      return null;
    }
  }

  if (typeof parsed === "object" && parsed !== null) {
    const rec = parsed as Record<string, unknown>;
    return {
      serverRelativeUrl:
        typeof rec.serverRelativeUrl === "string"
          ? rec.serverRelativeUrl
          : undefined,
      fileName: typeof rec.fileName === "string" ? rec.fileName : undefined,
      id: typeof rec.id === "string" ? rec.id : undefined,
    };
  }
  return null;
}

export function thumbnailPreviewUrl(
  webUrl: string,
  value: unknown
): string | null {
  const meta = parseThumbnailField(value);
  if (!meta || !meta.serverRelativeUrl) return null;
  return siteOrigin(webUrl) + meta.serverRelativeUrl;
}

async function getListId(
  client: SpListClient,
  listKey: SharePointListKey
): Promise<string> {
  const listName = SHAREPOINT_LISTS[listKey].listName;
  const url =
    client.webUrl +
    "/_api/web/lists/getbytitle('" +
    escapeOData(listName) +
    "')?$select=Id";

  const response = await client.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "odata-version": "",
      },
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      'Failed to resolve list id for "' + listName + '": ' + text
    );
  }

  const json = (await response.json()) as { Id?: string };
  if (!json.Id) {
    throw new Error('List id missing for "' + listName + '".');
  }
  return String(json.Id);
}

function safeImageFileName(file: File): string {
  const raw = (file.name || "logo.png").replace(/[^\w.\-]+/g, "_");
  const parts = raw.split(".");
  const ext =
    parts.length > 1 ? parts[parts.length - 1].toLowerCase() : "png";
  const base =
    parts.length > 1 ? parts.slice(0, -1).join(".") : parts[0] || "logo";
  return Date.now() + "-" + base.substring(0, 40) + "." + ext;
}

/**
 * Upload an image into the list's Site Assets folder via SharePoint UploadImage,
 * then set the Thumbnail/Image column (CompanyLogo) on the item.
 */
export async function uploadAndSetListImage(
  client: SpListClient,
  listKey: SharePointListKey,
  itemId: string,
  fieldInternalName: string,
  file: File
): Promise<ThumbnailMeta> {
  if (!file || !file.size) {
    throw new Error("Please choose an image file to upload.");
  }
  if (file.size > 10 * 1024 * 1024) {
    throw new Error("Image must be 10 MB or smaller.");
  }

  const listName = SHAREPOINT_LISTS[listKey].listName;
  const listId = await getListId(client, listKey);
  const imageName = safeImageFileName(file);

  const uploadUrl =
    client.webUrl +
    "/_api/web/uploadimage(listTitle=@a1,imageName=@a2,listId=@a3,itemId=@a4)" +
    "?@a1='" +
    escapeOData(listName) +
    "'&@a2='" +
    escapeOData(imageName) +
    "'&@a3='" +
    listId +
    "'&@a4=" +
    itemId;

  const buffer = await file.arrayBuffer();

  const response = await client.spHttpClient.post(
    uploadUrl,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/octet-stream",
        "odata-version": "",
      },
      body: buffer,
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      "Image upload failed (" + response.status + "): " + text
    );
  }

  const uploaded = (await response.json()) as {
    Name?: string;
    ServerRelativeUrl?: string;
    UniqueId?: string;
  };

  if (!uploaded.ServerRelativeUrl) {
    throw new Error("Image upload succeeded but no ServerRelativeUrl returned.");
  }

  const meta: ThumbnailMeta = {
    type: "thumbnail",
    fileName: uploaded.Name || imageName,
    fieldName: fieldInternalName,
    serverUrl: siteOrigin(client.webUrl),
    serverRelativeUrl: uploaded.ServerRelativeUrl,
    id: uploaded.UniqueId,
  };

  // Image columns must be a JSON *string*, not a nested object.
  await updateListItem(client, listKey, itemId, {
    [fieldInternalName]: JSON.stringify(meta),
  });

  return meta;
}

/** Allow CompanyLogo when it is already a stringified thumbnail JSON. */
export function isThumbnailJsonString(value: unknown): boolean {
  return typeof value === "string" && value.trim().charAt(0) === "{";
}
