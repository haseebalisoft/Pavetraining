import "server-only";

import { ResponseType } from "@microsoft/microsoft-graph-client";

import { getSharePointListId, getSharePointSiteApiRoot } from "@/lib/config/sharepoint";
import { getGraphClient } from "@/lib/graph/graphClient";
import {
  SHAREPOINT_SITE,
  type SharePointListKey,
} from "@/lib/schema/sharepointSchema";
import {
  getListItemByKey,
  updateListItemFieldsByKey,
} from "@/lib/services/sharePointListService";
import { ValidationError } from "@/lib/services/errorHandler";

export interface ThumbnailMeta {
  type: "thumbnail";
  fileName: string;
  fieldName: string;
  serverUrl: string;
  serverRelativeUrl: string;
  id?: string;
  driveId?: string;
}

const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

function siteOrigin(webUrl: string): string {
  try {
    return new URL(webUrl).origin;
  } catch {
    const i = webUrl.indexOf("/", webUrl.indexOf("//") + 2);
    return i > 0 ? webUrl.substring(0, i) : webUrl;
  }
}

export function parseThumbnailField(
  value: unknown,
): {
  serverRelativeUrl?: string;
  fileName?: string;
  id?: string;
  driveId?: string;
} | null {
  if (value === null || value === undefined || value === "") return null;

  let parsed: unknown = value;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (trimmed.startsWith("{")) {
      try {
        parsed = JSON.parse(trimmed);
      } catch {
        return trimmed.startsWith("/")
          ? { serverRelativeUrl: trimmed }
          : null;
      }
    } else if (trimmed.startsWith("/")) {
      return { serverRelativeUrl: trimmed };
    } else if (trimmed.startsWith("http")) {
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
      driveId: typeof rec.driveId === "string" ? rec.driveId : undefined,
    };
  }
  return null;
}

/** Absolute SharePoint URL for a thumbnail field (may require auth in browser). */
export function thumbnailAbsoluteUrl(value: unknown): string | null {
  const meta = parseThumbnailField(value);
  if (!meta?.serverRelativeUrl) return null;
  if (meta.serverRelativeUrl.startsWith("http")) return meta.serverRelativeUrl;
  return siteOrigin(SHAREPOINT_SITE.url) + meta.serverRelativeUrl;
}

function safeImageFileName(fileName: string): string {
  const raw = (fileName || "image.png").replace(/[^\w.\-]+/g, "_");
  const parts = raw.split(".");
  const ext =
    parts.length > 1 ? parts[parts.length - 1]!.toLowerCase() : "png";
  const base =
    parts.length > 1 ? parts.slice(0, -1).join(".") : parts[0] || "image";
  return `${Date.now()}-${base.substring(0, 40)}.${ext}`;
}

async function getDefaultDriveId(): Promise<string> {
  const siteRoot = getSharePointSiteApiRoot();
  const client = getGraphClient();
  const drive = (await client.api(`${siteRoot}/drive`).get()) as { id?: string };
  if (!drive.id) throw new Error("Could not resolve SharePoint default drive.");
  return drive.id;
}

/**
 * Upload an image via Graph into Shared Documents / PortalMedia, then set the
 * Thumbnail/Image column to stringified thumbnail JSON (SharePoint Image column).
 */
export async function uploadAndSetListImage(options: {
  listKey: SharePointListKey;
  itemId: string;
  fieldInternalName: string;
  fileName: string;
  bytes: Uint8Array;
  contentType?: string | null;
}): Promise<{ meta: ThumbnailMeta; previewUrl: string }> {
  const { listKey, itemId, fieldInternalName, fileName, bytes, contentType } =
    options;

  if (!bytes.length) {
    throw new ValidationError("Please choose an image file to upload.");
  }
  if (bytes.length > MAX_IMAGE_BYTES) {
    throw new ValidationError("Image must be 10 MB or smaller.");
  }

  const existing = await getListItemByKey(listKey, itemId);
  if (!existing) {
    throw new ValidationError("Record not found.");
  }

  const imageName = safeImageFileName(fileName);
  const folder = `PortalMedia/${listKey}-${itemId}`;
  const driveId = await getDefaultDriveId();
  const client = getGraphClient();
  const encodedPath = `${folder}/${imageName}`
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");

  const uploaded = (await client
    .api(`/drives/${driveId}/root:/${encodedPath}:/content`)
    .header("Content-Type", contentType || "application/octet-stream")
    .put(Buffer.from(bytes))) as {
    id?: string;
    name?: string;
    webUrl?: string;
  };

  if (!uploaded.id || !uploaded.webUrl) {
    throw new Error("Image upload succeeded but Graph returned no file id/url.");
  }

  const serverRelativeUrl = decodeURIComponent(
    uploaded.webUrl.replace(siteOrigin(SHAREPOINT_SITE.url), ""),
  );

  const meta: ThumbnailMeta = {
    type: "thumbnail",
    fileName: uploaded.name || imageName,
    fieldName: fieldInternalName,
    serverUrl: siteOrigin(SHAREPOINT_SITE.url),
    serverRelativeUrl,
    id: uploaded.id,
    driveId,
  };

  await updateListItemFieldsByKey(listKey, itemId, {
    [fieldInternalName]: JSON.stringify(meta),
  });

  // Touch list id env so callers can revalidate — updateListItemFieldsByKey already does.
  void getSharePointListId(listKey);

  return {
    meta,
    previewUrl: meta.serverUrl + meta.serverRelativeUrl,
  };
}

function sharedDocumentsRelativePath(serverRelativeUrl: string): string | null {
  const decoded = decodeURIComponent(serverRelativeUrl);
  const markers = ["Shared Documents/", "Shared%20Documents/"];
  for (const marker of markers) {
    const idx = decoded.indexOf(marker.replace("%20", " "));
    const idxEncoded = decoded.indexOf("Shared%20Documents/");
    if (idx >= 0) {
      return decoded.slice(idx + "Shared Documents/".length);
    }
    if (idxEncoded >= 0) {
      return decodeURIComponent(
        decoded.slice(idxEncoded + "Shared%20Documents/".length),
      );
    }
  }
  return null;
}

/** Stream image bytes for a thumbnail field via Graph (app-only). */
export async function fetchThumbnailContent(
  value: unknown,
): Promise<{ bytes: ArrayBuffer; contentType: string } | null> {
  const meta = parseThumbnailField(value);
  if (!meta) return null;

  const client = getGraphClient();

  try {
    if (meta.id) {
      const driveId = meta.driveId || (await getDefaultDriveId());
      const bytes = (await client
        .api(`/drives/${driveId}/items/${meta.id}/content`)
        .responseType(ResponseType.ARRAYBUFFER)
        .get()) as ArrayBuffer;
      return { bytes, contentType: "image/jpeg" };
    }

    if (meta.serverRelativeUrl) {
      const relative = sharedDocumentsRelativePath(meta.serverRelativeUrl);
      if (relative) {
        const driveId = await getDefaultDriveId();
        const encoded = relative
          .split("/")
          .filter(Boolean)
          .map((part) => encodeURIComponent(part))
          .join("/");
        const bytes = (await client
          .api(`/drives/${driveId}/root:/${encoded}:/content`)
          .responseType(ResponseType.ARRAYBUFFER)
          .get()) as ArrayBuffer;
        return { bytes, contentType: "image/jpeg" };
      }
    }
  } catch {
    return null;
  }

  return null;
}

export { MAX_IMAGE_BYTES };
