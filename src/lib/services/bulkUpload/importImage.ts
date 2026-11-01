import "server-only";

import type { SharePointListKey } from "@/lib/schema/sharepointSchema";
import {
  MAX_IMAGE_BYTES,
  uploadAndSetListImage,
} from "@/lib/services/listThumbnailService";

/**
 * Image cells in bulk uploads (Company Logo, Candidate Photo). Shared by the
 * Company and Workforce importers so both read cells the same way and store
 * images through the same path as the Admin upload buttons
 * (uploadAndSetListImage → Shared Documents/PortalMedia/<list>-<id>/ + the
 * list's Image column).
 */

const PLACEHOLDER =
  /^((logo|photo|picture|image)(\s+placeholder)?|placeholder|n\/?a|none|null|-|—|–)$/i;

const IMAGE_FILE_NAME = /^[^/\\:*?"<>|]+\.(png|jpe?g|gif|webp|bmp|svg)$/i;

export type ImportImagePlan =
  /** Blank or a placeholder ("Logo placeholder", "N/A", …). */
  | { kind: "missing"; placeholder: string | null }
  | { kind: "invalid"; reason: string }
  | { kind: "url"; url: string; fileName: string }
  /** Image file name — the admin attaches the file on the upload page. */
  | { kind: "file"; fileName: string };

export function planImportImage(
  value: string | null | undefined,
): ImportImagePlan {
  const raw = (value ?? "").trim();
  if (!raw) return { kind: "missing", placeholder: null };
  if (PLACEHOLDER.test(raw)) return { kind: "missing", placeholder: raw };

  if (IMAGE_FILE_NAME.test(raw) && !/^[a-z][a-z0-9+.-]*:\/\//i.test(raw)) {
    return { kind: "file", fileName: raw };
  }

  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    return {
      kind: "invalid",
      reason: `"${raw}" is not an image file name (.png/.jpg/…) or an https image URL`,
    };
  }

  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
    return {
      kind: "invalid",
      reason: `must be an http(s) image URL (got ${parsed.protocol})`,
    };
  }

  const host = parsed.hostname.toLowerCase();
  if (
    host === "localhost" ||
    host === "127.0.0.1" ||
    host === "0.0.0.0" ||
    host === "::1" ||
    host.endsWith(".local") ||
    host.endsWith(".internal") ||
    /^(10\.|192\.168\.|169\.254\.|127\.)/.test(host) ||
    /^172\.(1[6-9]|2\d|3[0-1])\./.test(host)
  ) {
    return { kind: "invalid", reason: "URL points to a local/internal host" };
  }

  const pathName = parsed.pathname || "/image.png";
  const base = pathName.split("/").pop() || "image.png";
  const fileName = /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(base)
    ? base
    : `${base.replace(/[^\w.-]+/g, "_") || "image"}.png`;

  return { kind: "url", url: parsed.toString(), fileName };
}

async function downloadImportImage(
  url: string,
): Promise<{ bytes: Uint8Array; contentType: string; fileName: string }> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20_000);
  try {
    const response = await fetch(url, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal,
      headers: { Accept: "image/*,*/*;q=0.8" },
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const contentType = (response.headers.get("content-type") || "").toLowerCase();
    if (contentType && !contentType.startsWith("image/") && !contentType.includes("octet-stream")) {
      throw new Error(`URL did not return an image (content-type: ${contentType})`);
    }
    const buffer = new Uint8Array(await response.arrayBuffer());
    if (!buffer.length) {
      throw new Error("Image URL returned empty content");
    }
    if (buffer.length > MAX_IMAGE_BYTES) {
      throw new Error("Image must be 10 MB or smaller");
    }
    const pathName = new URL(url).pathname;
    const base = pathName.split("/").pop() || "image.png";
    const fileName = /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(base)
      ? base
      : "image.png";
    return {
      bytes: buffer,
      contentType: contentType.startsWith("image/")
        ? contentType.split(";")[0]!.trim()
        : "image/png",
      fileName,
    };
  } finally {
    clearTimeout(timeout);
  }
}

/**
 * Download an image URL and store it on a list item (Company Logo, Candidate
 * Photo). Returns null on success or the failure reason — never throws, so an
 * image problem never fails the row it belongs to.
 */
export async function storeImportImageFromUrl(input: {
  listKey: SharePointListKey;
  itemId: string;
  fieldInternalName: string;
  url: string;
  fileName: string;
}): Promise<string | null> {
  try {
    const downloaded = await downloadImportImage(input.url);
    await uploadAndSetListImage({
      listKey: input.listKey,
      itemId: input.itemId,
      fieldInternalName: input.fieldInternalName,
      fileName: downloaded.fileName || input.fileName,
      bytes: downloaded.bytes,
      contentType: downloaded.contentType,
    });
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}
