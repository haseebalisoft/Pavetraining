import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  MAX_IMAGE_BYTES,
  uploadAndSetListImage,
} from "@/lib/services/listThumbnailService";

const companyFields = getSharePointFields("company");

const PLACEHOLDER_LOGO = /^(logo(\s+placeholder)?|placeholder|n\/?a|none|null|-|—|–)$/i;

export type CompanyLogoPlan =
  | { kind: "missing"; warning: string }
  | { kind: "invalid"; warning: string }
  | { kind: "url"; url: string; fileName: string };

/**
 * Company Logo cells accept an https/http image URL. Placeholders and blanks
 * are treated as missing (warning only — row still imports). Free text is
 * never written into the SharePoint Thumbnail column.
 */
export function planCompanyLogoImport(
  value: string | null | undefined,
): CompanyLogoPlan {
  const raw = (value ?? "").trim();
  if (!raw || PLACEHOLDER_LOGO.test(raw)) {
    return {
      kind: "missing",
      warning:
        "Company Logo missing — company will import without a logo. Put an https image URL in Company Logo to upload one.",
    };
  }

  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    return {
      kind: "invalid",
      warning: `Company Logo "${raw}" is not a valid image URL — company will import without a logo.`,
    };
  }

  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
    return {
      kind: "invalid",
      warning: `Company Logo must be an http(s) image URL (got ${parsed.protocol}) — company will import without a logo.`,
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
    return {
      kind: "invalid",
      warning:
        "Company Logo URL points to a local/internal host — company will import without a logo.",
    };
  }

  const pathName = parsed.pathname || "/logo.png";
  const base = pathName.split("/").pop() || "logo.png";
  const fileName = /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(base)
    ? base
    : `${base.replace(/[^\w.-]+/g, "_") || "logo"}.png`;

  return { kind: "url", url: parsed.toString(), fileName };
}

async function downloadLogoBytes(
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
    const base = pathName.split("/").pop() || "logo.png";
    const fileName = /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(base)
      ? base
      : "logo.png";
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
 * After the Company List row exists, download the logo URL and store it with
 * the same PortalMedia + CompanyLogo Thumbnail path as Admin logo upload.
 * Returns a warning string on skip/failure (never throws for logo problems).
 */
export async function applyCompanyLogoFromImportCell(
  companyId: string,
  cell: string | null | undefined,
): Promise<string | null> {
  const plan = planCompanyLogoImport(cell);
  if (plan.kind === "missing" || plan.kind === "invalid") {
    return plan.warning;
  }

  try {
    const downloaded = await downloadLogoBytes(plan.url);
    await uploadAndSetListImage({
      listKey: "company",
      itemId: companyId,
      fieldInternalName: companyFields.companyLogo,
      fileName: downloaded.fileName || plan.fileName,
      bytes: downloaded.bytes,
      contentType: downloaded.contentType,
    });
    return null;
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    return `Company Logo upload failed (${detail}) — company imported without a logo.`;
  }
}
