import { NextResponse } from "next/server";

const GENERIC_TYPES = new Set([
  "application/octet-stream",
  "binary/octet-stream",
  "application/binary",
  "application/force-download",
  "application/download",
]);

const EXTENSION_TYPES: Record<string, string> = {
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  txt: "text/plain; charset=utf-8",
  csv: "text/csv; charset=utf-8",
};

function fileExtension(fileName: string | null | undefined): string {
  const name = (fileName ?? "").trim();
  const dot = name.lastIndexOf(".");
  if (dot < 0 || dot === name.length - 1) return "";
  return name.slice(dot + 1).toLowerCase();
}

function sniffMagicType(content: ArrayBuffer): string | null {
  const bytes = new Uint8Array(content.slice(0, 16));
  if (
    bytes.length >= 4 &&
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46
  ) {
    return "application/pdf";
  }
  if (
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  ) {
    return "image/jpeg";
  }
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  ) {
    return "image/png";
  }
  if (
    bytes.length >= 6 &&
    bytes[0] === 0x47 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x38 &&
    (bytes[4] === 0x37 || bytes[4] === 0x39) &&
    bytes[5] === 0x61
  ) {
    return "image/gif";
  }
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x52 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x46 &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return "image/webp";
  }
  return null;
}

export function resolveFileContentType(
  content: ArrayBuffer,
  fileName: string | null | undefined,
  reportedType?: string | null,
): string {
  const sniffed = sniffMagicType(content);
  if (sniffed) return sniffed;

  const reported = reportedType?.split(";")[0]?.trim().toLowerCase() ?? "";
  if (reported && !GENERIC_TYPES.has(reported)) {
    return reportedType!.split(";")[0]!.trim();
  }

  return EXTENSION_TYPES[fileExtension(fileName)] ?? "application/octet-stream";
}

function asciiFileName(fileName: string): string {
  const cleaned = fileName.replace(/["\\\r\n]/g, "").trim() || "document";
  return cleaned.replace(/[^\x20-\x7E]/g, "_");
}

function rfc5987FileName(fileName: string): string {
  return encodeURIComponent(fileName.replace(/[\r\n]/g, "").trim() || "document");
}

export function fileResponse(
  content: ArrayBuffer,
  fileName: string | null | undefined,
  reportedType: string | null | undefined,
  disposition: "inline" | "attachment",
): NextResponse {
  const name = (fileName ?? "document").trim() || "document";
  const contentType = resolveFileContentType(content, name, reportedType);

  return new NextResponse(new Uint8Array(content), {
    status: 200,
    headers: {
      "Content-Type": contentType,
      "Content-Disposition": `${disposition}; filename="${asciiFileName(name)}"; filename*=UTF-8''${rfc5987FileName(name)}`,
      "Cache-Control": "private, no-store",
      "X-Content-Type-Options": "nosniff",
    },
  });
}
