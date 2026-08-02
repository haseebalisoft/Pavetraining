import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { PopularDocument } from "../models";
import { FIELDS, LIST_TITLES } from "./listSchema";
import { getSP } from "./SPContext";

function asString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  return String(value);
}

function asDateString(value: unknown): string | null {
  const raw = asString(value);
  if (!raw) return null;
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? raw : date.toISOString();
}

function iconTypeFromName(fileName: string): string {
  const ext = (fileName.split(".").pop() || "").toLowerCase();
  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].indexOf(ext) >= 0) {
    return "image";
  }
  if (ext === "pdf") return "pdf";
  if (["doc", "docx"].indexOf(ext) >= 0) return "word";
  if (["xls", "xlsx", "csv"].indexOf(ext) >= 0) return "excel";
  if (["ppt", "pptx"].indexOf(ext) >= 0) return "powerpoint";
  return "file";
}

function mapDocument(
  item: Record<string, unknown>,
  webUrl: string
): PopularDocument {
  const f = FIELDS.documents;
  const name = asString(item[f.name] ?? item.FileLeafRef ?? item.Title);
  const serverRelative = asString(item[f.url] ?? item.FileRef);
  const absolute =
    serverRelative.indexOf("http") === 0
      ? serverRelative
      : serverRelative
        ? new URL(serverRelative, webUrl).toString()
        : "";

  return {
    id: asString(item[f.id] ?? item.Id ?? item.ID),
    name,
    url: absolute,
    modified: asDateString(item[f.modified] ?? item.Modified),
    iconType: iconTypeFromName(name),
  };
}

async function queryLibrary(
  context: WebPartContext,
  listTitle: string,
  top: number
): Promise<PopularDocument[]> {
  const sp = getSP(context);
  const f = FIELDS.documents;
  const items = await sp.web.lists
    .getByTitle(listTitle)
    .items.select(f.id, f.name, f.url, f.modified, "FSObjType")
    .filter("FSObjType eq 0")
    .orderBy(f.modified, false)
    .top(Math.max(1, top))();

  const webUrl = context.pageContext.web.absoluteUrl;
  return (items as Record<string, unknown>[]).map((item) =>
    mapDocument(item, webUrl)
  );
}

/**
 * Popular documents for the homepage.
 * Tries the "Documents" library first, then "Customer Documents".
 */
export class DocumentsService {
  public static async getPopular(
    context: WebPartContext,
    top: number = 6
  ): Promise<PopularDocument[]> {
    try {
      return await queryLibrary(context, LIST_TITLES.documents, top);
    } catch (primaryError) {
      console.warn(
        '[DocumentsService.getPopular] "Documents" unavailable, trying Customer Documents',
        primaryError
      );
      try {
        return await queryLibrary(
          context,
          LIST_TITLES.customerDocuments,
          top
        );
      } catch (fallbackError) {
        console.error(
          "[DocumentsService.getPopular] Failed to load documents",
          fallbackError
        );
        throw fallbackError;
      }
    }
  }
}
