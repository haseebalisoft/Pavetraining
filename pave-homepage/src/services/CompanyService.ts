import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { Company } from "../models";
import { FIELDS, LIST_TITLES } from "./listSchema";
import { getSP } from "./SPContext";

function asString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (typeof value === "object") {
    const record = value as { Title?: unknown; LookupValue?: unknown };
    if (record.LookupValue != null) return asString(record.LookupValue);
    if (record.Title != null) return asString(record.Title);
  }
  return "";
}

function mapCompany(item: Record<string, unknown>): Company {
  const f = FIELDS.company;
  return {
    id: asString(item[f.id] ?? item.Id ?? item.ID),
    number: asString(item[f.number]),
    name: asString(item[f.name]),
    size: asString(item[f.size]),
    address: asString(item[f.address]),
    regNo: asString(item[f.regNo]),
    vatNo: asString(item[f.vatNo]),
    tel: asString(item[f.tel]),
    email: asString(item[f.email]),
    mainContact: asString(item[f.mainContact]),
  };
}

/**
 * Reads from SharePoint list "Company List".
 */
export class CompanyService {
  public static async getTop(
    context: WebPartContext,
    top: number = 12
  ): Promise<Company[]> {
    try {
      const sp = getSP(context);
      const f = FIELDS.company;
      const items = await sp.web.lists
        .getByTitle(LIST_TITLES.company)
        .items.select(
          f.id,
          f.number,
          f.name,
          f.size,
          f.address,
          f.regNo,
          f.vatNo,
          f.tel,
          f.email,
          f.mainContact
        )
        .orderBy(f.name, true)
        .top(Math.max(1, top))();

      return (items as Record<string, unknown>[]).map(mapCompany);
    } catch (error) {
      console.error("[CompanyService.getTop] Failed to load Company List", error);
      return [];
    }
  }
}
