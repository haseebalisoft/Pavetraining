import { WebPartContext } from "@microsoft/sp-webpart-base";

import type {
  EusrRecord,
  InHouseRecord,
  NporsRecord,
  NrswaRecord,
  NvqRecord,
  RegisterRecord,
  WorkforceMember,
} from "../models";
import { FIELDS, LIST_TITLES } from "./listSchema";
import { getSP } from "./SPContext";

function asString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return String(value);
}

function asDateString(value: unknown): string | null {
  const raw = asString(value);
  if (!raw) return null;
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? raw : date.toISOString();
}

function mapRegister(
  item: Record<string, unknown>,
  fields: {
    id: string;
    operator: string;
    category: string;
    expiryDate: string;
    status: string;
  }
): RegisterRecord {
  return {
    id: asString(item[fields.id] ?? item.Id ?? item.ID),
    operator: asString(item[fields.operator]),
    category: asString(item[fields.category]),
    expiryDate: asDateString(item[fields.expiryDate]),
    status: asString(item[fields.status]),
  };
}

async function getRegisterTop(
  context: WebPartContext,
  listTitle: string,
  fields: {
    id: string;
    operator: string;
    category: string;
    expiryDate: string;
    status: string;
  },
  top: number,
  label: string
): Promise<RegisterRecord[]> {
  try {
    const sp = getSP(context);
    const items = await sp.web.lists
      .getByTitle(listTitle)
      .items.select(
        fields.id,
        fields.operator,
        fields.category,
        fields.expiryDate,
        fields.status
      )
      .orderBy(fields.expiryDate, true)
      .top(Math.max(1, top))();

    return (items as Record<string, unknown>[]).map((item) =>
      mapRegister(item, fields)
    );
  } catch (error) {
    console.error(`[NporsService] Failed to load ${label} (${listTitle})`, error);
    return [];
  }
}

/**
 * Certification registers + workforce helpers used by the homepage tiles.
 * Field names match the live SharePoint schema.
 */
export class NporsService {
  public static getTopNpors(
    context: WebPartContext,
    top: number = 8
  ): Promise<NporsRecord[]> {
    return getRegisterTop(
      context,
      LIST_TITLES.npors,
      FIELDS.npors,
      top,
      "NPORS Register"
    );
  }

  public static getTopEusr(
    context: WebPartContext,
    top: number = 8
  ): Promise<EusrRecord[]> {
    return getRegisterTop(
      context,
      LIST_TITLES.eusr,
      FIELDS.eusr,
      top,
      "EUSR Register"
    );
  }

  public static getTopNrswa(
    context: WebPartContext,
    top: number = 8
  ): Promise<NrswaRecord[]> {
    return getRegisterTop(
      context,
      LIST_TITLES.nrswa,
      FIELDS.nrswa,
      top,
      "NRSWA Register"
    );
  }

  public static getTopInHouse(
    context: WebPartContext,
    top: number = 8
  ): Promise<InHouseRecord[]> {
    return getRegisterTop(
      context,
      LIST_TITLES.inHouse,
      FIELDS.inHouse,
      top,
      "In-House Certificates Register"
    );
  }

  public static async getTopNvq(
    context: WebPartContext,
    top: number = 8
  ): Promise<NvqRecord[]> {
    try {
      const sp = getSP(context);
      const f = FIELDS.nvq;
      const items = await sp.web.lists
        .getByTitle(LIST_TITLES.nvq)
        .items.select(
          f.id,
          f.operator,
          f.category,
          f.expiryDate,
          f.status,
          f.dateRegistered,
          f.completedDate
        )
        .orderBy(f.dateRegistered, false)
        .top(Math.max(1, top))();

      return (items as Record<string, unknown>[]).map((item) => ({
        ...mapRegister(item, f),
        dateRegistered: asDateString(item[f.dateRegistered]),
        completedDate: asDateString(item[f.completedDate]),
      }));
    } catch (error) {
      console.error("[NporsService.getTopNvq] Failed to load NVQ Register", error);
      return [];
    }
  }

  public static async getTopWorkforce(
    context: WebPartContext,
    top: number = 12
  ): Promise<WorkforceMember[]> {
    try {
      const sp = getSP(context);
      const f = FIELDS.workforce;
      const items = await sp.web.lists
        .getByTitle(LIST_TITLES.workforce)
        .items.select(f.id, f.name, f.role, f.phone, f.email, f.status)
        .orderBy(f.name, true)
        .top(Math.max(1, top))();

      return (items as Record<string, unknown>[]).map((item) => ({
        id: asString(item[f.id] ?? item.Id ?? item.ID),
        name: asString(item[f.name]),
        role: asString(item[f.role]),
        phone: asString(item[f.phone]),
        email: asString(item[f.email]),
      }));
    } catch (error) {
      console.error(
        "[NporsService.getTopWorkforce] Failed to load Workforce List",
        error
      );
      return [];
    }
  }

  /** Loads a larger page of register rows for stats (all five registers). */
  public static async getRegisterRowsForStats(
    context: WebPartContext,
    topPerList: number = 500
  ): Promise<{
    npors: NporsRecord[];
    eusr: EusrRecord[];
    nrswa: NrswaRecord[];
    inHouse: InHouseRecord[];
    nvq: NvqRecord[];
  }> {
    const [npors, eusr, nrswa, inHouse, nvq] = await Promise.all([
      this.getTopNpors(context, topPerList),
      this.getTopEusr(context, topPerList),
      this.getTopNrswa(context, topPerList),
      this.getTopInHouse(context, topPerList),
      this.getTopNvq(context, topPerList),
    ]);
    return { npors, eusr, nrswa, inHouse, nvq };
  }
}
