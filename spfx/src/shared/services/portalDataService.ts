import { getSharePointFields } from "../schema/sharepointSchema";
import type { SharePointListKey } from "../schema/sharepointSchema";
import {
  asBoolean,
  asString,
  fieldEqualsFilter,
  getListItems,
  type SpListClient,
} from "./sharePointListService";

export interface PortalTableRow {
  id: string;
  cells: string[];
}

function cell(value: unknown): string {
  return asString(value) ?? "";
}

/**
 * Loads list rows for portal tables (admin: all; customer: company + visible).
 */
export async function loadPortalListRows(
  client: SpListClient,
  listKey: SharePointListKey,
  options: {
    companyName?: string;
    companyFieldInternalName?: string;
    customerVisibleOnly?: boolean;
    visibleFieldInternalName?: string;
    columns: string[];
    top?: number;
  }
): Promise<PortalTableRow[]> {
  const filters: string[] = [];
  if (options.companyName && options.companyFieldInternalName) {
    filters.push(
      fieldEqualsFilter(
        options.companyFieldInternalName,
        options.companyName
      )
    );
  }
  if (options.customerVisibleOnly && options.visibleFieldInternalName) {
    filters.push(`${options.visibleFieldInternalName} eq 1`);
  }

  const items = await getListItems(client, listKey, {
    filter: filters.length ? filters.join(" and ") : undefined,
    top: options.top ?? 200,
  });

  return items.map((item) => ({
    id: item.id,
    cells: options.columns.map((col) => cell(item.fields[col])),
  }));
}

export async function loadDashboardCounts(
  client: SpListClient,
  companyName?: string
): Promise<Record<string, number>> {
  const workforce = getSharePointFields("workforce");
  const matrix = getSharePointFields("trainingMatrix");
  const events = getSharePointFields("events");
  const offers = getSharePointFields("offersPromotions");
  const docs = getSharePointFields("customerDocuments");

  const companyFilter = (field: string): string | undefined =>
    companyName ? fieldEqualsFilter(field, companyName) : undefined;

  const [w, m, e, o, d] = await Promise.all([
    getListItems(client, "workforce", {
      filter: companyFilter(workforce.companyName),
      top: 5000,
    }),
    getListItems(client, "trainingMatrix", {
      filter: companyFilter(matrix.companyName),
      top: 5000,
    }),
    getListItems(client, "events", {
      filter: companyName
        ? [
            fieldEqualsFilter(events.eventCompany, companyName),
            `${events.customerVisible} eq 1`,
          ].join(" and ")
        : undefined,
      top: 500,
    }),
    getListItems(client, "offersPromotions", {
      filter: companyName ? `${offers.customerVisible} eq 1` : undefined,
      top: 200,
    }),
    getListItems(client, "customerDocuments", {
      filter: companyName
        ? [
            fieldEqualsFilter(docs.company, companyName),
            `${docs.customerVisible} eq 1`,
          ].join(" and ")
        : undefined,
      top: 2000,
    }),
  ]);

  return {
    workforce: w.length,
    matrix: m.length,
    events: e.length,
    offers: o.filter((x) =>
      companyName ? asBoolean(x.fields[offers.customerVisible]) : true
    ).length,
    documents: d.length,
  };
}
