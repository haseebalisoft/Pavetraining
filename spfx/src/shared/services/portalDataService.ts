import {
  getSharePointFieldInternalNames,
  getSharePointFieldLabel,
  getSharePointFields,
  getSharePointList,
  type SharePointListKey,
} from "../schema/sharepointSchema";
import { formatDateTime } from "../utils/formatDate";
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
  /** Raw SharePoint fields — used by Admin edit/delete. */
  fields?: Record<string, unknown>;
}

function cell(value: unknown): string {
  return asString(value) ?? "";
}

function readField(
  fields: Record<string, unknown>,
  internalName: string
): unknown {
  if (fields[internalName] !== undefined && fields[internalName] !== null) {
    return fields[internalName];
  }
  // Lookup id companion (Company → CompanyId)
  const idKey = internalName + "Id";
  if (fields[idKey] !== undefined && fields[idKey] !== null) {
    return fields[idKey];
  }
  return fields[internalName];
}

/**
 * Lookup expands — only use when paired with a matching $select.
 * Prefer no expand by default (CompanyId etc. still return without it).
 */
const EXPAND_BY_LIST: Partial<Record<SharePointListKey, string[]>> = {};

/**
 * Returns every mapped schema field for Admin tables (closer to SharePoint list view).
 */
export function getAdminSchemaColumns(listKey: SharePointListKey): {
  columns: string[];
  headers: string[];
} {
  const list = getSharePointList(listKey);
  const fields = list.fields as Record<string, string>;
  const columns: string[] = [];
  const headers: string[] = [];

  for (const key in fields) {
    if (!Object.prototype.hasOwnProperty.call(fields, key)) continue;
    if (key === "id" || key === "companyLookupId") continue;
    const internal = fields[key];
    if (columns.indexOf(internal) >= 0) continue;
    columns.push(internal);
    headers.push(getSharePointFieldLabel(listKey, key));
  }

  return { columns, headers };
}

/**
 * Loads list rows for portal tables (admin: all pages; customer: company + visible).
 */
export async function loadPortalListRows(
  client: SpListClient,
  listKey: SharePointListKey,
  options: {
    companyName?: string;
    companyFieldInternalName?: string;
    /** Prefer for Lookup columns (CompanyId / EventCompanyId). */
    companyId?: string;
    companyIdFieldInternalName?: string;
    customerVisibleOnly?: boolean;
    visibleFieldInternalName?: string;
    columns: string[];
    top?: number;
    /** When true, page through the whole list (Admin). */
    loadAll?: boolean;
  }
): Promise<PortalTableRow[]> {
  const filters: string[] = [];
  if (options.companyId && options.companyIdFieldInternalName) {
    filters.push(
      options.companyIdFieldInternalName + " eq " + Number(options.companyId)
    );
  } else if (options.companyName && options.companyFieldInternalName) {
    filters.push(
      fieldEqualsFilter(
        options.companyFieldInternalName,
        options.companyName
      )
    );
  }
  if (options.customerVisibleOnly && options.visibleFieldInternalName) {
    filters.push(options.visibleFieldInternalName + " eq 1");
  }

  const expand = EXPAND_BY_LIST[listKey] || [];
  const loadAll = options.loadAll !== false && !options.companyName && !options.companyId;

  const items = await getListItems(client, listKey, {
    filter: filters.length ? filters.join(" and ") : undefined,
    top: options.top != null ? options.top : loadAll ? 5000 : 500,
    maxItems: loadAll ? 20000 : options.top != null ? options.top : 500,
    expand: expand.length ? expand : undefined,
  });

  return items.map((item) => ({
    id: item.id,
    cells: options.columns.map((col) => cell(readField(item.fields, col))),
    fields: item.fields,
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

  const adminTop = companyName ? 5000 : 5000;
  const adminMax = companyName ? 5000 : 20000;

  const [w, m, e, o, d] = await Promise.all([
    getListItems(client, "workforce", {
      filter: companyFilter(workforce.companyName),
      top: adminTop,
      maxItems: adminMax,
    }),
    getListItems(client, "trainingMatrix", {
      // Training Matrix Update has no company column — count all (or filter later).
      top: adminTop,
      maxItems: adminMax,
    }),
    getListItems(client, "events", {
      filter: companyName
        ? [
            fieldEqualsFilter(events.eventCompany, companyName),
            events.customerVisible + " eq 1",
          ].join(" and ")
        : undefined,
      top: adminTop,
      maxItems: adminMax,
    }),
    getListItems(client, "offersPromotions", {
      filter: companyName ? offers.customerVisible + " eq 1" : undefined,
      top: 2000,
      maxItems: 2000,
    }).catch((): Awaited<ReturnType<typeof getListItems>> => []),
    getListItems(client, "customerDocuments", {
      filter: companyName
        ? [
            fieldEqualsFilter(docs.company, companyName),
            docs.customerVisible + " eq 1",
          ].join(" and ")
        : undefined,
      top: adminTop,
      maxItems: adminMax,
    }),
  ]);

  let matrixCount = m.length;
  if (companyName) {
    const names: { [n: string]: boolean } = {};
    for (let i = 0; i < w.length; i++) {
      const n = asString(w[i].fields[workforce.candidateName]);
      if (n) names[n.trim().toLowerCase()] = true;
    }
    matrixCount = m.filter((item) => {
      const title = asString(item.fields[matrix.title] || item.fields.Title);
      return Boolean(title && names[title.trim().toLowerCase()]);
    }).length;
  }

  return {
    workforce: w.length,
    matrix: matrixCount,
    events: e.length,
    offers: o.filter((x) =>
      companyName ? asBoolean(x.fields[offers.customerVisible]) : true
    ).length,
    documents: d.length,
  };
}

function daysUntil(iso: string): number | null {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.ceil((d.getTime() - today.getTime()) / 86400000);
}

function nvqProgressPct(stage: string): number {
  const s = stage.toLowerCase();
  if (s.indexOf("complete") >= 0 || s.indexOf("achieved") >= 0) return 100;
  if (s.indexOf("assess") >= 0 || s.indexOf("portfolio") >= 0) return 75;
  if (s.indexOf("induct") >= 0 || s.indexOf("register") >= 0) return 25;
  if (s.indexOf("progress") >= 0 || s.indexOf("active") >= 0) return 50;
  if (!s) return 0;
  return 40;
}

export interface CustomerDashboardDocTile {
  id: string;
  label: string;
  meta: string;
  kind: "folder" | "pdf";
}

export interface CustomerDashboardNvqRow {
  id: string;
  name: string;
  course: string;
  pct: number;
}

export interface CustomerDashboardEventRow {
  id: string;
  title: string;
  when: string;
  where: string;
  dateRaw: string;
}

export interface CustomerDashboardOfferCard {
  id: string;
  badge: string;
  title: string;
  code: string;
}

export interface CustomerDashboardData {
  counts: Record<string, number>;
  matrixRows: PortalTableRow[];
  documentTiles: CustomerDashboardDocTile[];
  nvqRows: CustomerDashboardNvqRow[];
  eventRows: CustomerDashboardEventRow[];
  offerCards: CustomerDashboardOfferCard[];
}

/**
 * Customer dashboard payload — company-scoped SharePoint lists (real data).
 */
export async function loadCustomerDashboardData(
  client: SpListClient,
  companyName: string,
  companyId?: string
): Promise<CustomerDashboardData> {
  const matrix = getSharePointFields("trainingMatrix");
  const docs = getSharePointFields("customerDocuments");
  const nvq = getSharePointFields("nvqRegister");
  const events = getSharePointFields("events");
  const offers = getSharePointFields("offersPromotions");
  const workforce = getSharePointFields("workforce");

  const safe = async <T>(p: Promise<T>, fallback: T): Promise<T> => {
    try {
      return await p;
    } catch {
      return fallback;
    }
  };

  const lookupId =
    companyId && companyId !== "0" && /^\d+$/.test(companyId)
      ? Number(companyId)
      : NaN;

  const docsFilter =
    !isNaN(lookupId)
      ? ["CompanyId eq " + lookupId, docs.customerVisible + " eq 1"].join(
          " and "
        )
      : [
          fieldEqualsFilter(docs.company, companyName),
          docs.customerVisible + " eq 1",
        ].join(" and ");

  const eventsFilter =
    !isNaN(lookupId)
      ? [
          "EventCompanyId eq " + lookupId,
          events.customerVisible + " eq 1",
        ].join(" and ")
      : [
          fieldEqualsFilter(events.eventCompany, companyName),
          events.customerVisible + " eq 1",
        ].join(" and ");

  const [wItems, mItems, dItems, nItems, eItems, oItems] = await Promise.all([
    safe(
      getListItems(client, "workforce", {
        filter: fieldEqualsFilter(workforce.companyName, companyName),
        top: 5000,
        maxItems: 5000,
      }),
      []
    ),
    safe(
      getListItems(client, "trainingMatrix", {
        top: 5000,
        maxItems: 5000,
      }),
      []
    ),
    safe(
      getListItems(client, "customerDocuments", {
        filter: docsFilter,
        top: 200,
        maxItems: 200,
      }),
      []
    ),
    safe(
      getListItems(client, "nvqRegister", {
        filter: [
          fieldEqualsFilter(nvq.companyName, companyName),
          nvq.customerVisible + " eq 1",
        ].join(" and "),
        top: 200,
        maxItems: 200,
      }),
      []
    ),
    safe(
      getListItems(client, "events", {
        filter: eventsFilter,
        top: 200,
        maxItems: 200,
      }),
      []
    ),
    safe(
      getListItems(client, "offersPromotions", {
        filter: offers.customerVisible + " eq 1",
        top: 50,
        maxItems: 50,
      }),
      []
    ),
  ]);

  const workforceNames: { [n: string]: boolean } = {};
  for (let i = 0; i < wItems.length; i++) {
    const n = asString(wItems[i].fields[workforce.candidateName]);
    if (n) workforceNames[n.trim().toLowerCase()] = true;
  }

  const companyMatrix = mItems.filter((item) => {
    const title = asString(item.fields[matrix.title] || item.fields.Title);
    return Boolean(title && workforceNames[title.trim().toLowerCase()]);
  });

  let expiringSoon = 0;
  let missingData = 0;
  const matrixRows: PortalTableRow[] = companyMatrix.map((item) => {
    const name = cell(item.fields[matrix.title] || item.fields.Title);
    const cscs = cell(item.fields[matrix.cscsExpiry] || item.fields.CSCSExpiry);
    const days = cscs ? daysUntil(cscs) : null;
    if (!cscs) {
      missingData += 1;
    } else if (days !== null && days >= 0 && days <= 60) {
      expiringSoon += 1;
    } else if (days !== null && days < 0) {
      missingData += 1;
    }
    const status =
      days === null
        ? "Missing Data"
        : days < 0
          ? "Expired"
          : days <= 60
            ? "Expiring Soon"
            : "Compliant";
    return {
      id: item.id,
      cells: [name, status, cscs],
      fields: item.fields,
    };
  });

  const documentTiles: CustomerDashboardDocTile[] = dItems
    .slice(0, 8)
    .map((item) => {
      const name =
        cell(item.fields[docs.fileLeafRef]) ||
        cell(item.fields[docs.title]) ||
        "Document";
      const type = cell(item.fields[docs.documentType]) || "";
      const fsObj = item.fields[docs.fsObjType];
      const isFolder = fsObj === 1 || fsObj === "1";
      const isPdf = !isFolder && (/\.pdf$/i.test(name) || /pdf/i.test(type));
      return {
        id: item.id,
        label: name,
        meta: type || (isFolder ? "Folder" : isPdf ? "PDF" : "File"),
        kind: isPdf ? "pdf" : "folder",
      };
    });

  const nvqRows: CustomerDashboardNvqRow[] = nItems.slice(0, 5).map((item) => {
    const stage = cell(item.fields[nvq.stageOfNvq]);
    const title = cell(item.fields[nvq.nvqTitle]);
    return {
      id: item.id,
      name: cell(item.fields[nvq.candidateName]) || "—",
      course: [title, stage].filter(Boolean).join(" · ") || "NVQ",
      pct: nvqProgressPct(stage),
    };
  });

  const eventRows: CustomerDashboardEventRow[] = eItems.slice(0, 5).map((item) => {
    const dateRaw = cell(item.fields[events.eventDate]);
    return {
      id: item.id,
      title: cell(item.fields[events.title]) || "Event",
      when: formatDateTime(dateRaw),
      where: cell(item.fields[events.trainingAddress]) || "—",
      dateRaw: dateRaw,
    };
  });

  const offerCards: CustomerDashboardOfferCard[] = oItems
    .filter((x) => asBoolean(x.fields[offers.customerVisible]))
    .slice(0, 4)
    .map((item) => {
      const category = cell(item.fields[offers.category]) || "Offer";
      const title =
        cell(item.fields[offers.title]) ||
        cell(item.fields[offers.shortDescription]) ||
        "Promotion";
      return {
        id: item.id,
        badge: category.toUpperCase(),
        title: title,
        code: "Learn more →",
      };
    });

  return {
    counts: {
      workforce: wItems.length,
      matrix: mItems.length,
      events: eItems.length,
      offers: offerCards.length,
      documents: dItems.length,
      expiringSoon: expiringSoon,
      missingData: missingData,
    },
    matrixRows: matrixRows,
    documentTiles: documentTiles,
    nvqRows: nvqRows,
    eventRows: eventRows,
    offerCards: offerCards,
  };
}

/** Convenience: schema field names for a list (debug / future edit forms). */
export function listSchemaFieldNames(listKey: SharePointListKey): string[] {
  return getSharePointFieldInternalNames(listKey);
}

