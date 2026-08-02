import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  extractCandidateLookupId,
  getAllowedCandidateIds,
  getAllowedCandidateNames,
  getCompanyCandidateNameMap,
  isCompanyWideScope,
  resolveCandidateDisplayNameSync,
} from "@/lib/services/customerAccessService";
import { getCompanyById } from "@/lib/services/companyService";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  buildFieldLookupIdEqualsFilter,
  buildSchemaFieldEqualsFilter,
  getListItemByKey,
  getListItemFileContent,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import { parseThumbnailField } from "@/lib/services/listThumbnailService";
import type {
  CustomerContext,
  CustomerDocumentRecord,
  CustomerEventRecord,
  CustomerNvqRecord,
  CustomerNvqStatus,
  CustomerOfferRecord,
} from "@/types/models";

const nvqFields = getSharePointFields("nvqRegister");
const documentFields = getSharePointFields("customerDocuments");
const eventFields = getSharePointFields("events");
const offerFields = getSharePointFields("offersPromotions");

async function resolveCompanyName(companyId: string): Promise<string> {
  const company = await getCompanyById(companyId);
  if (!company?.companyName) {
    throw new Error(`Unable to resolve company name for id ${companyId}.`);
  }
  return company.companyName;
}

function companyAndVisibleFilter(
  listKey: "nvqRegister" | "customerDocuments" | "events" | "offersPromotions",
  companyFieldKey: string,
  companyName: string,
  companyId?: string,
): string {
  const id = Number(companyId);
  let companyFilter: string;
  if (listKey === "nvqRegister" && Number.isFinite(id)) {
    // Admin writes NVQCompanyLookupId — prefer that over display-name text.
    companyFilter = buildFieldLookupIdEqualsFilter("NVQCompanyLookupId", id);
  } else {
    companyFilter = buildSchemaFieldEqualsFilter(
      listKey,
      companyFieldKey,
      companyName,
    );
  }
  const fields = getSharePointFields(listKey) as Record<string, string>;
  const visibleField = fields.customerVisible;
  return `${companyFilter} and fields/${visibleField} eq true`;
}

/**
 * Prefer CompanyLookupId (indexed) over Company display-name text filters.
 * FSObjType is filtered in memory — including it in OData often makes Graph hang.
 */
function customerDocumentsByCompanyIdFilter(companyId: string): string {
  const id = Number(companyId);
  const companyClause = Number.isFinite(id)
    ? `fields/${documentFields.companyLookupId} eq ${id}`
    : buildSchemaFieldEqualsFilter("customerDocuments", "company", companyId);
  return `${companyClause} and fields/${documentFields.customerVisible} eq true`;
}

function matchesCompany(
  value: string | null | undefined,
  companyName: string,
): boolean {
  return (value ?? "").trim().toLowerCase() === companyName.trim().toLowerCase();
}

/** SharePoint FSObjType: 0 = file, 1 = folder. */
function isSharePointFile(fields: SharePointFields): boolean {
  const fs = fields[documentFields.fsObjType];
  if (fs === 1 || fs === "1") {
    return false;
  }
  if (fs === 0 || fs === "0") {
    return true;
  }
  // Fallback when FSObjType is missing: treat leaf refs with an extension as files.
  const leaf = asString(fields[documentFields.fileLeafRef]);
  if (!leaf) {
    return false;
  }
  return leaf.includes(".");
}

function isActiveOfferStatus(status: string | null): boolean {
  if (!status?.trim()) {
    return true;
  }
  const normalized = status.trim().toLowerCase();
  return (
    normalized === "active" ||
    normalized === "open" ||
    normalized === "live" ||
    normalized === "published"
  );
}

function offerHasNotExpired(endDate: string | null): boolean {
  if (!endDate?.trim()) return true;

  const dateOnly = /^(\d{4}-\d{2}-\d{2})/.exec(endDate.trim())?.[1];
  if (dateOnly) {
    return dateOnly >= new Date().toISOString().slice(0, 10);
  }

  const expiry = new Date(endDate).getTime();
  return Number.isNaN(expiry) || expiry >= Date.now();
}

function offerImageUrl(id: string, value: unknown): string | null {
  if (parseThumbnailField(value)) {
    return `/api/media/offer/${id}/image`;
  }
  return asNullableString(value);
}

function nvqStatus(completedDate: string | null): CustomerNvqStatus {
  return completedDate?.trim() ? "Completed" : "Active";
}

function mapNvq(id: string, fields: SharePointFields): CustomerNvqRecord | null {
  if (!asBoolean(fields[nvqFields.customerVisible])) {
    return null;
  }

  const candidateName = asString(fields[nvqFields.candidateName]);
  if (!candidateName) {
    return null;
  }

  const completedDate = asNullableString(fields[nvqFields.completedDate]);

  return {
    id,
    candidateName,
    nvqTitle: asNullableString(fields[nvqFields.nvqTitle]),
    boltOn: asNullableString(fields[nvqFields.boltonNvq]),
    dateRegistered: asNullableString(fields[nvqFields.dateRegistered]),
    inductionDate: asNullableString(fields[nvqFields.dateInductionBooked]),
    stageOfNvq: asNullableString(fields[nvqFields.stageOfNvq]),
    notes: asNullableString(fields[nvqFields.customerUpdateNotes]),
    completedDate,
    status: nvqStatus(completedDate),
  };
}

function mapDocument(
  id: string,
  fields: SharePointFields,
  uploadedDate: string | null,
  canDownload: boolean,
  candidateName: string | null,
): CustomerDocumentRecord | null {
  if (!isSharePointFile(fields)) {
    return null;
  }

  if (!asBoolean(fields[documentFields.customerVisible])) {
    return null;
  }

  const name =
    asString(fields[documentFields.fileLeafRef]) ??
    asString(fields[documentFields.title]);
  if (!name) {
    return null;
  }

  const hasFile = Boolean(
    asString(fields[documentFields.fileRef]) ||
      asString(fields[documentFields.fileLeafRef]),
  );
  if (!hasFile) {
    return null;
  }

  return {
    id,
    name,
    documentType: asNullableString(fields[documentFields.documentType]),
    candidate: candidateName,
    uploadedDate,
    canDownload,
    viewPath: `/api/customer/documents/${id}/view`,
    downloadPath: canDownload
      ? `/api/customer/documents/${id}/download`
      : null,
  };
}

function mapEvent(
  id: string,
  fields: SharePointFields,
  companyId: string,
  companyName: string,
): CustomerEventRecord | null {
  const visible =
    asBoolean(fields[eventFields.customerVisible]) ||
    asBoolean(fields.CustomerVisible);
  if (!visible) {
    return null;
  }

  const eventCompanyId =
    asString(fields[eventFields.eventCompanyLookupId]) ??
    asString(fields.EventCompanyId) ??
    (typeof fields[eventFields.eventCompany] === "object"
      ? asString(
          (fields[eventFields.eventCompany] as { LookupId?: unknown }).LookupId,
        )
      : null);

  // Prefer lookup ID match — Graph often has no LookupValue text.
  const idMatches =
    Boolean(eventCompanyId) &&
    String(eventCompanyId).trim() === String(companyId).trim();
  const nameMatches = matchesCompany(
    asLookupOrString(fields[eventFields.eventCompany]),
    companyName,
  );
  if (!idMatches && !nameMatches) {
    return null;
  }

  const title = asString(fields[eventFields.title]);
  if (!title) {
    return null;
  }

  return {
    id,
    title,
    eventDate: asNullableString(fields[eventFields.eventDate]),
    endDate: asNullableString(fields[eventFields.endDate]),
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[eventFields.trainingAddress]),
    ),
    location: asNullableString(fields[eventFields.location]),
    description: stripSharePointHtml(
      asNullableString(fields[eventFields.description]),
    ),
    company: companyName,
  };
}

function mapOffer(
  id: string,
  fields: SharePointFields,
): CustomerOfferRecord | null {
  if (!asBoolean(fields[offerFields.customerVisible])) {
    return null;
  }

  const status = asNullableString(fields[offerFields.status]);
  if (!isActiveOfferStatus(status)) {
    return null;
  }

  const endDate = asNullableString(fields[offerFields.endDate]);
  if (!offerHasNotExpired(endDate)) {
    return null;
  }

  const title = asString(fields[offerFields.title]);
  if (!title) {
    return null;
  }

  return {
    id,
    title,
    category: asNullableString(fields[offerFields.category]),
    description: asNullableString(fields[offerFields.shortDescription]),
    startDate: asNullableString(fields[offerFields.startDate]),
    endDate,
    status: status ?? "Active",
    image: offerImageUrl(id, fields[offerFields.image]),
    ctaLabel: asNullableString(fields[offerFields.ctaLabel]),
    ctaLink: asNullableString(fields[offerFields.ctaLink]),
  };
}

export async function getCustomerNvqRecords(
  companyId: string,
  context?: CustomerContext,
): Promise<CustomerNvqRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const items = await getListItemsByKey("nvqRegister", {
    filter: companyAndVisibleFilter(
      "nvqRegister",
      "companyName",
      companyName,
      companyId,
    ),
    top: 5000,
  });

  let rows = items
    .map((item) => mapNvq(item.id, item.fields))
    .filter((row): row is CustomerNvqRecord => row !== null);

  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    const allowedNames = await getAllowedCandidateNames(context);
    rows = rows.filter((row) =>
      allowedNames.has(row.candidateName.trim().toLowerCase()),
    );
  }

  return rows;
}

export async function getCustomerDocumentRecords(
  companyId: string,
  canDownload: boolean,
  context?: CustomerContext,
): Promise<CustomerDocumentRecord[]> {
  // Prefer company name already resolved on the permission context.
  const companyName =
    context?.companyName?.trim() || (await resolveCompanyName(companyId));

  // Indexed CompanyLookupId filter — avoid slow Company text + FSObjType OData.
  const [items, nameCache, allowedIds] = await Promise.all([
    getListItemsByKey("customerDocuments", {
      filter: customerDocumentsByCompanyIdFilter(companyId),
      top: 500,
    }),
    getCompanyCandidateNameMap(companyName),
    context && !isCompanyWideScope(context.normalizedAccessScope)
      ? getAllowedCandidateIds(context)
      : Promise.resolve(null),
  ]);

  const needsScope = allowedIds != null;
  const rows: CustomerDocumentRecord[] = [];

  for (const item of items) {
    if (!isSharePointFile(item.fields)) {
      continue;
    }

    const company = asLookupOrString(item.fields[documentFields.company]);
    const companyLookupId = asString(
      item.fields[documentFields.companyLookupId],
    );
    const companyOk =
      companyLookupId === companyId ||
      matchesCompany(company, companyName);
    if (!companyOk) {
      continue;
    }

    const candidateRaw = item.fields[documentFields.candidate];
    const candidateId =
      asString(item.fields[documentFields.candidateLookupId]) ||
      extractCandidateLookupId(item.fields, documentFields.candidate);

    if (needsScope && allowedIds && context) {
      // Company-level docs (no candidate) — Training Manager only.
      if (!candidateId) {
        if (context.customerRole !== "TrainingManager") {
          continue;
        }
      } else if (!allowedIds.has(candidateId)) {
        continue;
      }
    }

    const candidateName = resolveCandidateDisplayNameSync(
      candidateRaw,
      candidateId,
      nameCache,
    );

    // Never surface raw numeric IDs as the candidate label.
    const safeCandidateName =
      candidateName && !/^\d+$/.test(candidateName.trim())
        ? candidateName
        : null;

    const modified =
      item.lastModifiedDateTime ??
      asNullableString(item.fields[documentFields.modified]) ??
      item.createdDateTime ??
      asNullableString(item.fields.Created) ??
      asNullableString(item.fields.CreatedDateTime);

    const mapped = mapDocument(
      item.id,
      item.fields,
      modified,
      canDownload,
      safeCandidateName,
    );
    if (mapped) {
      rows.push(mapped);
    }
  }

  return rows;
}

export async function getCustomerDocumentForAccess(
  companyId: string,
  documentId: string,
  context?: CustomerContext,
): Promise<{
  id: string;
  name: string;
  companyMatches: boolean;
  customerVisible: boolean;
  isFile: boolean;
  candidateId: string | null;
  scopeAllowed: boolean;
} | null> {
  const companyName = await resolveCompanyName(companyId);
  const item = await getListItemByKey("customerDocuments", documentId);
  if (!item) {
    return null;
  }

  const company = asLookupOrString(item.fields[documentFields.company]);
  const name =
    asString(item.fields[documentFields.fileLeafRef]) ??
    asString(item.fields[documentFields.title]) ??
    "document";

  const candidateId =
    asString(item.fields[documentFields.candidateLookupId]) ||
    extractCandidateLookupId(item.fields, documentFields.candidate);

  let scopeAllowed = true;
  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    if (!candidateId) {
      scopeAllowed = context.customerRole === "TrainingManager";
    } else {
      const allowedIds = await getAllowedCandidateIds(context);
      scopeAllowed = allowedIds.has(candidateId);
    }
  }

  return {
    id: item.id,
    name,
    companyMatches: matchesCompany(company, companyName),
    customerVisible: asBoolean(item.fields[documentFields.customerVisible]),
    isFile: isSharePointFile(item.fields),
    candidateId,
    scopeAllowed,
  };
}

/** @deprecated Prefer getCustomerDocumentForAccess — kept for call-site clarity. */
export async function getCustomerDocumentForDownload(
  companyId: string,
  documentId: string,
  context?: CustomerContext,
) {
  return getCustomerDocumentForAccess(companyId, documentId, context);
}

export async function downloadCustomerDocumentFile(documentId: string) {
  return getListItemFileContent("customerDocuments", documentId);
}

export async function getCustomerEventRecords(
  companyId: string,
): Promise<CustomerEventRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const id = Number(companyId);

  // IMPORTANT: Do not AND Customer_x0020_Visible into the Graph OData filter.
  // Live tenant returns 0 rows for that combined filter even when both fields match
  // (EventCompanyLookupId eq 27 works alone; Visible filter is unreliable with the encoded name).
  // Filter company via lookup ID when possible, then enforce visibility in mapEvent.
  let items;
  if (Number.isFinite(id)) {
    items = await getListItemsByKey("events", {
      filter: `fields/${eventFields.eventCompanyLookupId} eq ${id}`,
      top: 5000,
    }).catch(async () => getListItemsByKey("events", { top: 5000 }));
  } else {
    items = await getListItemsByKey("events", { top: 5000 });
  }

  return items
    .map((item) => mapEvent(item.id, item.fields, companyId, companyName))
    .filter((row): row is CustomerEventRecord => row !== null)
    .sort((a, b) => {
      const aTime = a.eventDate ? new Date(a.eventDate).getTime() : 0;
      const bTime = b.eventDate ? new Date(b.eventDate).getTime() : 0;
      return aTime - bTime;
    });
}

export async function getCustomerOfferRecords(
  companyId?: string,
): Promise<CustomerOfferRecord[]> {
  // Offers / Promotions is site-wide (no Company column). Show visible + active only.
  void companyId;
  const items = await getListItemsByKey("offersPromotions", {
    filter: `fields/${offerFields.customerVisible} eq true`,
    top: 5000,
  });

  return items
    .map((item) => mapOffer(item.id, item.fields))
    .filter((row): row is CustomerOfferRecord => row !== null)
    .sort((a, b) => {
      const aTime = a.startDate ? new Date(a.startDate).getTime() : 0;
      const bTime = b.startDate ? new Date(b.startDate).getTime() : 0;
      return bTime - aTime;
    });
}
