import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  getAllowedCandidateNames,
  isCompanyWideScope,
} from "@/lib/services/customerAccessService";
import { getCompanyById } from "@/lib/services/companyService";
import {
  asBoolean,
  asLookupOrString,
  asMultiChoiceText,
  asNullableString,
  asString,
  buildFieldLookupIdEqualsFilter,
  buildSchemaFieldEqualsFilter,
  extractLookupId,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getWorkforceByCompanyName } from "@/lib/services/workforceService";
import { toCustomerOutcome } from "@/lib/training/customerOutcome";
import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import type {
  CustomerContext,
  CustomerEusrRecord,
  CustomerInHouseRecord,
  CustomerNporsRecord,
  CustomerStreetworksRecord,
} from "@/types/models";

const nporsFields = getSharePointFields("nporsRegister");
const eusrFields = getSharePointFields("eusrRegister");
const streetworksFields = getSharePointFields("nrswaRegister");
const inHouseFields = getSharePointFields("inHouseCertificates");

async function resolveCompanyName(companyId: string): Promise<string> {
  const company = await getCompanyById(companyId);
  if (!company?.companyName) {
    throw new Error(`Unable to resolve company name for id ${companyId}.`);
  }
  return company.companyName;
}

/**
 * Company-only Graph filter. Do not AND CustomerVisible here — the live
 * tenant returns 0 rows for that combined query even when both fields match
 * (same issue as Events / Customer Documents). Visibility is enforced in
 * the per-row mappers.
 */
function companyOnlyFilter(
  listKey:
    | "nporsRegister"
    | "eusrRegister"
    | "nrswaRegister"
    | "inHouseCertificates",
  companyId: string,
  companyName: string,
): string {
  // Prefer LookupId — avoids OData breakage on company names with & / quotes.
  const id = Number(companyId);
  if (Number.isFinite(id)) {
    return buildFieldLookupIdEqualsFilter("CompanyNameLookupId", id);
  }
  return buildSchemaFieldEqualsFilter(listKey, "companyName", companyName);
}

function isCustomerVisibleField(
  fields: SharePointFields,
  internalName: string,
): boolean {
  return (
    asBoolean(fields[internalName]) ||
    asBoolean(fields.CustomerVisible) ||
    asBoolean(fields.Customer_x0020_Visible)
  );
}

/**
 * THIS company's Workforce, indexed both ways. Register rows reference the
 * candidate by Lookup, and Graph often returns only `CandidateNameLookupId`
 * (no name text) — the mappers used to drop every such row, so customers saw
 * "No records found" for NPORS / EUSR / Streetworks / In-House records the
 * Admin register showed fine. Names are resolved only against this company's
 * candidates, so a lookup to anyone else never surfaces here.
 */
interface CompanyWorkforceIndex {
  idByName: Map<string, string>;
  nameById: Map<string, string>;
}

async function loadCompanyWorkforceIndex(
  companyName: string,
): Promise<CompanyWorkforceIndex> {
  const workforce = await getWorkforceByCompanyName(companyName);
  const idByName = new Map<string, string>();
  const nameById = new Map<string, string>();
  for (const candidate of workforce) {
    const id = String(candidate.id).trim();
    const name = candidate.candidateName.trim();
    if (!id || !name) continue;
    nameById.set(id, name);
    const key = name.toLowerCase();
    if (!idByName.has(key)) idByName.set(key, id);
  }
  return { idByName, nameById };
}

/**
 * Candidate name + Workforce id for a register row. The row's own Lookup id is
 * authoritative; the name -> id map is a same-name-unsafe fallback only for
 * legacy rows with no Lookup id. Null when the candidate is not one of this
 * company's Workforce (never shown).
 */
function resolveCandidate(
  fields: SharePointFields,
  internalName: string,
  index: CompanyWorkforceIndex,
): { candidateName: string; workforceId: string | null } | null {
  const lookupId = extractLookupId(fields, internalName);
  const text = asString(fields[internalName])?.trim() || null;
  if (lookupId) {
    const name = index.nameById.get(lookupId);
    if (!name) return null;
    return { candidateName: text ?? name, workforceId: lookupId };
  }
  if (!text) return null;
  return {
    candidateName: text,
    workforceId: index.idByName.get(text.toLowerCase()) ?? null,
  };
}

function mapNpors(
  id: string,
  fields: SharePointFields,
  index: CompanyWorkforceIndex,
): CustomerNporsRecord | null {
  if (!isCustomerVisibleField(fields, nporsFields.customerVisible)) {
    return null;
  }

  const candidate = resolveCandidate(
    fields,
    nporsFields.candidateName,
    index,
  );
  if (!candidate) {
    return null;
  }
  const { candidateName, workforceId } = candidate;

  return {
    id,
    candidateName,
    workforceId,
    nporsNumber: asNullableString(fields[nporsFields.nporsNumber]),
    trainingDate: asNullableString(fields[nporsFields.trainingDate]),
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[nporsFields.trainingAddress]),
    ),
    noviceOrEwt: asNullableString(fields[nporsFields.noviceOrEwt]),
    nporsCategory: asMultiChoiceText(fields[nporsFields.nporsCategory]),
    outcome: toCustomerOutcome(
      asNullableString(fields[nporsFields.trainingOutcome]),
    ),
    expiry: asNullableString(fields[nporsFields.expiry]),
  };
}

function mapEusr(
  id: string,
  fields: SharePointFields,
  index: CompanyWorkforceIndex,
): CustomerEusrRecord | null {
  if (!isCustomerVisibleField(fields, eusrFields.customerVisible)) {
    return null;
  }

  const candidate = resolveCandidate(
    fields,
    eusrFields.candidateName,
    index,
  );
  if (!candidate) {
    return null;
  }
  const { candidateName, workforceId } = candidate;

  return {
    id,
    candidateName,
    workforceId,
    eusrNumber: asNullableString(fields[eusrFields.eusrNumber]),
    eusrCategory: asMultiChoiceText(fields[eusrFields.eusrCategory]),
    cardType: asNullableString(fields[eusrFields.cardType]),
    trainingDate: asNullableString(fields[eusrFields.trainingDate]),
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[eusrFields.trainingAddress]),
    ),
    outcome: toCustomerOutcome(
      asNullableString(fields[eusrFields.trainingOutcome]),
    ),
    expiry: asNullableString(fields[eusrFields.expiry]),
  };
}

function mapStreetworks(
  id: string,
  fields: SharePointFields,
  index: CompanyWorkforceIndex,
): CustomerStreetworksRecord | null {
  if (!isCustomerVisibleField(fields, streetworksFields.customerVisible)) {
    return null;
  }

  const candidate = resolveCandidate(
    fields,
    streetworksFields.candidateName,
    index,
  );
  if (!candidate) {
    return null;
  }
  const { candidateName, workforceId } = candidate;

  const rawNotes = asNullableString(fields[streetworksFields.outcomeNotes]);
  const endMatch = rawNotes?.match(/\[TrainingEnd:(\d{4}-\d{2}-\d{2})\]/i);

  return {
    id,
    candidateName,
    workforceId,
    swqrNumber: asNullableString(fields[streetworksFields.swqrNumber]),
    trainingDate: asNullableString(fields[streetworksFields.trainingDate]),
    trainingDateEnd: endMatch?.[1] ?? null,
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[streetworksFields.trainingAddress]),
    ),
    course: asNullableString(fields[streetworksFields.course]),
    streetworksCategory: asMultiChoiceText(
      fields[streetworksFields.streetworksCategory],
    ),
    outcome: toCustomerOutcome(
      asNullableString(fields[streetworksFields.trainingOutcome]),
    ),
    expiry: asNullableString(fields[streetworksFields.expiryDate]),
  };
}

function mapInHouse(
  id: string,
  fields: SharePointFields,
  index: CompanyWorkforceIndex,
): CustomerInHouseRecord | null {
  if (!isCustomerVisibleField(fields, inHouseFields.customerVisible)) {
    return null;
  }

  const candidate = resolveCandidate(
    fields,
    inHouseFields.candidateName,
    index,
  );
  if (!candidate) {
    return null;
  }
  const { candidateName, workforceId } = candidate;

  // Prefer Certificate Category (live SharePoint choice); fall back to legacy CourseCategory.
  const course =
    asNullableString(fields[inHouseFields.certificateCategory]) ??
    asNullableString(fields[inHouseFields.courseCategory]);

  return {
    id,
    candidateName,
    workforceId,
    course,
    certificationNumber:
      asNullableString(fields[inHouseFields.certificationNumber]) ??
      asLookupOrString(
        fields.Candidate_x0020_Name_x003a__x002,
      ),
    trainingDate: asNullableString(fields[inHouseFields.courseDate]),
    trainingAddress: stripSharePointHtml(
      asNullableString(fields[inHouseFields.trainingAddress]),
    ),
    expiry: asNullableString(fields[inHouseFields.expiryDate]),
    outcome: toCustomerOutcome(
      asNullableString(fields[inHouseFields.trainingOutcome]),
    ),
  };
}

/**
 * Customer-scoped NPORS records. Company comes from Permissions-derived companyId.
 * Only CustomerVisible records are returned. Admin-only fields are omitted.
 */
export async function getCustomerNporsRecords(
  companyId: string,
  context?: CustomerContext,
): Promise<CustomerNporsRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIndex] = await Promise.all([
    getListItemsByKey("nporsRegister", {
      filter: companyOnlyFilter("nporsRegister", companyId, companyName),
      top: 5000,
    }),
    loadCompanyWorkforceIndex(companyName),
  ]);

  let rows = items
    .map((item) => mapNpors(item.id, item.fields, workforceIndex))
    .filter((row): row is CustomerNporsRecord => row !== null);

  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    const allowedNames = await getAllowedCandidateNames(context);
    rows = rows.filter((row) =>
      allowedNames.has(row.candidateName.trim().toLowerCase()),
    );
  }

  return rows;
}

export async function getCustomerEusrRecords(
  companyId: string,
  context?: CustomerContext,
): Promise<CustomerEusrRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIndex] = await Promise.all([
    getListItemsByKey("eusrRegister", {
      filter: companyOnlyFilter("eusrRegister", companyId, companyName),
      top: 5000,
    }),
    loadCompanyWorkforceIndex(companyName),
  ]);

  let rows = items
    .map((item) => mapEusr(item.id, item.fields, workforceIndex))
    .filter((row): row is CustomerEusrRecord => row !== null);

  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    const allowedNames = await getAllowedCandidateNames(context);
    rows = rows.filter((row) =>
      allowedNames.has(row.candidateName.trim().toLowerCase()),
    );
  }

  return rows;
}

/**
 * Streetworks Training (SharePoint list: NRSWA Register).
 * Customer-facing name must never expose "NRSWA".
 */
export async function getCustomerStreetworksRecords(
  companyId: string,
  context?: CustomerContext,
): Promise<CustomerStreetworksRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIndex] = await Promise.all([
    getListItemsByKey("nrswaRegister", {
      filter: companyOnlyFilter("nrswaRegister", companyId, companyName),
      top: 5000,
    }),
    loadCompanyWorkforceIndex(companyName),
  ]);

  let rows = items
    .map((item) => mapStreetworks(item.id, item.fields, workforceIndex))
    .filter((row): row is CustomerStreetworksRecord => row !== null);

  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    const allowedNames = await getAllowedCandidateNames(context);
    rows = rows.filter((row) =>
      allowedNames.has(row.candidateName.trim().toLowerCase()),
    );
  }

  return rows;
}

export async function getCustomerInHouseRecords(
  companyId: string,
  context?: CustomerContext,
): Promise<CustomerInHouseRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIndex] = await Promise.all([
    getListItemsByKey("inHouseCertificates", {
      filter: companyOnlyFilter(
        "inHouseCertificates",
        companyId,
        companyName,
      ),
      top: 5000,
    }),
    loadCompanyWorkforceIndex(companyName),
  ]);

  let rows = items
    .map((item) => mapInHouse(item.id, item.fields, workforceIndex))
    .filter((row): row is CustomerInHouseRecord => row !== null);

  if (context && !isCompanyWideScope(context.normalizedAccessScope)) {
    const allowedNames = await getAllowedCandidateNames(context);
    rows = rows.filter((row) =>
      allowedNames.has(row.candidateName.trim().toLowerCase()),
    );
  }

  return rows;
}
