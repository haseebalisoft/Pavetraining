import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { getCompanyById } from "@/lib/services/companyService";
import {
  asBoolean,
  asNullableString,
  asString,
  buildSchemaFieldEqualsFilter,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import { getWorkforceIdByCandidateName } from "@/lib/services/workforceService";
import { toCustomerOutcome } from "@/lib/training/customerOutcome";
import type {
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

function companyAndVisibleFilter(
  listKey:
    | "nporsRegister"
    | "eusrRegister"
    | "nrswaRegister"
    | "inHouseCertificates",
  companyName: string,
): string {
  const companyFilter = buildSchemaFieldEqualsFilter(
    listKey,
    "companyName",
    companyName,
  );
  const fields = getSharePointFields(listKey);
  const visibleField = fields.customerVisible;
  return `${companyFilter} and fields/${visibleField} eq true`;
}

function resolveWorkforceId(
  candidateName: string,
  workforceIds: Map<string, string>,
): string | null {
  return workforceIds.get(candidateName.trim().toLowerCase()) ?? null;
}

function mapNpors(
  id: string,
  fields: SharePointFields,
  workforceIds: Map<string, string>,
): CustomerNporsRecord | null {
  if (!asBoolean(fields[nporsFields.customerVisible])) {
    return null;
  }

  const candidateName = asString(fields[nporsFields.candidateName]);
  if (!candidateName) {
    return null;
  }

  return {
    id,
    candidateName,
    workforceId: resolveWorkforceId(candidateName, workforceIds),
    nporsNumber: asNullableString(fields[nporsFields.nporsNumber]),
    trainingDate: asNullableString(fields[nporsFields.trainingDate]),
    trainingAddress: asNullableString(fields[nporsFields.trainingAddress]),
    noviceOrEwt: asNullableString(fields[nporsFields.noviceOrEwt]),
    nporsCategory: asNullableString(fields[nporsFields.nporsCategory]),
    outcome: toCustomerOutcome(
      asNullableString(fields[nporsFields.trainingOutcome]),
    ),
    expiry: asNullableString(fields[nporsFields.expiry]),
  };
}

function mapEusr(
  id: string,
  fields: SharePointFields,
  workforceIds: Map<string, string>,
): CustomerEusrRecord | null {
  if (!asBoolean(fields[eusrFields.customerVisible])) {
    return null;
  }

  const candidateName = asString(fields[eusrFields.candidateName]);
  if (!candidateName) {
    return null;
  }

  return {
    id,
    candidateName,
    workforceId: resolveWorkforceId(candidateName, workforceIds),
    eusrNumber: asNullableString(fields[eusrFields.eusrNumber]),
    eusrCategory: asNullableString(fields[eusrFields.eusrCategory]),
    trainingDate: asNullableString(fields[eusrFields.trainingDate]),
    trainingAddress: asNullableString(fields[eusrFields.trainingAddress]),
    outcome: toCustomerOutcome(
      asNullableString(fields[eusrFields.trainingOutcome]),
    ),
    expiry: asNullableString(fields[eusrFields.expiry]),
  };
}

function mapStreetworks(
  id: string,
  fields: SharePointFields,
  workforceIds: Map<string, string>,
): CustomerStreetworksRecord | null {
  if (!asBoolean(fields[streetworksFields.customerVisible])) {
    return null;
  }

  const candidateName = asString(fields[streetworksFields.candidateName]);
  if (!candidateName) {
    return null;
  }

  return {
    id,
    candidateName,
    workforceId: resolveWorkforceId(candidateName, workforceIds),
    swqrNumber: asNullableString(fields[streetworksFields.swqrNumber]),
    trainingDate: asNullableString(fields[streetworksFields.trainingDate]),
    trainingAddress: asNullableString(
      fields[streetworksFields.trainingAddress],
    ),
    course: asNullableString(fields[streetworksFields.course]),
    streetworksCategory: asNullableString(
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
  workforceIds: Map<string, string>,
): CustomerInHouseRecord | null {
  if (!asBoolean(fields[inHouseFields.customerVisible])) {
    return null;
  }

  const candidateName = asString(fields[inHouseFields.candidateName]);
  if (!candidateName) {
    return null;
  }

  const course =
    asNullableString(fields[inHouseFields.courseCategory]) ??
    asNullableString(fields[inHouseFields.certificateCategory]);

  return {
    id,
    candidateName,
    workforceId: resolveWorkforceId(candidateName, workforceIds),
    course,
    trainingDate: asNullableString(fields[inHouseFields.courseDate]),
    trainingAddress: asNullableString(fields[inHouseFields.trainingAddress]),
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
): Promise<CustomerNporsRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIds] = await Promise.all([
    getListItemsByKey("nporsRegister", {
      filter: companyAndVisibleFilter("nporsRegister", companyName),
      top: 5000,
    }),
    getWorkforceIdByCandidateName(companyName),
  ]);

  return items
    .map((item) => mapNpors(item.id, item.fields, workforceIds))
    .filter((row): row is CustomerNporsRecord => row !== null);
}

export async function getCustomerEusrRecords(
  companyId: string,
): Promise<CustomerEusrRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIds] = await Promise.all([
    getListItemsByKey("eusrRegister", {
      filter: companyAndVisibleFilter("eusrRegister", companyName),
      top: 5000,
    }),
    getWorkforceIdByCandidateName(companyName),
  ]);

  return items
    .map((item) => mapEusr(item.id, item.fields, workforceIds))
    .filter((row): row is CustomerEusrRecord => row !== null);
}

/**
 * Streetworks Training (SharePoint list: NRSWA Register).
 * Customer-facing name must never expose "NRSWA".
 */
export async function getCustomerStreetworksRecords(
  companyId: string,
): Promise<CustomerStreetworksRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIds] = await Promise.all([
    getListItemsByKey("nrswaRegister", {
      filter: companyAndVisibleFilter("nrswaRegister", companyName),
      top: 5000,
    }),
    getWorkforceIdByCandidateName(companyName),
  ]);

  return items
    .map((item) => mapStreetworks(item.id, item.fields, workforceIds))
    .filter((row): row is CustomerStreetworksRecord => row !== null);
}

export async function getCustomerInHouseRecords(
  companyId: string,
): Promise<CustomerInHouseRecord[]> {
  const companyName = await resolveCompanyName(companyId);
  const [items, workforceIds] = await Promise.all([
    getListItemsByKey("inHouseCertificates", {
      filter: companyAndVisibleFilter("inHouseCertificates", companyName),
      top: 5000,
    }),
    getWorkforceIdByCandidateName(companyName),
  ]);

  return items
    .map((item) => mapInHouse(item.id, item.fields, workforceIds))
    .filter((row): row is CustomerInHouseRecord => row !== null);
}
