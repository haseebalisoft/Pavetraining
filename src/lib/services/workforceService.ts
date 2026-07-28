import "server-only";

import { cache } from "react";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { getAllCompanies } from "@/lib/services/companyService";
import {
  asLookupOrString,
  asNullableString,
  asString,
  buildSchemaFieldEqualsFilter,
  extractLookupId,
  getListItemByKey,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import type { WorkforceCandidate } from "@/types/models";

const workforceFields = getSharePointFields("workforce");

function asDepartment(value: unknown): string | null {
  if (Array.isArray(value)) {
    const parts = value
      .map((entry) => {
        if (typeof entry === "string") return entry.trim();
        if (entry && typeof entry === "object") {
          const record = entry as { LookupValue?: unknown; Label?: unknown };
          return (
            asString(record.LookupValue) ?? asString(record.Label) ?? ""
          ).trim();
        }
        return "";
      })
      .filter(Boolean);
    return parts.length ? parts.join(", ") : null;
  }
  return asLookupOrString(value) ?? asNullableString(value);
}

function mapWorkforceItem(
  id: string,
  fields: SharePointFields,
  companyNameById?: Map<string, string>,
  permissionNameById?: Map<string, string>,
  departmentNameById?: Map<string, string>,
): WorkforceCandidate | null {
  const candidateName = asString(fields[workforceFields.candidateName]);
  const companyLookupId = extractLookupId(fields, workforceFields.companyName);
  const companyName =
    asLookupOrString(fields[workforceFields.companyName]) ??
    asString(fields[workforceFields.companyName]) ??
    (companyLookupId && companyNameById
      ? (companyNameById.get(companyLookupId) ?? undefined)
      : undefined);

  if (!candidateName || !companyName) {
    return null;
  }

  const trainingManagerLookupId = extractLookupId(
    fields,
    workforceFields.trainingManager,
  );
  const supervisorLookupId = extractLookupId(
    fields,
    workforceFields.supervisor,
  );
  const departmentLookupId = extractLookupId(
    fields,
    workforceFields.departmentText,
  );

  return {
    id,
    candidateName,
    companyName,
    workforceNumber: asNullableString(fields[workforceFields.workforceNumber]),
    dateOfBirth: asNullableString(fields[workforceFields.dateOfBirth]),
    department:
      asDepartment(fields[workforceFields.departmentText]) ??
      (departmentLookupId && departmentNameById
        ? (departmentNameById.get(departmentLookupId) ?? null)
        : null) ??
      asDepartment(fields[workforceFields.department]),
    status: asNullableString(fields[workforceFields.status]),
    trainingManager:
      asLookupOrString(fields[workforceFields.trainingManager]) ??
      asNullableString(fields[workforceFields.trainingManager]) ??
      (trainingManagerLookupId && permissionNameById
        ? (permissionNameById.get(trainingManagerLookupId) ?? null)
        : null),
    supervisor:
      asLookupOrString(fields[workforceFields.supervisor]) ??
      asNullableString(fields[workforceFields.supervisor]) ??
      (supervisorLookupId && permissionNameById
        ? (permissionNameById.get(supervisorLookupId) ?? null)
        : null),
    email: asNullableString(fields[workforceFields.email])?.toLowerCase() ?? null,
    cscsNumber: asNullableString(fields[workforceFields.cscsNumber]),
    swqrNumber: asNullableString(fields[workforceFields.swqrNumber]),
    eusrNumber: asNullableString(fields[workforceFields.eusrNumber]),
    nporsNumbers: asNullableString(fields[workforceFields.nporsNumbers]),
    inHouseCertificationNumber: asNullableString(
      fields[workforceFields.inHouseCertificationNumber],
    ),
    cscsExpiry: asNullableString(fields[workforceFields.cscsExpiry]),
    swqrExpiry: asNullableString(fields[workforceFields.swqrExpiry]),
    eusrExpiry: asNullableString(fields[workforceFields.eusrExpiry]),
  };
}

async function companyNameLookupMap(): Promise<Map<string, string>> {
  const companies = await getAllCompanies();
  return new Map(companies.map((row) => [row.id, row.companyName] as const));
}

async function permissionNameLookupMap(): Promise<Map<string, string>> {
  const permissionFields = getSharePointFields("permissions");
  const items = await getListItemsByKey("permissions", { top: 5000 });
  const map = new Map<string, string>();
  for (const item of items) {
    const name =
      asNullableString(item.fields[permissionFields.name]) ??
      asNullableString(item.fields[permissionFields.userEmail]);
    if (name) map.set(item.id, name);
  }
  return map;
}

async function departmentNameLookupMap(): Promise<Map<string, string>> {
  const departmentFields = getSharePointFields("departments");
  const items = await getListItemsByKey("departments", { top: 5000 });
  const map = new Map<string, string>();
  for (const item of items) {
    const name =
      asNullableString(item.fields[departmentFields.name]) ??
      asNullableString(item.fields[departmentFields.title]);
    if (name) map.set(item.id, name);
  }
  return map;
}

export async function getWorkforceById(
  candidateId: string,
): Promise<WorkforceCandidate | null> {
  const item = await getListItemByKey("workforce", candidateId);
  if (!item) {
    return null;
  }

  const [companyNameById, permissionNameById, departmentNameById] =
    await Promise.all([
      companyNameLookupMap(),
      permissionNameLookupMap(),
      departmentNameLookupMap(),
    ]);
  return mapWorkforceItem(
    item.id,
    item.fields,
    companyNameById,
    permissionNameById,
    departmentNameById,
  );
}

/** Deduped per request so name-map + scope filters share one Graph read. */
export const getWorkforceByCompanyName = cache(
  async (companyName: string): Promise<WorkforceCandidate[]> => {
    const [companies, permissionNameById, departmentNameById] =
      await Promise.all([
        getAllCompanies(),
        permissionNameLookupMap(),
        departmentNameLookupMap(),
      ]);
    const companyNameById = new Map(
      companies.map((row) => [row.id, row.companyName] as const),
    );
    const company = companies.find(
      (row) =>
        row.companyName.trim().toLowerCase() ===
        companyName.trim().toLowerCase(),
    );

    // Prefer LookupId filter — Graph text filters on Lookup CompanyName are unreliable.
    const items = await getListItemsByKey("workforce", {
      filter: company
        ? `fields/CompanyNameLookupId eq ${Number(company.id)}`
        : buildSchemaFieldEqualsFilter("workforce", "companyName", companyName),
      top: 5000,
    });

    return items
      .map((item) =>
        mapWorkforceItem(
          item.id,
          item.fields,
          companyNameById,
          permissionNameById,
          departmentNameById,
        ),
      )
      .filter((row): row is WorkforceCandidate => row !== null)
      .filter(
        (row) =>
          row.companyName.trim().toLowerCase() ===
          companyName.trim().toLowerCase(),
      );
  },
);

/**
 * Builds a case-insensitive candidate-name → workforce id map for profile links.
 */
export async function getWorkforceIdByCandidateName(
  companyName: string,
): Promise<Map<string, string>> {
  const workforce = await getWorkforceByCompanyName(companyName);
  const map = new Map<string, string>();

  for (const candidate of workforce) {
    const key = candidate.candidateName.trim().toLowerCase();
    if (key && !map.has(key)) {
      map.set(key, candidate.id);
    }
  }

  return map;
}
