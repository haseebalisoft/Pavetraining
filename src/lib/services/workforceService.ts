import "server-only";

import { cache } from "react";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asLookupOrString,
  asNullableString,
  asString,
  buildSchemaFieldEqualsFilter,
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
): WorkforceCandidate | null {
  const candidateName = asString(fields[workforceFields.candidateName]);
  const companyName =
    asLookupOrString(fields[workforceFields.companyName]) ??
    asString(fields[workforceFields.companyName]);

  if (!candidateName || !companyName) {
    return null;
  }

  return {
    id,
    candidateName,
    companyName,
    workforceNumber: asNullableString(fields[workforceFields.workforceNumber]),
    dateOfBirth: asNullableString(fields[workforceFields.dateOfBirth]),
    department: asDepartment(fields[workforceFields.department]),
    status: asNullableString(fields[workforceFields.status]),
    trainingManager:
      asLookupOrString(fields[workforceFields.trainingManager]) ??
      asNullableString(fields[workforceFields.trainingManager]),
    supervisor:
      asLookupOrString(fields[workforceFields.supervisor]) ??
      asNullableString(fields[workforceFields.supervisor]),
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

export async function getWorkforceById(
  candidateId: string,
): Promise<WorkforceCandidate | null> {
  const item = await getListItemByKey("workforce", candidateId);
  if (!item) {
    return null;
  }

  return mapWorkforceItem(item.id, item.fields);
}

/** Deduped per request so name-map + scope filters share one Graph read. */
export const getWorkforceByCompanyName = cache(
  async (companyName: string): Promise<WorkforceCandidate[]> => {
    const items = await getListItemsByKey("workforce", {
      filter: buildSchemaFieldEqualsFilter(
        "workforce",
        "companyName",
        companyName,
      ),
      top: 5000,
    });

    return items
      .map((item) => mapWorkforceItem(item.id, item.fields))
      .filter((row): row is WorkforceCandidate => row !== null);
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
