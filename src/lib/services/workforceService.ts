import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asNullableString,
  asString,
  buildSchemaFieldEqualsFilter,
  getListItemByKey,
  getListItemsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import type { WorkforceCandidate } from "@/types/models";

const workforceFields = getSharePointFields("workforce");

function mapWorkforceItem(
  id: string,
  fields: SharePointFields,
): WorkforceCandidate | null {
  const candidateName = asString(fields[workforceFields.candidateName]);
  const companyName = asString(fields[workforceFields.companyName]);

  if (!candidateName || !companyName) {
    return null;
  }

  return {
    id,
    candidateName,
    companyName,
    workforceNumber: asNullableString(fields[workforceFields.workforceNumber]),
    dateOfBirth: asNullableString(fields[workforceFields.dateOfBirth]),
    department: asNullableString(fields[workforceFields.department]),
    status: asNullableString(fields[workforceFields.status]),
    trainingManager: asNullableString(fields[workforceFields.trainingManager]),
    supervisor: asNullableString(fields[workforceFields.supervisor]),
    cscsNumber: asNullableString(fields[workforceFields.cscsNumber]),
    swqrNumber: asNullableString(fields[workforceFields.swqrNumber]),
    eusrNumber: asNullableString(fields[workforceFields.eusrNumber]),
    nporsNumbers: asNullableString(fields[workforceFields.nporsNumbers]),
    inHouseCertificationNumber: asNullableString(
      fields[workforceFields.inHouseCertificationNumber],
    ),
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

export async function getWorkforceByCompanyName(
  companyName: string,
): Promise<WorkforceCandidate[]> {
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
}

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
