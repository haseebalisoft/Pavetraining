import type { AdminWorkforceRecord } from "@/lib/services/adminCrudService";
import type { Company } from "@/types/models";

export function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

export function normalizeCompanyKey(value: string | null | undefined): string {
  return nameKey(value)
    .replace(/\bltd\b\.?/g, "")
    .replace(/\blimited\b/g, "")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function sameDate(
  a: string | null | undefined,
  b: string | null | undefined,
): boolean {
  if (!a?.trim() || !b?.trim()) return false;
  return nameKey(a).slice(0, 10) === nameKey(b).slice(0, 10);
}

export function findCompanyByName(
  companies: Company[],
  companyName: string | null | undefined,
): Company | null {
  if (!companyName?.trim()) return null;
  const exact = companies.find(
    (c) => nameKey(c.companyName) === nameKey(companyName),
  );
  if (exact) return exact;
  const normalized = normalizeCompanyKey(companyName);
  return (
    companies.find((c) => normalizeCompanyKey(c.companyName) === normalized) ??
    null
  );
}

export type CandidateDuplicateMatch =
  | { kind: "workforceNumber"; record: AdminWorkforceRecord }
  | { kind: "nameDobCompany"; record: AdminWorkforceRecord }
  | { kind: "nameCompany"; record: AdminWorkforceRecord }
  | null;

/**
 * Duplicate priority:
 * 1. WorkforceNumber + Company
 * 2. CandidateName + DOB + Company
 * 3. CandidateName + Company (warning-level only)
 */
export function findCandidateDuplicate(
  workforce: AdminWorkforceRecord[],
  input: {
    candidateName: string;
    companyName: string;
    workforceNumber: string | null;
    dateOfBirth: string | null;
  },
): CandidateDuplicateMatch {
  const companyExact = nameKey(input.companyName);
  const companyNorm = normalizeCompanyKey(input.companyName);
  const sameCompany = (row: AdminWorkforceRecord) =>
    nameKey(row.companyName) === companyExact ||
    normalizeCompanyKey(row.companyName) === companyNorm;

  if (input.workforceNumber?.trim()) {
    const numberKey = nameKey(input.workforceNumber);
    const byNumber = workforce.find(
      (w) => nameKey(w.workforceNumber) === numberKey && sameCompany(w),
    );
    if (byNumber) {
      return { kind: "workforceNumber", record: byNumber };
    }
  }

  const cName = nameKey(input.candidateName);
  if (input.dateOfBirth?.trim()) {
    const byNameDob = workforce.find(
      (w) =>
        nameKey(w.candidateName) === cName &&
        sameCompany(w) &&
        sameDate(w.dateOfBirth, input.dateOfBirth),
    );
    if (byNameDob) {
      return { kind: "nameDobCompany", record: byNameDob };
    }
  }

  const byNameCompany = workforce.find(
    (w) => nameKey(w.candidateName) === cName && sameCompany(w),
  );
  if (byNameCompany) {
    return { kind: "nameCompany", record: byNameCompany };
  }

  return null;
}
