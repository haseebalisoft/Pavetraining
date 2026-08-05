/**
 * Company Number integrity rules for bulk import (client rule: company numbers
 * must never be lost, overwritten, or duplicated).
 *
 * Kept dependency-free so scripts/test-company-number-rules.mjs can compile and
 * exercise it standalone.
 */

export const COMPANY_NUMBER_CHANGE_ERROR =
  "Company Number cannot be changed for an existing company.";

function companyNumberKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

export type CompanyNumberResolution =
  /** Row left the number blank; the matched company already has one — keep it. */
  | { kind: "preserve"; companyNumber: string }
  /** Row's number equals the matched company's number (case/space-insensitive). */
  | { kind: "match"; companyNumber: string }
  /** Matched company has no number yet; incoming (or a fresh allocation) may be assigned. */
  | { kind: "assign"; companyNumber: string | null }
  /** Row tries to change an existing number — there is no safe renumber mode. */
  | { kind: "reject"; message: string };

export function resolveCompanyNumberOnUpdate(input: {
  incoming: string | null | undefined;
  existing: string | null | undefined;
}): CompanyNumberResolution {
  const incoming = input.incoming?.trim() ?? "";
  const existing = input.existing?.trim() ?? "";

  if (!existing) {
    return { kind: "assign", companyNumber: incoming || null };
  }
  if (!incoming) {
    return { kind: "preserve", companyNumber: existing };
  }
  if (companyNumberKey(incoming) === companyNumberKey(existing)) {
    return { kind: "match", companyNumber: existing };
  }
  return { kind: "reject", message: COMPANY_NUMBER_CHANGE_ERROR };
}

export function companyNumberTakenError(
  number: string,
  ownerId: string,
): string {
  return `Company Number "${number}" already belongs to another company (#${ownerId}).`;
}

export function companyNumberDuplicateInFileError(number: string): string {
  return `Duplicate Company Number "${number}" appears more than once in this file.`;
}

/** A minimal view of an existing company for import resolution. */
export type CompanyImportMatch = {
  id: string;
  companyNumber: string | null;
} | null;

export type CompanyImportDecision =
  | { action: "reject"; message: string }
  | { action: "skip"; targetId: string }
  /** Update the matched company; `companyNumber` is the value to KEEP (never a new one). */
  | { action: "update"; targetId: string; companyNumber: string | null }
  /** Create a new company; `companyNumber` null means "allocate a fresh portal number". */
  | { action: "create"; companyNumber: string | null };

/**
 * The single decision the importer makes per row, covering:
 *  - update mode: preserve a blank number, reject a changed number;
 *  - reject a supplied number already owned by a different SharePoint company;
 *  - create mode: keep a unique supplied number, or allocate for a brand-new company.
 *
 * In-file duplicate-number rejection is layered on top by the caller, which
 * tracks the final numbers already used earlier in the same file.
 *
 * @param matchByNumber existing company whose number equals the supplied number
 * @param matchByName   existing company matched by name
 */
export function resolveCompanyImport(input: {
  incoming: string | null | undefined;
  matchByNumber: CompanyImportMatch;
  matchByName: CompanyImportMatch;
  duplicateMode: "skip" | "update" | "create";
}): CompanyImportDecision {
  const incoming = input.incoming?.trim() ?? "";
  const { matchByNumber, matchByName, duplicateMode } = input;

  // A supplied number that belongs to a company OTHER than the one this row
  // targets by name is a duplicate against SharePoint — never reusable.
  if (
    incoming &&
    matchByNumber &&
    (!matchByName || matchByName.id !== matchByNumber.id)
  ) {
    return {
      action: "reject",
      message: companyNumberTakenError(incoming, matchByNumber.id),
    };
  }

  // The existing company this row refers to (prefer the name match).
  const target = matchByName ?? matchByNumber;

  if (!target) {
    // Brand-new company. A supplied number is unique here (the check above
    // proved it is not owned by anyone); blank means "allocate".
    return { action: "create", companyNumber: incoming || null };
  }

  const sub = resolveCompanyNumberOnUpdate({
    incoming,
    existing: target.companyNumber,
  });
  if (sub.kind === "reject") {
    return { action: "reject", message: sub.message };
  }
  if (duplicateMode === "skip") {
    return { action: "skip", targetId: target.id };
  }
  if (duplicateMode === "create") {
    // Create-despite-duplicate: a NEW company with a freshly allocated number,
    // never the matched company's number.
    return { action: "create", companyNumber: null };
  }
  const keep =
    sub.kind === "assign"
      ? (sub.companyNumber ?? target.companyNumber ?? null)
      : (target.companyNumber ?? null);
  return { action: "update", targetId: target.id, companyNumber: keep };
}
