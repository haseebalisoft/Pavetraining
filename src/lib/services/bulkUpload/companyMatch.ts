/**
 * Company matching for WORKFORCE bulk import.
 *
 * Client rule: match a workforce row to its company by Company Number FIRST,
 * and only fall back to Company Name when it uniquely and safely identifies one
 * company. Never assign by fuzzy name when a Company Number is present, and
 * never silently create a duplicate company from a spelling variation.
 *
 * Only depends on the sibling normalization helpers (no server-only or Graph
 * imports), so scripts/test-company-match.mjs can exercise it under Node with a
 * lightweight "@/" resolver hook.
 */
import { nameKey, normalizeCompanyKey } from "@/lib/services/bulkUpload/matching";

export const AMBIGUOUS_COMPANY_NAME_ERROR =
  "Company name is ambiguous. Use Company Number.";

export type CompanyMatchRef = {
  id: string;
  companyName: string;
  companyNumber: string | null;
};

/** Report fields surfaced per row (requirement 7). */
export type CompanyMatchReport = {
  matchedCompanyId: string | null;
  matchedCompanyNumber: string | null;
  matchedCompanyName: string | null;
  warning: string | null;
  error: string | null;
};

export type CompanyMatchResult =
  | {
      kind: "matched";
      company: CompanyMatchRef;
      matchedBy: "companyNumber" | "companyName";
      report: CompanyMatchReport;
    }
  /** Not found, and auto-create mode is enabled — caller should create it. */
  | { kind: "create"; report: CompanyMatchReport }
  /** Row must fail (wrong number, ambiguous name, or missing without auto-create). */
  | { kind: "error"; report: CompanyMatchReport };

function makeReport(partial: Partial<CompanyMatchReport>): CompanyMatchReport {
  return {
    matchedCompanyId: partial.matchedCompanyId ?? null,
    matchedCompanyNumber: partial.matchedCompanyNumber ?? null,
    matchedCompanyName: partial.matchedCompanyName ?? null,
    warning: partial.warning ?? null,
    error: partial.error ?? null,
  };
}

function dedupeById(rows: CompanyMatchRef[]): CompanyMatchRef[] {
  const seen = new Set<string>();
  const out: CompanyMatchRef[] = [];
  for (const row of rows) {
    if (seen.has(row.id)) continue;
    seen.add(row.id);
    out.push(row);
  }
  return out;
}

function namesLooselyEqual(a: string, b: string): boolean {
  return (
    nameKey(a) === nameKey(b) ||
    normalizeCompanyKey(a) === normalizeCompanyKey(b)
  );
}

export function resolveWorkforceCompanyMatch(
  companies: CompanyMatchRef[],
  input: { companyNumber?: string | null; companyName?: string | null },
  options: { autoCreateMissing: boolean },
): CompanyMatchResult {
  const number = input.companyNumber?.trim() ?? "";
  const name = input.companyName?.trim() ?? "";

  // ---- Priority 1: Company Number (authoritative; name is never used to
  // override a supplied number) ----
  if (number) {
    const numberKey = nameKey(number);
    const byNumber = companies.filter(
      (c) => nameKey(c.companyNumber) === numberKey,
    );

    if (byNumber.length === 1) {
      const company = byNumber[0];
      const base = {
        matchedCompanyId: company.id,
        matchedCompanyNumber: company.companyNumber,
        matchedCompanyName: company.companyName,
      };
      if (name && !namesLooselyEqual(company.companyName, name)) {
        return {
          kind: "matched",
          company,
          matchedBy: "companyNumber",
          report: makeReport({
            ...base,
            warning: `Company Number ${number} belongs to "${company.companyName}", but the row's Company Name is "${name}". Matched by Company Number.`,
          }),
        };
      }
      return {
        kind: "matched",
        company,
        matchedBy: "companyNumber",
        report: makeReport(base),
      };
    }

    if (byNumber.length > 1) {
      return {
        kind: "error",
        report: makeReport({
          matchedCompanyNumber: number,
          error: `Company Number "${number}" matches more than one company — data needs cleanup.`,
        }),
      };
    }

    // Number supplied but not found in the Company List.
    if (options.autoCreateMissing) {
      return {
        kind: "create",
        report: makeReport({
          matchedCompanyNumber: number,
          matchedCompanyName: name || null,
          warning: `Company Number "${number}" not found — a new company will be created on import.`,
        }),
      };
    }
    return {
      kind: "error",
      report: makeReport({
        matchedCompanyNumber: number,
        error: `Company Number "${number}" was not found. Upload the company first or enable auto-create.`,
      }),
    };
  }

  // ---- Priority 2: Company Name, only if it uniquely identifies one company ----
  if (name) {
    const exact = companies.filter(
      (c) => nameKey(c.companyName) === nameKey(name),
    );
    const chosen = exact.length
      ? exact
      : companies.filter(
          (c) => normalizeCompanyKey(c.companyName) === normalizeCompanyKey(name),
        );
    const unique = dedupeById(chosen);

    if (unique.length === 1) {
      const company = unique[0];
      return {
        kind: "matched",
        company,
        matchedBy: "companyName",
        report: makeReport({
          matchedCompanyId: company.id,
          matchedCompanyNumber: company.companyNumber,
          matchedCompanyName: company.companyName,
        }),
      };
    }

    if (unique.length > 1) {
      return {
        kind: "error",
        report: makeReport({
          matchedCompanyName: name,
          error: AMBIGUOUS_COMPANY_NAME_ERROR,
        }),
      };
    }

    // No name match.
    if (options.autoCreateMissing) {
      return {
        kind: "create",
        report: makeReport({
          matchedCompanyName: name,
          warning: `Company "${name}" not found — a new company will be created on import.`,
        }),
      };
    }
    return {
      kind: "error",
      report: makeReport({
        matchedCompanyName: name,
        error: `Company "${name}" was not found. Upload the company first, provide a Company Number, or enable auto-create.`,
      }),
    };
  }

  return {
    kind: "error",
    report: makeReport({ error: "Company Name or Company Number is required." }),
  };
}
