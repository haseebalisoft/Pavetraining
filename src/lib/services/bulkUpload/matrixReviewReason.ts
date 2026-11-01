/**
 * Plain-English reasons a Training Matrix row could not be linked to exactly
 * one Workforce record ("Needs Review" / "Orphan"): DOB mismatch, candidate
 * not found, multiple matches, company mismatch, …
 *
 * Used by the matrix upload (preview + results) and by Admin Matrix. Kept
 * dependency-free (no server-only / Graph imports) so scripts can test it —
 * see scripts/test-matrix-review-reasons.mjs.
 *
 * Reasons name the Workforce records that were compared (name, DOB, company),
 * so they are ADMIN-ONLY — never attach them to customer-facing data.
 */

import { nameKey, normalizeCompanyKey } from "@/lib/services/bulkUpload/matching";
import {
  candidateNameKey,
  isoDateKey,
} from "@/lib/services/bulkUpload/workforceMatrixSync";
import {
  MATRIX_REVIEW_REASON_LABELS,
  type MatrixReviewReason,
  type MatrixReviewReasonCode,
} from "@/types/bulkUpload";

/** The Workforce fields a reason quotes back to the admin. */
export interface ReasonPerson {
  candidateName: string;
  dateOfBirth: string | null;
  companyName: string | null;
  workforceNumber?: string | null;
}

export interface MatchableWorkforce extends ReasonPerson {
  companyNumber?: string | null;
}

export function matrixReviewReason(
  code: MatrixReviewReasonCode,
  detail: string,
): MatrixReviewReason {
  return { code, label: MATRIX_REVIEW_REASON_LABELS[code], detail };
}

/** `DD/MM/YYYY` (the format the spreadsheet and UI use), or the raw value. */
export function formatReasonDate(value: string | null | undefined): string {
  const iso = isoDateKey(value);
  if (!iso) return String(value ?? "").trim();
  const [year, month, day] = iso.split("-");
  return `${day}/${month}/${year}`;
}

const MAX_LISTED = 3;

/** "Zaid Khan (DOB 02/02/1980, Acme Ltd, Workforce # W12); …; and 2 more" */
export function describePeople(people: ReasonPerson[]): string {
  const shown = people.slice(0, MAX_LISTED).map((person) => {
    const dob = formatReasonDate(person.dateOfBirth);
    const bits = [
      dob ? `DOB ${dob}` : "no DOB",
      person.companyName?.trim() || "no company",
    ];
    if (person.workforceNumber?.trim()) {
      bits.push(`Workforce # ${person.workforceNumber.trim()}`);
    }
    return `${person.candidateName.trim() || "(no name)"} (${bits.join(", ")})`;
  });
  const extra = people.length - shown.length;
  return shown.join("; ") + (extra > 0 ? `; and ${extra} more` : "");
}

function recordsNamed(count: number, name: string): string {
  return count === 1
    ? `1 Workforce record is named "${name}"`
    : `${count} Workforce records are named "${name}"`;
}

/**
 * Explain why same-name Workforce records did not resolve to one person,
 * comparing DOB first, then company. Shared by the upload and Admin Matrix.
 */
export function explainSameNameCandidates(input: {
  candidateName: string;
  dateOfBirth: string | null;
  companyName: string | null;
  sameName: ReasonPerson[];
}): { code: MatrixReviewReasonCode; text: string } {
  const name = input.candidateName.trim();
  const { sameName } = input;
  const rowDob = isoDateKey(input.dateOfBirth);
  if (!rowDob) {
    return {
      code: "dobMissing",
      text: `${recordsNamed(sameName.length, name)} and this row has no DOB to confirm which person it is — a name alone is never used to link: ${describePeople(sameName)}.`,
    };
  }
  const dobMatches = sameName.filter(
    (person) => isoDateKey(person.dateOfBirth) === rowDob,
  );
  if (!dobMatches.length) {
    const target =
      sameName.length === 1
        ? `the Workforce record named "${name}"`
        : `any of the ${sameName.length} Workforce records named "${name}"`;
    return {
      code: "dobMismatch",
      text: `This row's DOB (${formatReasonDate(input.dateOfBirth)}) doesn't match ${target}: ${describePeople(sameName)}.`,
    };
  }
  if (dobMatches.length === 1) {
    return {
      code: "multipleMatches",
      text:
        sameName.length === 1
          ? `1 Workforce record matches Name + DOB: ${describePeople(dobMatches)}.`
          : `${recordsNamed(sameName.length, name)}; only ${describePeople(dobMatches)} has the same DOB.`,
    };
  }
  const company = input.companyName?.trim() || null;
  if (company) {
    const companyKey = normalizeCompanyKey(company);
    const atCompany = dobMatches.filter(
      (person) => normalizeCompanyKey(person.companyName) === companyKey,
    );
    if (!atCompany.length) {
      return {
        code: "companyMismatch",
        text: `${dobMatches.length} Workforce records match Name + DOB, but none is at "${company}": ${describePeople(dobMatches)}.`,
      };
    }
    if (atCompany.length === 1) {
      return {
        code: "multipleMatches",
        text: `${dobMatches.length} Workforce records match Name + DOB; only ${describePeople(atCompany)} is at "${company}".`,
      };
    }
  }
  return {
    code: "multipleMatches",
    text: `${dobMatches.length} Workforce records match Name + DOB: ${describePeople(dobMatches)}.`,
  };
}

export type WorkforceMatchType =
  | "workforceNumber"
  | "companyNameDob"
  | "nameDob"
  | "none";

export interface WorkforceRowMatch<T> {
  candidate: T | null;
  matchType: WorkforceMatchType;
  /** Several Workforce records tied — never auto-link, leave Needs Review. */
  ambiguous: boolean;
  /** Set whenever `candidate` is null: exactly why no record was chosen. */
  reason: MatrixReviewReason | null;
}

/**
 * Resolve the Workforce record a matrix spreadsheet row belongs to.
 *
 * Locked ladder (client rule): Workforce Number → Company + Name + DOB →
 * Name + DOB when unique. Name alone NEVER auto-matches — two candidates can
 * share a name, and attaching one person's training to the other is worse than
 * leaving the row for an admin to resolve. Every miss carries a reason.
 */
export function matchWorkforceForMatrixRow<T extends MatchableWorkforce>(
  workforce: T[],
  fields: Record<string, string | null>,
): WorkforceRowMatch<T> {
  const noMatch = (
    code: MatrixReviewReasonCode,
    detail: string,
    ambiguous = false,
  ): WorkforceRowMatch<T> => ({
    candidate: null,
    matchType: "none",
    ambiguous,
    reason: matrixReviewReason(code, detail),
  });

  // Prefixed to name-based reasons so the admin also learns the number missed.
  let numberNote = "";
  const sheetNumber = fields.workforceNumber?.trim();
  if (sheetNumber) {
    const numberKey = nameKey(sheetNumber);
    const byNumber = workforce.filter(
      (row) => nameKey(row.workforceNumber) === numberKey,
    );
    if (byNumber.length === 1) {
      return {
        candidate: byNumber[0]!,
        matchType: "workforceNumber",
        ambiguous: false,
        reason: null,
      };
    }
    if (byNumber.length > 1) {
      return noMatch(
        "multipleMatches",
        `Workforce Number "${sheetNumber}" is on ${byNumber.length} Workforce records: ${describePeople(byNumber)}. Fix the duplicate number in Workforce, then re-upload.`,
        true,
      );
    }
    numberNote = `Workforce Number "${sheetNumber}" is not on any Workforce record. `;
  }

  const name = fields.candidateName?.trim() ?? "";
  const cName = candidateNameKey(name);
  if (!cName) {
    return noMatch(
      "candidateNotFound",
      `${numberNote}The row has no Candidate Name to match by.`,
    );
  }

  const sameName = workforce.filter(
    (row) => candidateNameKey(row.candidateName) === cName,
  );
  const rowDob = isoDateKey(fields.dateOfBirth);
  const sheetCompany =
    fields.company?.trim() ||
    (fields.companyNumber?.trim()
      ? `company number ${fields.companyNumber.trim()}`
      : null);

  if (!sameName.length) {
    // A same-DOB record at the same company is usually a spelling difference.
    let hint = "";
    if (rowDob) {
      const companyNorm = fields.company?.trim()
        ? normalizeCompanyKey(fields.company)
        : null;
      const sameDob = workforce.filter(
        (row) =>
          isoDateKey(row.dateOfBirth) === rowDob &&
          (!companyNorm || normalizeCompanyKey(row.companyName) === companyNorm),
      );
      if (sameDob.length >= 1 && sameDob.length <= MAX_LISTED) {
        hint = ` Possible spelling difference: ${describePeople(sameDob)} ${sameDob.length === 1 ? "has" : "have"} the same DOB${companyNorm ? ` at "${fields.company!.trim()}"` : ""}.`;
      }
    }
    return noMatch(
      "candidateNotFound",
      `${numberNote}No Workforce record is named "${name}".${hint}`,
    );
  }

  // No DOB means Name is the only key left, and Name alone must not auto-match.
  if (!rowDob) {
    const why = explainSameNameCandidates({
      candidateName: name,
      dateOfBirth: null,
      companyName: sheetCompany,
      sameName,
    });
    return noMatch(why.code, `${numberNote}${why.text} Add the DOB to the spreadsheet.`);
  }

  const nameDobMatches = sameName.filter(
    (row) => isoDateKey(row.dateOfBirth) === rowDob,
  );
  if (!nameDobMatches.length) {
    const why = explainSameNameCandidates({
      candidateName: name,
      dateOfBirth: fields.dateOfBirth,
      companyName: sheetCompany,
      sameName,
    });
    return noMatch(
      why.code,
      `${numberNote}${why.text} Correct the DOB in the spreadsheet or in Workforce.`,
    );
  }

  // Company narrows a same-name + same-DOB tie (the client's primary key).
  const companyHint = fields.company?.trim() ? nameKey(fields.company) : null;
  const companyNorm = fields.company?.trim()
    ? normalizeCompanyKey(fields.company)
    : null;
  const companyNumberHint = fields.companyNumber?.trim()
    ? nameKey(fields.companyNumber)
    : null;

  if (companyHint || companyNumberHint) {
    const scoped = nameDobMatches.filter(
      (row) =>
        (companyNumberHint != null &&
          nameKey(row.companyNumber) === companyNumberHint) ||
        (companyHint != null &&
          (nameKey(row.companyName) === companyHint ||
            normalizeCompanyKey(row.companyName) === companyNorm)),
    );
    if (scoped.length === 1) {
      return {
        candidate: scoped[0]!,
        matchType: "companyNameDob",
        ambiguous: false,
        reason: null,
      };
    }
    if (scoped.length > 1) {
      return noMatch(
        "multipleMatches",
        `${numberNote}${scoped.length} Workforce records match Name + DOB + Company: ${describePeople(scoped)}. Add the Workforce Number to the spreadsheet to pick one.`,
        true,
      );
    }
  }

  if (nameDobMatches.length === 1) {
    return {
      candidate: nameDobMatches[0]!,
      matchType: "nameDob",
      ambiguous: false,
      reason: null,
    };
  }

  // Same Name AND same DOB in more than one company: Company is required to
  // link safely, and the spreadsheet either had none or one matching nobody.
  if (sheetCompany) {
    return noMatch(
      "companyMismatch",
      `${numberNote}${nameDobMatches.length} Workforce records match Name + DOB, but none is at "${sheetCompany}" (the company on the spreadsheet): ${describePeople(nameDobMatches)}. Correct the Company on the spreadsheet, or add the Workforce Number.`,
      true,
    );
  }
  return noMatch(
    "multipleMatches",
    `${numberNote}${nameDobMatches.length} Workforce records match Name + DOB: ${describePeople(nameDobMatches)}. Add Company or Workforce Number to the spreadsheet to pick one.`,
    true,
  );
}

/**
 * Why Admin Matrix shows a row as Needs Review / Orphan. Mirrors
 * `deriveMatrixLinkStatus` (name-keyed), then compares DOB and company so the
 * admin sees the exact difference. Null for Linked rows.
 */
export function explainMatrixLinkStatus(input: {
  status: "Linked" | "Needs Review" | "Orphan";
  candidateName: string;
  dateOfBirth: string | null;
  companyName: string | null;
  /** Stored WorkforceItemId that no longer resolves to a live record. */
  staleWorkforceItemId: string | null;
  /** Live Workforce records sharing this row's name (within any company filter). */
  sameName: ReasonPerson[];
  /** Set when the same-name person is already linked to another matrix row. */
  duplicateOf: { person: ReasonPerson; matrixRowId: string | null } | null;
  /** Company the list was scoped to, if any. */
  scopedCompany: string | null;
}): MatrixReviewReason | null {
  if (input.status === "Linked") return null;
  const name = input.candidateName.trim() || "(no name)";
  const action = " Use Link to Workforce on the row to choose the right record.";

  if (input.duplicateOf) {
    const other = input.duplicateOf.matrixRowId
      ? ` (#${input.duplicateOf.matrixRowId})`
      : "";
    return matrixReviewReason(
      "duplicateMatrixRow",
      `${describePeople([input.duplicateOf.person])} is already linked to another Training Matrix row${other}. This row is a second copy — delete it, or use Link to Workforce if it is a different person.`,
    );
  }

  const stale = input.staleWorkforceItemId
    ? `The Workforce record this row was linked to (#${input.staleWorkforceItemId}) no longer exists — it was deleted or re-created. `
    : "";

  if (!input.sameName.length) {
    const where = input.scopedCompany ? ` at "${input.scopedCompany}"` : "";
    return matrixReviewReason(
      stale ? "linkedRecordMissing" : "candidateNotFound",
      `${stale}No Workforce record is named "${name}"${where}. Add the candidate to Workforce, or delete this row if they have left.`,
    );
  }

  const why = explainSameNameCandidates({
    candidateName: name,
    dateOfBirth: input.dateOfBirth,
    companyName: input.companyName,
    sameName: input.sameName,
  });
  return matrixReviewReason(
    stale ? "linkedRecordMissing" : why.code,
    `${stale}${why.text}${action}`,
  );
}
