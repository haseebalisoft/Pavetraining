/**
 * Re-link a Training Matrix row after an admin corrects a candidate's name or
 * DOB (on the Matrix row or on the Workforce record).
 *
 * Decides — without touching SharePoint — what the correction means:
 *  - `alreadyLinked` — the row is linked to a live Workforce record; a
 *    corrected display name/DOB never moves an id link;
 *  - `link`          — exactly one Workforce record now matches and it owns
 *    no other row: link this row in place (expiry cells untouched);
 *  - `merge`         — the match already owns a linked row: fold this row's
 *    expiry dates into that row, then remove this copy (no duplicate rows);
 *  - `unmatched`     — still no single match: the row stays Needs Review and
 *    the result carries the exact reason.
 *
 * Dependency-free so scripts can test it — see
 * scripts/test-matrix-review-reasons.mjs. adminCrudService applies the plan.
 */

import {
  matchWorkforceForMatrixRow,
  type MatchableWorkforce,
} from "@/lib/services/bulkUpload/matrixReviewReason";
import { isoDateKey } from "@/lib/services/bulkUpload/workforceMatrixSync";
import type { MatrixReviewReason } from "@/types/bulkUpload";

/** The Matrix row fields the plan reads (a TrainingMatrixExampleRow subset). */
export interface RelinkMatrixRow {
  id: string;
  candidateName: string;
  dateOfBirth: string | null;
  workforceItemId: string | null;
  companyNumber: string | null;
  storedCompanyName: string | null;
  columnValues: Record<string, string | null>;
  manualOverrides: string[];
  categoryTrainingDates?: Record<string, string>;
}

export type MatrixRelinkPlan<T> =
  | { action: "alreadyLinked"; candidate: T }
  | { action: "link"; candidate: T }
  | { action: "merge"; candidate: T; target: RelinkMatrixRow }
  | { action: "unmatched"; reason: MatrixReviewReason };

export function planMatrixRelink<T extends MatchableWorkforce & { id: string }>(
  row: RelinkMatrixRow,
  workforce: T[],
  matrixRows: RelinkMatrixRow[],
): MatrixRelinkPlan<T> {
  const linkedId = String(row.workforceItemId ?? "").trim();
  const owner = linkedId
    ? workforce.find((candidate) => candidate.id === linkedId)
    : undefined;
  if (owner) return { action: "alreadyLinked", candidate: owner };

  // Name + DOB (+ Company) only. A stale row's old Workforce Number is NOT
  // used: a deleted record's number can since belong to someone else.
  const match = matchWorkforceForMatrixRow(workforce, {
    candidateName: row.candidateName,
    dateOfBirth: row.dateOfBirth,
    company: row.storedCompanyName,
    companyNumber: row.companyNumber,
    workforceNumber: null,
  });
  if (!match.candidate) {
    return {
      action: "unmatched",
      reason: match.reason ?? {
        code: "candidateNotFound",
        label: "Candidate not found",
        detail: `No Workforce record matches "${row.candidateName}".`,
      },
    };
  }

  const candidate = match.candidate;
  const target = matrixRows.find(
    (other) =>
      other.id !== row.id &&
      String(other.workforceItemId ?? "").trim() === candidate.id,
  );
  return target
    ? { action: "merge", candidate, target }
    : { action: "link", candidate };
}

/** Columns that identify the person rather than hold training dates. */
const IDENTITY_HEADERS = new Set(["name", "dob"]);

export interface MatrixCellMerge {
  /** Only the cells to write onto the target (partial update). */
  cells: Record<string, string>;
  manualOverrides: string[];
  categoryTrainingDates: Record<string, string>;
  /** Headers whose date came from the duplicate row. */
  mergedHeaders: string[];
  /** Training-date keys filled from the duplicate row. */
  mergedTrainingDates: string[];
}

/**
 * Fold the duplicate row's dates into the target without losing any:
 *  - a target cell marked Manual is never changed;
 *  - a blank target cell takes the duplicate's date;
 *  - when both have a date the later expiry wins (same rule as register
 *    sync: a newer Pass extends, nothing ever shortens an expiry);
 *  - training dates only fill blanks.
 * `allowedHeaders` limits writes to columns the list actually stores.
 */
export function mergeMatrixCells(
  target: RelinkMatrixRow,
  duplicate: RelinkMatrixRow,
  allowedHeaders: readonly string[],
): MatrixCellMerge {
  const allowed = new Set(allowedHeaders);
  const targetManual = new Set(target.manualOverrides ?? []);
  const duplicateManual = new Set(duplicate.manualOverrides ?? []);
  const cells: Record<string, string> = {};
  const mergedHeaders: string[] = [];
  const manual = new Set(targetManual);

  for (const [header, value] of Object.entries(duplicate.columnValues ?? {})) {
    if (!allowed.has(header) || IDENTITY_HEADERS.has(header.toLowerCase())) {
      continue;
    }
    const incoming = isoDateKey(value);
    if (!incoming || targetManual.has(header)) continue;
    const existing = isoDateKey(target.columnValues?.[header]);
    if (existing && existing >= incoming) continue;
    cells[header] = incoming;
    mergedHeaders.push(header);
    if (duplicateManual.has(header)) manual.add(header);
  }

  const categoryTrainingDates: Record<string, string> = {
    ...(target.categoryTrainingDates ?? {}),
  };
  const mergedTrainingDates: string[] = [];
  for (const [key, value] of Object.entries(
    duplicate.categoryTrainingDates ?? {},
  )) {
    if (!value || categoryTrainingDates[key]) continue;
    categoryTrainingDates[key] = value;
    mergedTrainingDates.push(key);
  }

  return {
    cells,
    manualOverrides: [...manual],
    categoryTrainingDates,
    mergedHeaders,
    mergedTrainingDates,
  };
}
