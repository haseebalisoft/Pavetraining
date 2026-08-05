/**
 * Pure logic for syncing a Workforce record into the "Training Matrix Update"
 * list (schema key `trainingMatrixExample`).
 *
 * Dependency-free (no server-only / Graph imports) so scripts can exercise the
 * matching, dedupe, id-guard, and blank-preservation rules under Node. The
 * Graph I/O wrapper lives in trainingMatrixExampleService.syncWorkforceToTrainingMatrix,
 * which composes these helpers.
 *
 * Client rules encoded here:
 *  - Reuse an existing matrix row for a candidate — never create a duplicate.
 *  - Never use a synthetic id (e.g. `workforce-only:<id>`) as a SharePoint
 *    update id; treat it as "no real row yet" so the caller creates one.
 *  - Never overwrite existing expiry/training columns with blanks: only
 *    non-blank values are placed in the source, and absent keys are left
 *    untouched by the upsert.
 */

/** Minimum shape needed from a Workforce record to seed a matrix row. */
export interface WorkforceMatrixProfile {
  id?: string | null;
  candidateName: string;
  companyName?: string | null;
  workforceNumber?: string | null;
  dateOfBirth?: string | null;
  department?: string | null;
  trainingManager?: string | null;
  supervisor?: string | null;
  cscsNumber?: string | null;
  cscsExpiry?: string | null;
  eusrNumber?: string | null;
  eusrExpiry?: string | null;
  swqrNumber?: string | null;
  swqrExpiry?: string | null;
  nporsNumbers?: string | null;
}

/** Minimal existing-row reference used for matching. */
export interface MatrixRowRef {
  id: string;
  candidateName: string;
  dateOfBirth?: string | null;
}

const SYNTHETIC_MATRIX_PREFIX = "workforce-only:";
const EXAMPLE_MATRIX_PREFIX = "example:";

function clean(value: string | null | undefined): string | null {
  if (value == null) return null;
  const text = String(value).trim();
  return text ? text : null;
}

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

/** A matrix id the admin UI fabricates for a candidate that has no real row. */
export function isSyntheticMatrixId(id: string | null | undefined): boolean {
  return Boolean(id && id.trim().startsWith(SYNTHETIC_MATRIX_PREFIX));
}

/**
 * Resolve a matrix id down to a REAL SharePoint item id, or null when there is
 * none to PATCH (blank, or a synthetic `workforce-only:` id). Strips the
 * `example:` display prefix. This is the guard that stops fake ids ever being
 * sent to SharePoint as update targets.
 */
export function realMatrixItemId(id: string | null | undefined): string | null {
  const text = clean(id);
  if (!text) return null;
  if (text.startsWith(SYNTHETIC_MATRIX_PREFIX)) return null;
  if (text.startsWith(EXAMPLE_MATRIX_PREFIX)) {
    return clean(text.slice(EXAMPLE_MATRIX_PREFIX.length));
  }
  return text;
}

/**
 * Find the existing Training Matrix Update row for a candidate. The list has no
 * company column, so matching is by candidate name; when several rows share a
 * name, disambiguate by DOB, otherwise reuse the first (stable) so we never
 * create a duplicate.
 */
export function findMatrixRowForCandidate<T extends MatrixRowRef>(
  rows: T[],
  profile: Pick<WorkforceMatrixProfile, "candidateName" | "dateOfBirth">,
): T | null {
  const key = nameKey(profile.candidateName);
  if (!key) return null;
  const matches = rows.filter((row) => nameKey(row.candidateName) === key);
  if (matches.length === 0) return null;
  if (matches.length === 1) return matches[0]!;

  const dob = (clean(profile.dateOfBirth) ?? "").slice(0, 10);
  if (dob) {
    const byDob = matches.find(
      (row) => (clean(row.dateOfBirth) ?? "").slice(0, 10) === dob,
    );
    if (byDob) return byDob;
  }
  return matches[0]!;
}

/**
 * Build the upsert payload for a workforce record.
 *  - `source` carries the matrix DATE columns (Name always; DOB + CSCS/EUSR/
 *    NRSWA expiry only when non-blank). Omitting a key means the upsert leaves
 *    that column untouched — so existing expiry data is never wiped by blanks.
 *  - `profileFields` carries text profile columns (Company, Workforce Number,
 *    Department, Training Manager, Supervisor, NPORS/EUSR/SWQR/CSCS numbers).
 *    Only non-blank values are included, and the upsert writes each only if the
 *    list actually has that column ("if fields exist").
 */
export function buildWorkforceMatrixSource(profile: WorkforceMatrixProfile): {
  source: Record<string, string | null>;
  profileFields: Record<string, string | null>;
} {
  const source: Record<string, string | null> = {
    Name: profile.candidateName.trim(),
  };
  const dob = clean(profile.dateOfBirth);
  if (dob) source.DOB = dob;
  const cscsExpiry = clean(profile.cscsExpiry);
  if (cscsExpiry) source["CSCS Expiry"] = cscsExpiry;
  const eusrExpiry = clean(profile.eusrExpiry);
  if (eusrExpiry) source["EUSR Expiry"] = eusrExpiry;
  const swqrExpiry = clean(profile.swqrExpiry);
  if (swqrExpiry) source["NRSWA Expiry"] = swqrExpiry;

  const profileFields: Record<string, string | null> = {};
  const add = (header: string, value: string | null | undefined) => {
    const v = clean(value);
    if (v) profileFields[header] = v;
  };
  // Aliases cover both likely SharePoint display names / casings; the upsert
  // resolves each against the live column map and skips those that don't exist.
  add("Company", profile.companyName);
  add("Company Name", profile.companyName);
  add("Workforce Number", profile.workforceNumber);
  add("Department", profile.department);
  add("Training Manager", profile.trainingManager);
  add("Training manager", profile.trainingManager);
  add("Supervisor", profile.supervisor);
  add("NPORS Number", profile.nporsNumbers);
  add("NPORS Numbers", profile.nporsNumbers);
  add("EUSR Number", profile.eusrNumber);
  add("SWQR Number", profile.swqrNumber);
  add("CSCS Number", profile.cscsNumber);

  return { source, profileFields };
}
