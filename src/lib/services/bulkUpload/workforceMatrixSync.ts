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
  /** SharePoint Workforce item id — the strong matrix link key (defaults to `id`). */
  workforceItemId?: string | number | null;
  candidateName: string;
  companyName?: string | null;
  /** SharePoint Company item id (Company lookup id) — a link key. */
  companyItemId?: string | number | null;
  companyNumber?: string | null;
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

/** Typed link columns written onto a Training Matrix Update row. */
export interface MatrixLinkFields {
  /** Display name → numeric value (SharePoint Number columns). */
  numbers: Record<string, number>;
  /** Display name → text value (text or choice columns, e.g. MatrixLinkStatus). */
  text: Record<string, string>;
}

export type MatrixLinkStatus = "Linked" | "Orphan" | "Needs Review";

/** How a matrix row was matched to a workforce record. */
export type MatrixMatchType = "id" | "legacy" | "name" | "none";

/** Minimal existing-row reference used for matching. */
export interface MatrixRowRef {
  id: string;
  candidateName: string;
  dateOfBirth?: string | null;
  /** Link fields (present once a row has been linked to a workforce record). */
  workforceItemId?: string | number | null;
  workforceNumber?: string | null;
  companyItemId?: string | number | null;
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

/** Normalize an item id (number or string) to a comparable trimmed string. */
function idKey(value: string | number | null | undefined): string {
  if (value == null) return "";
  return String(value).trim();
}

/** Coerce an id-ish value to a finite number, or null when not numeric. */
function toNumericId(value: string | number | null | undefined): number | null {
  const key = idKey(value);
  if (!key) return null;
  const n = Number(key);
  return Number.isFinite(n) ? n : null;
}

/**
 * The dedupe unit findMatrixRowForCandidate matches on (normalized candidate
 * name). Exposed so bulk import can group candidates by it — same-key
 * candidates must sync sequentially (they share one matrix row) while distinct
 * names can sync in parallel without ever creating a duplicate row.
 */
export function candidateNameKey(value: string | null | undefined): string {
  return nameKey(value);
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
  linkFields: MatrixLinkFields;
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

  // Strong link columns (created by scripts/ensure-matrix-link-columns.mjs).
  // Numbers vs text are kept apart so the upsert can honor each column's
  // storage type; every write is still optional (skipped if the column is
  // absent), so nothing breaks before the columns exist.
  const numbers: Record<string, number> = {};
  const workforceItemId = toNumericId(profile.workforceItemId ?? profile.id);
  if (workforceItemId != null) numbers.WorkforceItemId = workforceItemId;
  const companyItemId = toNumericId(profile.companyItemId);
  if (companyItemId != null) numbers.CompanyItemId = companyItemId;

  const text: Record<string, string> = { MatrixLinkStatus: "Linked" };
  const workforceNumber = clean(profile.workforceNumber);
  if (workforceNumber) text.WorkforceNumber = workforceNumber;
  const companyNumber = clean(profile.companyNumber);
  if (companyNumber) text.CompanyNumber = companyNumber;
  const candidateName = clean(profile.candidateName);
  if (candidateName) text.CandidateName = candidateName;

  return { source, profileFields, linkFields: { numbers, text } };
}

/**
 * Find the matrix row for a workforce record using the STRONG link keys, in
 * priority order — this is what makes create/update/delete target the right row
 * even when two candidates share a name:
 *   1. `id`     — exact WorkforceItemId match (authoritative once linked).
 *   2. `legacy` — WorkforceNumber + CompanyItemId (a row linked before the id
 *                 was stored, or from a differently-shaped import).
 *   3. `name`   — only an UNAMBIGUOUS, still-unlinked same-name row (exactly one
 *                 row carries the name and it has no WorkforceItemId yet). Never
 *                 adopt when the name is ambiguous or the match already belongs
 *                 to another workforce — the caller then creates a fresh linked
 *                 row (or flags Needs Review) instead of hijacking someone
 *                 else's row.
 * Returns `{ row: null, matchType: "none" }` when nothing safe matches.
 */
export function findMatrixRowByWorkforce<T extends MatrixRowRef>(
  rows: T[],
  profile: Pick<
    WorkforceMatrixProfile,
    | "candidateName"
    | "dateOfBirth"
    | "id"
    | "workforceItemId"
    | "workforceNumber"
    | "companyItemId"
  >,
): { row: T | null; matchType: MatrixMatchType } {
  const wfId = idKey(profile.workforceItemId ?? profile.id);
  if (wfId) {
    const byId = rows.find((row) => idKey(row.workforceItemId) === wfId);
    if (byId) return { row: byId, matchType: "id" };
  }

  const wfNumber = idKey(profile.workforceNumber);
  const companyId = idKey(profile.companyItemId);
  if (wfNumber && companyId) {
    const byLegacy = rows.find(
      (row) =>
        idKey(row.workforceNumber) === wfNumber &&
        idKey(row.companyItemId) === companyId,
    );
    if (byLegacy) return { row: byLegacy, matchType: "legacy" };
  }

  const key = nameKey(profile.candidateName);
  if (key) {
    const nameMatches = rows.filter((row) => nameKey(row.candidateName) === key);
    const unlinked = nameMatches.filter((row) => !idKey(row.workforceItemId));
    // Only adopt a same-name row when there is exactly one row with that name
    // AND it is still unlinked. Multiple same-name rows, or a lone row already
    // linked to a different workforce, are left untouched (ambiguous).
    if (nameMatches.length === 1 && unlinked.length === 1) {
      return { row: unlinked[0]!, matchType: "name" };
    }
  }

  return { row: null, matchType: "none" };
}

/**
 * Classify a matrix row for display (Admin Matrix badge + hide-orphans filter):
 *  - stored WorkforceItemId resolves to a live workforce row → `Linked`
 *  - stored WorkforceItemId no longer resolves (workforce deleted/recreated):
 *    name still matches ≥1 live record → `Needs Review` (re-linkable), else `Orphan`
 *  - no stored id, exactly one live workforce shares the name → `Linked`
 *    (a safe implicit/legacy link — keeps the pre-migration matrix working)
 *  - no stored id, ≥2 live workforce share the name → `Needs Review` (ambiguous)
 *  - no stored id and no name match → `Orphan` (stale row, hidden by default)
 */
export function deriveMatrixLinkStatus(input: {
  hasWorkforceItemId: boolean;
  workforceResolved: boolean;
  nameMatchCount: number;
}): MatrixLinkStatus {
  if (input.hasWorkforceItemId) {
    if (input.workforceResolved) return "Linked";
    return input.nameMatchCount >= 1 ? "Needs Review" : "Orphan";
  }
  if (input.nameMatchCount === 0) return "Orphan";
  if (input.nameMatchCount === 1) return "Linked";
  return "Needs Review";
}
