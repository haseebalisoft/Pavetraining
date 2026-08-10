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

/**
 * How a matrix row was matched to a workforce record, strongest first.
 * `name` is the legacy no-DOB-anywhere case; a name match is NEVER used to
 * adopt a row when either side carries a DOB (see findMatrixRowByWorkforce).
 */
export type MatrixMatchType =
  | "id"
  | "legacy"
  | "companyNameDob"
  | "nameDob"
  | "name"
  | "none";

/** Minimal existing-row reference used for matching. */
export interface MatrixRowRef {
  id: string;
  candidateName: string;
  dateOfBirth?: string | null;
  /** Link fields (present once a row has been linked to a workforce record). */
  workforceItemId?: string | number | null;
  workforceNumber?: string | null;
  companyItemId?: string | number | null;
  companyNumber?: string | null;
}

/** Outcome of matching one workforce record against the matrix rows. */
export interface MatrixMatchResult<T> {
  row: T | null;
  /** Always "none" when `row` is null. */
  matchType: MatrixMatchType;
  /**
   * True when more than one row tied at the strongest step reached. The caller
   * must NOT auto-link (and must not create a row either) — the candidate needs
   * an admin decision, so the rows stay Needs Review.
   */
  ambiguous: boolean;
  /** The competing rows when `ambiguous`, for warning text. */
  candidates: T[];
  /** Which step produced the ambiguity. Diagnostics only. */
  ambiguousAt?: MatrixMatchType;
}

export interface MatrixMatchOptions {
  /**
   * Live Workforce records sharing this candidate's NAME (including itself).
   * `> 1` disables the name-only step: name alone must never auto-link when
   * another person shares the name.
   */
  workforceNamePeers?: number;
  /**
   * Live Workforce records sharing NAME + DOB (including itself). `> 1` makes a
   * company-less Name+DOB hit ambiguous: two people with the same name AND date
   * of birth need Company to link safely.
   */
  workforceNameDobPeers?: number;
}

/** Emitted whenever a workforce record ties against several unlinked rows. */
export const AMBIGUOUS_MATRIX_MATCH_WARNING =
  "Multiple unlinked Matrix rows match this Workforce record.";

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

const BLANK_CELL_SENTINELS = /^(—|–|-|n\/?a|null|none)$/i;

/**
 * Any date-ish value → `YYYY-MM-DD`, or "" when it cannot be parsed confidently.
 *
 * DOB reaches this from two very different places: a matrix row's DOB is already
 * normalized ISO (excelSerialToIsoDate), but a Workforce record's dateOfBirth is
 * the RAW SharePoint string, which can be `DD/MM/YYYY`. A naive slice(0,10)
 * would compare "13/07/1981" against "1981-07-13", silently never match, and
 * produce a duplicate matrix row per candidate with no error anywhere — so the
 * UK format is parsed explicitly. Anything ambiguous (2-digit years, junk)
 * returns "" = "unknown", which never counts as a match.
 */
export function isoDateKey(value: string | number | null | undefined): string {
  if (value == null) return "";
  const text = String(value).trim();
  if (!text || BLANK_CELL_SENTINELS.test(text)) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const uk = text.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/);
  if (uk) {
    return `${uk[3]}-${uk[2]!.padStart(2, "0")}-${uk[1]!.padStart(2, "0")}`;
  }
  return "";
}

type DobRelation = "equal" | "conflict" | "unknown";

/**
 * Compare two DOBs. "unknown" when either side is missing/unparseable — it is
 * never treated as a match, but it also does not block a company-keyed match
 * (a company id + name is already strong, and rows legitimately have blank DOB).
 */
function compareDob(
  a: string | number | null | undefined,
  b: string | number | null | undefined,
): DobRelation {
  const x = isoDateKey(a);
  const y = isoDateKey(b);
  if (!x || !y) return "unknown";
  return x === y ? "equal" : "conflict";
}

/**
 * Drop blank / sentinel cells. `upsertTrainingMatrixExampleRow` writes any
 * own-property of `source`, so a present-but-null key is a WIPE, not a no-op.
 * Spreadsheet parsing sets EVERY header (blanks included) as an own-property,
 * so uploaded cells must pass through here or a blank column erases live data.
 */
export function stripBlankCells(
  cells: Record<string, string | null> | undefined,
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(cells ?? {})) {
    if (value == null) continue;
    const text = String(value).trim();
    if (!text || BLANK_CELL_SENTINELS.test(text)) continue;
    out[key] = text;
  }
  return out;
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
 *   1. `id`             — exact WorkforceItemId match (authoritative once linked).
 *   2. `legacy`         — WorkforceNumber + CompanyItemId (a row linked before the
 *                         id was stored, or from a differently-shaped import).
 *   3. `companyNameDob` — an UNLINKED row matching Company (item id or number)
 *                         + Name + DOB. The safest re-link key when the client
 *                         uploads a matrix that carries company info.
 *   4. `nameDob`        — an UNLINKED row matching Name + DOB, when that pair
 *                         hits exactly one row. This is the key that links the
 *                         client's matrix template, which has only Name + DOB.
 *   5. `name`           — legacy fallback for the no-DOB-anywhere case only:
 *                         exactly one row carries the name, it is unlinked, and
 *                         NEITHER side has a DOB to compare. Name alone never
 *                         adopts a row when a DOB exists on either side.
 *
 * Steps 3-5 only ever consider rows with no WorkforceItemId, so a row already
 * linked to a different workforce record is never hijacked.
 *
 * When several unlinked rows tie at the strongest step reached, the result is
 * `ambiguous` with `matchType: "none"` — the caller must neither link nor create
 * (the rows stay Needs Review for an admin to resolve).
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
    | "companyNumber"
  >,
  options: MatrixMatchOptions = {},
): MatrixMatchResult<T> {
  const NONE: MatrixMatchResult<T> = {
    row: null,
    matchType: "none",
    ambiguous: false,
    candidates: [],
  };
  const found = (
    row: T,
    matchType: MatrixMatchType,
    candidates: T[] = [row],
  ): MatrixMatchResult<T> => ({ row, matchType, ambiguous: false, candidates });
  const tie = (
    candidates: T[],
    at: MatrixMatchType,
  ): MatrixMatchResult<T> => ({
    row: null,
    matchType: "none",
    ambiguous: true,
    candidates,
    ambiguousAt: at,
  });

  const wfId = idKey(profile.workforceItemId ?? profile.id);
  const namePeers = options.workforceNamePeers ?? 1;
  const nameDobPeers = options.workforceNameDobPeers ?? 1;

  /**
   * A row may be reused only when nobody else owns it. Rows already stamped
   * with a DIFFERENT WorkforceItemId are structurally invisible to every
   * name/DOB step, which is what makes hijacking impossible.
   */
  const claimable = (row: T): boolean => {
    const owner = idKey(row.workforceItemId);
    return !owner || (wfId !== "" && owner === wfId);
  };

  // 1. WorkforceItemId — authoritative; ignores name so renames still resolve.
  if (wfId) {
    const byId = rows.filter((row) => idKey(row.workforceItemId) === wfId);
    if (byId.length === 1) return found(byId[0]!, "id");
    if (byId.length > 1) {
      // Duplicate rows carrying the same id already exist in live data (created
      // by the pre-fix name-only path). Never report ambiguity here — that would
      // freeze the primary sync path on dirty data. Pick deterministically so
      // the same row is updated on every run.
      const dobHit = byId.find(
        (row) => compareDob(profile.dateOfBirth, row.dateOfBirth) === "equal",
      );
      const pick =
        dobHit ??
        [...byId].sort(
          (a, b) => (toNumericId(a.id) ?? 0) - (toNumericId(b.id) ?? 0),
        )[0]!;
      return found(pick, "id", byId);
    }
  }

  // 2. Legacy WorkforceNumber + CompanyItemId.
  const wfNumber = idKey(profile.workforceNumber);
  const companyId = idKey(profile.companyItemId);
  if (wfNumber && companyId) {
    const byLegacy = rows.filter(
      (row) =>
        claimable(row) &&
        idKey(row.workforceNumber) === wfNumber &&
        idKey(row.companyItemId) === companyId,
    );
    if (byLegacy.length === 1) return found(byLegacy[0]!, "legacy");
    if (byLegacy.length > 1) return tie(byLegacy, "legacy");
  }

  const key = nameKey(profile.candidateName);
  if (!key) return NONE;

  const allNameRows = rows.filter((row) => nameKey(row.candidateName) === key);
  const nameRows = allNameRows.filter(claimable);
  // Every same-name row belongs to someone else — caller creates its own row.
  if (nameRows.length === 0) return NONE;

  const dobRelation = (row: T) => compareDob(profile.dateOfBirth, row.dateOfBirth);

  // 3. Company (item id or number) + Name, with a DOB that does not contradict.
  //    CompanyNumber is the only company key an unlinked matrix-upload row can
  //    carry, so this is the step that re-links matrix-first uploads.
  const companyNumber = nameKey(profile.companyNumber);
  if (companyId || companyNumber) {
    const byCompany = nameRows.filter((row) => {
      if (dobRelation(row) === "conflict") return false;
      const rowCompanyId = idKey(row.companyItemId);
      if (companyId && rowCompanyId && rowCompanyId === companyId) return true;
      const rowCompanyNumber = nameKey(row.companyNumber);
      return Boolean(
        companyNumber && rowCompanyNumber && rowCompanyNumber === companyNumber,
      );
    });
    if (byCompany.length === 1) return found(byCompany[0]!, "companyNameDob");
    if (byCompany.length > 1) return tie(byCompany, "companyNameDob");
  }

  // 4. Name + DOB — the client matrix template's only available key.
  const byNameDob = nameRows.filter((row) => dobRelation(row) === "equal");
  if (byNameDob.length === 1) {
    // One matrix row, but two Workforce people share this name AND DOB: company
    // is required to link safely, so leave it for an admin.
    if (nameDobPeers > 1) return tie(byNameDob, "nameDob");
    return found(byNameDob[0]!, "nameDob");
  }
  // Same name AND same DOB across several unlinked rows cannot be resolved
  // without company info — leave them Needs Review.
  if (byNameDob.length > 1) return tie(byNameDob, "nameDob");

  // 5. Name alone, ONLY when there is no DOB to contradict: the name is unique
  //    across the whole matrix, the row is claimable and carries no DOB, and no
  //    other Workforce record shares the name.
  if (
    namePeers <= 1 &&
    allNameRows.length === 1 &&
    nameRows.length === 1 &&
    !isoDateKey(nameRows[0]!.dateOfBirth)
  ) {
    return found(nameRows[0]!, "name");
  }

  if (nameRows.length > 1) return tie(nameRows, "name");
  return NONE;
}

/**
 * Merge uploaded matrix cells over the workforce-derived source.
 *
 * Blanks are stripped FIRST, so an empty spreadsheet cell can never erase a
 * live expiry (an absent key means "leave that column untouched").
 * For expiry columns the uploaded value wins — the spreadsheet is the newer
 * client truth — but identity stays with the matched Workforce record:
 *  - `Name` is never taken from the sheet (it is the row Title).
 *  - `DOB` only fills when the workforce record has none, because the match was
 *    made ON the DOB; letting a sheet typo rewrite it would break the next sync.
 */
export function mergeUploadedCells(
  built: Record<string, string | null>,
  uploaded?: Record<string, string | null>,
): Record<string, string | null> {
  const cells = stripBlankCells(uploaded);
  const merged: Record<string, string | null> = { ...built };
  for (const [header, value] of Object.entries(cells)) {
    if (header === "Name") continue;
    if (header === "DOB" && clean(built.DOB)) continue;
    merged[header] = value;
  }
  return merged;
}

/**
 * Build the upsert payload for a matrix row that has NO Workforce match yet
 * (matrix uploaded before the candidate exists).
 *
 * Deliberately a separate function rather than a flag on
 * `buildWorkforceMatrixSource`: "never write WorkforceItemId/CompanyItemId" has
 * to be a structural guarantee, not caller discipline. `linkFields.numbers` is
 * always empty, so nothing claims ownership of the row, and a later Workforce
 * import can adopt it via the Name+DOB (or CompanyNumber+Name+DOB) step.
 */
export function buildUnlinkedMatrixSource(input: {
  candidateName: string;
  dateOfBirth?: string | null;
  companyName?: string | null;
  companyNumber?: string | null;
  /** Uploaded cells keyed by canonical matrix display header. */
  uploadedCells?: Record<string, string | null>;
  /** Defaults to "Needs Review". */
  status?: MatrixLinkStatus;
}): {
  source: Record<string, string | null>;
  profileFields: Record<string, string | null>;
  linkFields: MatrixLinkFields;
} {
  const name = input.candidateName.trim();
  if (!name) {
    throw new Error("Candidate name is required to build a matrix row.");
  }

  const cells = stripBlankCells(input.uploadedCells);
  const source: Record<string, string | null> = { Name: name };
  const dob = isoDateKey(input.dateOfBirth) || isoDateKey(cells.DOB);
  if (dob) source.DOB = dob;
  for (const [header, value] of Object.entries(cells)) {
    if (header === "Name" || header === "DOB") continue;
    source[header] = value;
  }

  const profileFields: Record<string, string | null> = {};
  const company = clean(input.companyName);
  if (company) {
    profileFields.Company = company;
    profileFields["Company Name"] = company;
  }

  const text: Record<string, string> = {
    MatrixLinkStatus: input.status ?? "Needs Review",
    CandidateName: name,
  };
  // CompanyNumber is written even though CompanyItemId is not: it is the only
  // company key an unlinked row can carry, and it lets a later Workforce sync
  // adopt this row through the Company + Name + DOB step.
  const companyNumber = clean(input.companyNumber);
  if (companyNumber) text.CompanyNumber = companyNumber;

  return { source, profileFields, linkFields: { numbers: {}, text } };
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
