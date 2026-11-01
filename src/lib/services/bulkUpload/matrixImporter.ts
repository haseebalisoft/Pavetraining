import "server-only";

import {
  listAdminWorkforce,
  toMatrixProfile,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import {
  createBulkLogger,
  type BulkLogger,
} from "@/lib/services/bulkUpload/bulkUploadLog";
import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";
import {
  extractCategoryWritesFromRow,
  loadMatrixCategoryLookupCaches,
  upsertMatrixCategoryRecords,
} from "@/lib/services/bulkUpload/matrixCategoryService";
import {
  listTrainingMatrixExampleRows,
  syncWorkforceToTrainingMatrix,
  upsertUnlinkedMatrixRow,
  type TrainingMatrixExampleRow,
} from "@/lib/services/bulkUpload/trainingMatrixExampleService";
import {
  candidateNameKey,
  findMatrixRowByWorkforce,
  isoDateKey,
} from "@/lib/services/bulkUpload/workforceMatrixSync";
import { normalizeCompanyKey } from "@/lib/services/bulkUpload/matching";
import {
  matchWorkforceForMatrixRow,
  matrixReviewReason,
} from "@/lib/services/bulkUpload/matrixReviewReason";
import {
  normalizeDateValue,
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import { stopIfPastDeadline } from "@/lib/services/bulkUpload/commitDeadline";
import {
  toSharePointFields,
  updateListItemFieldsByKey,
} from "@/lib/services/sharePointListService";
import type { JobDeadline } from "@/lib/jobs/jobOutcome";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkLinkOutcome,
  BulkPreviewRow,
  MatrixReviewReason,
} from "@/types/bulkUpload";

function mapMatrixFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  const fields: Record<string, string | null> = {
    candidateName: pickField(raw, [
      "Candidate Name",
      "CandidateName",
      "Name",
    ]),
    workforceNumber: pickField(raw, [
      "Workforce Number",
      "WorkforceNumber",
      "Workforce No",
    ]),
    company: pickField(raw, ["Company", "Company Name", "CompanyName"]),
    companyNumber: pickField(raw, [
      "Company Number",
      "CompanyNumber",
      "Company No",
    ]),
    department: pickField(raw, ["Department", "Dept", " Department"]),
    dateOfBirth: normalizeDateValue(
      pickField(raw, ["DOB", "Date of birth", "DateOfBirth"]),
    ),
    overallStatus: pickField(raw, ["Overall Status", "OverallStatus", "Status"]),
    needsReview: pickField(raw, ["Needs Review", "NeedsReview"]),
    matrixNotes: pickField(raw, ["Matrix Notes", "MatrixNotes", "Notes"]),
    nextExpiryDate: normalizeDateValue(
      pickField(raw, ["Next Expiry Date", "NextExpiryDate"]),
    ),
    cscsExpiry: normalizeDateValue(
      pickField(raw, ["CSCS Expiry", "Cscs Expiry"]),
    ),
    ssstsExpiry: normalizeDateValue(pickField(raw, ["SSSTS Expiry"])),
    smstsExpiry: normalizeDateValue(pickField(raw, ["SMSTS Expiry"])),
    nrswaExpiry: normalizeDateValue(pickField(raw, ["NRSWA Expiry"])),
    eusrExpiry: normalizeDateValue(
      pickField(raw, ["EUSR Expiry", "Eusr Expiry"]),
    ),
    faceFitExpiry: normalizeDateValue(
      pickField(raw, ["Face ift", "Face Fit", "FaceFit"]),
    ),
  };

  // Map known Training Matrix list expiry columns from client N-code headers.
  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    if (!column.matrixField) continue;
    const value = normalizeDateValue(
      pickField(raw, [column.header, `${column.code} Expiry`, column.code]),
    );
    if (value) fields[column.matrixField] = value;
  }

  // Count filled category cells for preview messaging.
  const categoryWrites = extractCategoryWritesFromRow(raw);
  fields._categoryCount = String(categoryWrites.length);

  return fields;
}

function findWorkforceForMatrix(
  workforce: AdminWorkforceRecord[],
  fields: Record<string, string | null>,
) {
  return matchWorkforceForMatrixRow(workforce, fields);
}

function earliestExpiry(fields: Record<string, string | null>): string | null {
  const dates: number[] = [];
  for (const [key, value] of Object.entries(fields)) {
    if (!value?.trim()) continue;
    if (
      !/expiry$/i.test(key) &&
      !["cscsExpiry", "ssstsExpiry", "smstsExpiry", "nrswaExpiry", "eusrExpiry", "faceFitExpiry", "nextExpiryDate"].includes(
        key,
      )
    ) {
      continue;
    }
    const t = new Date(value).getTime();
    if (!Number.isNaN(t)) dates.push(t);
  }
  if (!dates.length) return null;
  return new Date(Math.min(...dates)).toISOString().slice(0, 10);
}

/**
 * Uploaded matrix cells keyed by the canonical display headers the Training
 * Matrix Update list uses. Blank cells are kept as null here and stripped by
 * `mergeUploadedCells`, so an empty spreadsheet cell never erases a live expiry.
 */
function buildUploadedCells(
  fields: Record<string, string | null>,
  source: Record<string, string | null>,
): Record<string, string | null> {
  return {
    ...source,
    DOB: fields.dateOfBirth ?? source.DOB ?? null,
    "CSCS Expiry": fields.cscsExpiry ?? source["CSCS Expiry"] ?? null,
    "SSSTS Expiry": fields.ssstsExpiry ?? source["SSSTS Expiry"] ?? null,
    "SMSTS Expiry": fields.smstsExpiry ?? source["SMSTS Expiry"] ?? null,
    "NRSWA Expiry": fields.nrswaExpiry ?? source["NRSWA Expiry"] ?? null,
    "Face ift": fields.faceFitExpiry ?? source["Face ift"] ?? null,
  };
}

interface WorkforcePeers {
  namePeers: Map<string, number>;
  nameDobPeers: Map<string, number>;
}

function buildWorkforcePeers(workforce: AdminWorkforceRecord[]): WorkforcePeers {
  const namePeers = new Map<string, number>();
  const nameDobPeers = new Map<string, number>();
  for (const row of workforce) {
    const key = candidateNameKey(row.candidateName);
    if (!key) continue;
    namePeers.set(key, (namePeers.get(key) ?? 0) + 1);
    const dob = isoDateKey(row.dateOfBirth);
    if (!dob) continue;
    const dobKey = `${key}|${dob}`;
    nameDobPeers.set(dobKey, (nameDobPeers.get(dobKey) ?? 0) + 1);
  }
  return { namePeers, nameDobPeers };
}

function peerCounts(
  peers: WorkforcePeers,
  candidate: AdminWorkforceRecord,
): { workforceNamePeers: number; workforceNameDobPeers: number } {
  const key = candidateNameKey(candidate.candidateName);
  const dob = isoDateKey(candidate.dateOfBirth);
  return {
    workforceNamePeers: peers.namePeers.get(key) ?? 1,
    workforceNameDobPeers: dob
      ? (peers.nameDobPeers.get(`${key}|${dob}`) ?? 1)
      : 1,
  };
}

/**
 * The Needs Review row an earlier upload already made for this spreadsheet
 * person: unlinked, same name, and same DOB — or, when the sheet has no DOB,
 * no DOB and the same company (name alone is too weak to merge two people).
 */
function findUnlinkedRowFor(
  rows: TrainingMatrixExampleRow[],
  fields: Record<string, string | null>,
): TrainingMatrixExampleRow | null {
  const name = candidateNameKey(fields.candidateName);
  if (!name) return null;
  const dob = isoDateKey(fields.dateOfBirth);
  const company = normalizeCompanyKey(fields.company);
  const matches = rows.filter((row) => {
    if (String(row.workforceItemId ?? "").trim()) return false;
    if (candidateNameKey(row.candidateName) !== name) return false;
    const rowDob = isoDateKey(row.dateOfBirth);
    if (dob) return rowDob === dob;
    return !rowDob && normalizeCompanyKey(row.storedCompanyName) === company;
  });
  // Several copies already exist (older uploads): update the oldest rather
  // than adding yet another.
  matches.sort((a, b) => Number(a.id) - Number(b.id));
  return matches[0] ?? null;
}

/** Message for a row kept as Needs Review: the exact reason, then what happens next. */
function needsReviewMessage(reason: MatrixReviewReason): string {
  const next =
    reason.code === "candidateNotFound"
      ? "Kept as a Needs Review row — it links automatically once the candidate is added to Workforce with the same Name + DOB."
      : "Kept as a Needs Review row — fix the data and re-upload, or use Link to Workforce in Admin Matrix.";
  return `Needs Review — ${reason.label}: ${reason.detail} ${next}`;
}

/**
 * Validate one matrix spreadsheet row AND predict exactly what the commit will
 * do, so the preview and the commit can never disagree.
 */
function validateMatrixRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  workforce: AdminWorkforceRecord[],
  exampleRows: TrainingMatrixExampleRow[],
  peers: WorkforcePeers,
): BulkPreviewRow {
  const messages: string[] = [];

  // A row with no candidate cannot be attached to anyone. Client matrix exports
  // routinely carry trailing rows like this — either fully blank, or holding a
  // stray expiry (their Face Fit column often runs past the last name). Both are
  // SKIPPED rather than reported as errors, but any value we drop is named
  // explicitly so a spreadsheet mistake is never hidden.
  // (`_categoryCount` is always set by mapMatrixFields and never counts as data.)
  if (!fields.candidateName?.trim() && !fields.workforceNumber?.trim()) {
    const ignored = Object.entries(fields)
      .filter(([key, value]) => key !== "_categoryCount" && value?.trim())
      .map(([key, value]) => `${key}=${value}`);
    return {
      rowNumber,
      status: "Skipped",
      messages: ignored.length
        ? [
            `Skipped — no Candidate Name or Workforce Number on this row. ${ignored.length} value(s) ignored: ${ignored.slice(0, 4).join(", ")}${ignored.length > 4 ? ", …" : ""}. Add the candidate name in the spreadsheet to import them.`,
          ]
        : ["Empty row — skipped."],
      fields,
      resolvedCompanyName: fields.company,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
      linkOutcome: null,
      matrixRowId: null,
    };
  }

  const categoryCount = Number(fields._categoryCount ?? "0");
  const match = findWorkforceForMatrix(workforce, fields);

  // Task C: keep the uploaded training data as a Needs Review row instead of
  // rejecting it, so a later Workforce import can adopt it.
  if (!match.candidate) {
    const outcome: BulkLinkOutcome = match.ambiguous
      ? "needsReviewMultipleMatches"
      : "needsReviewNoMatch";
    const reviewReason =
      match.reason ??
      matrixReviewReason(
        "candidateNotFound",
        `No Workforce record matches "${fields.candidateName ?? ""}".`,
      );
    messages.push(needsReviewMessage(reviewReason));
    if (categoryCount > 0) {
      messages.push(
        `${categoryCount} expiry value(s) will be saved on the Needs Review row (category backup records are skipped until it is linked).`,
      );
    }
    return {
      rowNumber,
      status: "Warning",
      messages,
      fields: {
        ...fields,
        nextExpiryDate: fields.nextExpiryDate ?? earliestExpiry(fields),
      },
      resolvedCompanyName: fields.company,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
      linkOutcome: outcome,
      reviewReason,
      matrixRowId: null,
    };
  }

  const candidate = match.candidate;
  const resolvedCompanyName = candidate.companyName;
  const enrichedFields = {
    ...fields,
    candidateName: candidate.candidateName,
    company: resolvedCompanyName,
    companyNumber: fields.companyNumber ?? candidate.companyNumber,
    workforceNumber: candidate.workforceNumber,
    department: fields.department ?? candidate.department,
    dateOfBirth: fields.dateOfBirth ?? candidate.dateOfBirth,
    nextExpiryDate: fields.nextExpiryDate ?? earliestExpiry(fields),
  };

  // Notes describe normal, healthy work. Kept out of `messages` (which drives
  // Warning status) so a clean 50-row upload does not report 50 warnings.
  const notes: string[] = [];
  if (categoryCount > 0) {
    notes.push(
      `${categoryCount} expiry value(s) will be written to Training Matrix Update (+ category records backup).`,
    );
  }

  // Same resolution the commit will run, against the same snapshot.
  const rowMatch = findMatrixRowByWorkforce(
    exampleRows,
    toMatrixProfile(candidate),
    peerCounts(peers, candidate),
  );

  if (rowMatch.ambiguous) {
    const rowIds = rowMatch.candidates.map((row) => `#${row.id}`).join(", ");
    const reviewReason = matrixReviewReason(
      "multipleMatrixRows",
      `${rowMatch.candidates.length || "Several"} unlinked Training Matrix rows${rowIds ? ` (${rowIds})` : ""} match ${candidate.candidateName} — skipped to avoid attaching training to the wrong person. Delete the extra rows or use Link to Workforce in Admin Matrix, then re-upload.`,
    );
    return {
      rowNumber,
      status: "Warning",
      reviewReason,
      messages: [
        `Needs Review — ${reviewReason.label}: ${reviewReason.detail}`,
        ...messages,
        ...notes,
      ],
      fields: enrichedFields,
      resolvedCompanyName,
      matchedEntityId: candidate.id,
      matchedEntityName: candidate.candidateName,
      duplicateMatch: null,
      linkOutcome: "skippedAmbiguous",
      matrixRowId: null,
    };
  }

  const existingRow = rowMatch.row;
  const adoptsUnlinked =
    existingRow != null && !String(existingRow.workforceItemId ?? "").trim();

  let linkOutcome: BulkLinkOutcome;
  if (!existingRow) {
    linkOutcome = "createdLinked";
  } else if (adoptsUnlinked) {
    linkOutcome = "linkedExistingNeedsReview";
  } else if (match.matchType === "nameDob") {
    linkOutcome = "linkedNameDob";
  } else {
    linkOutcome = "linkedCompanyNameDob";
  }

  if (existingRow) {
    notes.push(
      adoptsUnlinked
        ? `Existing Needs Review row #${existingRow.id} will be linked to this candidate.`
        : `Existing Training Matrix row #${existingRow.id} will be updated.`,
    );
  }

  return {
    rowNumber,
    status: messages.length ? "Warning" : "Ready",
    messages: [...messages, ...notes],
    fields: enrichedFields,
    resolvedCompanyName,
    matchedEntityId: candidate.id,
    matchedEntityName: candidate.candidateName,
    duplicateMatch: null,
    linkOutcome,
    matrixRowId: existingRow ? `example:${existingRow.id}` : null,
  };
}

async function syncWorkforceMetaExpiries(
  candidate: AdminWorkforceRecord,
  fields: Record<string, string | null>,
): Promise<string[]> {
  // Matrix list shows CSCS/EUSR/NRSWA via Workforce lookup columns — keep them
  // in sync. Only the changed expiry fields are PATCHed: this used to call the
  // full updateAdminWorkforce per row, which re-read Permissions, Departments,
  // Companies and the whole Training Matrix list and re-wrote the matrix row —
  // the bulk of a matrix upload's time.
  const patch: Record<string, string> = {};
  const set = (key: "cscsExpiry" | "eusrExpiry" | "swqrExpiry", value?: string | null) => {
    const next = isoDateKey(value);
    if (next && next !== isoDateKey(candidate[key])) patch[key] = next;
  };
  set("cscsExpiry", fields.cscsExpiry);
  set("eusrExpiry", fields.eusrExpiry);
  set("swqrExpiry", fields.nrswaExpiry);
  if (!Object.keys(patch).length) return [];
  await updateListItemFieldsByKey(
    "workforce",
    candidate.id,
    toSharePointFields("workforce", patch),
  );
  return [
    `Workforce meta expiries updated (${Object.keys(patch).join(", ")}).`,
  ];
}

async function writeCategoriesForRow(
  fields: Record<string, string | null>,
  raw: Record<string, string | null>,
  candidate: AdminWorkforceRecord,
  caches?: Awaited<ReturnType<typeof loadMatrixCategoryLookupCaches>>,
): Promise<string[]> {
  const categories = extractCategoryWritesFromRow({ ...raw, ...fields });
  if (!categories.length) return [];
  const result = await upsertMatrixCategoryRecords({
    candidateName: fields.candidateName ?? "",
    companyName: fields.company ?? "",
    // The person this row matched by Name + DOB — never a name lookup.
    candidateId: candidate.id,
    companyId: candidate.companyId,
    categories,
    caches,
  });
  const messages = [
    `Category records: ${result.written} written, ${result.unchanged} already up to date` +
      (result.failed ? `, ${result.failed} failed` : ""),
  ];
  if (result.errors.length) {
    messages.push(...result.errors.slice(0, 5));
  }
  return messages;
}

export async function previewMatrixImport(
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [workforce, exampleRows] = await Promise.all([
    listAdminWorkforce(),
    listTrainingMatrixExampleRows(),
  ]);
  const peers = buildWorkforcePeers(workforce);
  // Preview simulates the commit against a growing snapshot, so two spreadsheet
  // rows for one candidate report "update" on the second, not a second create.
  const liveRows = [...exampleRows];

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapMatrixFields(raw);
    const validated = validateMatrixRow(
      index + 2,
      fields,
      workforce,
      liveRows,
      peers,
    );
    if (validated.matrixRowId) return { ...validated, source: raw };

    // Predict the row this import would create so later rows see it.
    const name = validated.fields.candidateName?.trim();
    if (name && validated.linkOutcome !== "skippedAmbiguous") {
      const candidate = findWorkforceForMatrix(workforce, fields).candidate;
      liveRows.push({
        id: `preview:${index}`,
        candidateName: name,
        dateOfBirth: validated.fields.dateOfBirth ?? null,
        columnValues: {},
        nextExpiryDate: validated.fields.nextExpiryDate ?? null,
        manualOverrides: [],
        workforceItemId: candidate?.id ?? null,
        workforceNumber: candidate?.workforceNumber ?? null,
        companyItemId: candidate?.companyId ?? null,
        companyNumber:
          validated.fields.companyNumber ?? candidate?.companyNumber ?? null,
        storedCompanyName: candidate?.companyName ?? null,
        matrixLinkStatus: candidate ? "Linked" : "Needs Review",
      });
    }
    return { ...validated, source: raw };
  });
}

export async function commitMatrixImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
  log?: BulkLogger;
  /** Stop between rows once passed; unreached rows return as Not processed. */
  deadline?: JobDeadline;
}): Promise<BulkPreviewRow[]> {
  // Each row writes a matrix row + its category records + the workforce meta
  // expiries, serially — a 50-row sheet is minutes of Graph traffic. Without
  // these events the admin sees a spinner with no output and assumes it hung.
  const log = input.log ?? createBulkLogger("commit:trainingMatrix");
  const loadPhase = log.phase("load");
  const timed = async <T>(label: string, fn: () => Promise<T>): Promise<T> => {
    const start = Date.now();
    const result = await fn();
    log.info(`load:${label}`, { ms: Date.now() - start });
    return result;
  };
  const [workforce, exampleRows, categoryCaches] = await Promise.all([
    timed("workforce", listAdminWorkforce),
    timed("exampleRows", listTrainingMatrixExampleRows),
    timed("categoryCaches", loadMatrixCategoryLookupCaches),
  ]);
  loadPhase.end({ workforce: workforce.length, matrixRows: exampleRows.length });
  const peers = buildWorkforcePeers(workforce);
  const liveRows = [...exampleRows];
  const replaceLive = (row: TrainingMatrixExampleRow) => {
    const at = liveRows.findIndex((existing) => existing.id === row.id);
    if (at >= 0) liveRows[at] = row;
    else liveRows.push(row);
  };
  const results: BulkPreviewRow[] = [];
  const categoryBacklog: Array<{
    resultIndex: number;
    fields: Record<string, string | null>;
    source: Record<string, string | null>;
    candidate: AdminWorkforceRecord;
  }> = [];

  const rowsPhase = log.phase("rows");
  let done = 0;
  const total = input.rows.length;
  const tick = () => {
    done += 1;
    // Every row: rows are slow serial writes, so the admin's gauge should
    // move with each one rather than sit on "…" for the first ten.
    log.info("progress", { done, total });
  };

  for (let index = 0; index < input.rows.length; index += 1) {
    const notProcessed = stopIfPastDeadline(
      input.deadline,
      input.rows,
      index,
      log,
    );
    if (notProcessed) {
      results.push(...notProcessed);
      break;
    }
    const row = input.rows[index];
    const source = row.source ?? row.fields;
    const fields = mapMatrixFields(source);
    for (const [key, value] of Object.entries(row.fields)) {
      if (value && !fields[key]) fields[key] = value;
      if (
        value &&
        [
          "candidateName",
          "company",
          "companyNumber",
          "workforceNumber",
          "dateOfBirth",
        ].includes(key)
      ) {
        fields[key] = value;
      }
    }
    fields.dateOfBirth = normalizeDateValue(fields.dateOfBirth);
    for (const key of Object.keys(fields)) {
      if (/expiry$/i.test(key) || key === "nextExpiryDate") {
        fields[key] = normalizeDateValue(fields[key]);
      }
    }
    // Keep category count accurate for messaging after remap.
    fields._categoryCount = String(extractCategoryWritesFromRow(source).length);

    const validated = validateMatrixRow(
      row.rowNumber,
      fields,
      workforce,
      liveRows,
      peers,
    );

    if (
      validated.status === "Error" ||
      validated.status === "Skipped" ||
      validated.linkOutcome === "skippedAmbiguous"
    ) {
      results.push(validated);
      tick();
      continue;
    }

    const uploadedCells = buildUploadedCells(validated.fields, source);

    try {
      const match = findWorkforceForMatrix(workforce, validated.fields);

      // Task C: no safe Workforce owner — preserve the training data on a
      // Needs Review row that a later Workforce import can adopt.
      if (!match.candidate) {
        const unlinked = await upsertUnlinkedMatrixRow({
          candidateName: validated.fields.candidateName ?? "",
          dateOfBirth: validated.fields.dateOfBirth,
          companyName: validated.fields.company,
          companyNumber: validated.fields.companyNumber,
          uploadedCells,
          // Re-uploading the same sheet (e.g. a retry after a dropped
          // connection) must update this person's Needs Review row, not add
          // another copy every time.
          existingRow: findUnlinkedRowFor(liveRows, validated.fields),
        });
        replaceLive(unlinked.row);
        results.push({
          ...validated,
          status: "Warning",
          matrixRowId: `example:${unlinked.id}`,
          messages: [
            ...validated.messages,
            unlinked.created
              ? `Needs Review row created (#${unlinked.id}).`
              : `Needs Review row updated (#${unlinked.id}).`,
          ],
        });
        tick();
        continue;
      }

      const candidate = match.candidate;
      // Task D/E: the single sync path finds-or-creates exactly one linked row.
      const sync = await syncWorkforceToTrainingMatrix(
        toMatrixProfile(candidate),
        {
          existingRows: liveRows,
          uploadedCells,
          ...peerCounts(peers, candidate),
        },
      );

      if (sync.skipped) {
        results.push({
          ...validated,
          status: "Warning",
          linkOutcome: "skippedAmbiguous",
          reviewReason: matrixReviewReason(
            "multipleMatrixRows",
            sync.warnings.join(" ") ||
              "Several unlinked Training Matrix rows match this candidate — skipped to avoid attaching training to the wrong person.",
          ),
          messages: [...validated.messages, ...sync.warnings],
        });
        tick();
        continue;
      }

      if (sync.row) replaceLive(sync.row);
      const duplicateNote =
        input.duplicateMode === "create" && !sync.created
          ? [
              "Duplicate mode 'create' does not apply to Training Matrix rows — the existing linked row was updated instead.",
            ]
          : [];
      // Category backup records are written in a second pass (below), after
      // every row's Matrix dates — what customers actually see — are saved.
      categoryBacklog.push({
        resultIndex: results.length,
        fields: validated.fields,
        source,
        candidate,
      });
      const wfMessages = await syncWorkforceMetaExpiries(
        candidate,
        validated.fields,
      );

      results.push({
        ...validated,
        status: "Imported",
        matrixRowId: sync.matrixId,
        linkOutcome: sync.created ? "createdLinked" : validated.linkOutcome,
        messages: [
          ...validated.messages,
          ...sync.warnings,
          sync.created
            ? `Training Matrix Update: created linked row #${sync.id}.`
            : `Training Matrix Update: updated linked row #${sync.id}.`,
          ...duplicateNote,
          ...wfMessages,
        ],
      });
    } catch (error) {
      results.push({
        ...validated,
        status: "Error",
        messages: [
          ...validated.messages,
          error instanceof Error
            ? error.message
            : "Failed to import matrix row.",
        ],
      });
    }
    tick();
  }
  rowsPhase.end({ rows: total });

  // Pass 2: category backup records (one per filled N-code, ~150 per row on
  // the client's sheet). Only the legacy matrix editor reads them, so they
  // never hold up the customer-visible Matrix rows above. Whatever the time
  // budget doesn't cover is reported and finishes on re-upload (unchanged
  // records are skipped, so a re-run only writes what is missing).
  const categoriesPhase = log.phase("categories");
  let pendingBackups = 0;
  let backupsDone = 0;
  for (const pending of categoryBacklog) {
    backupsDone += 1;
    log.info("progress", { done: backupsDone, total: categoryBacklog.length });
    const row = results[pending.resultIndex]!;
    if (input.deadline?.expired()) {
      pendingBackups += 1;
      row.backupPending = true;
      row.messages = [
        ...row.messages,
        "Category backup records not written yet (time limit) — the Matrix dates are saved; re-upload the file to finish the backup.",
      ];
      continue;
    }
    try {
      row.messages = [
        ...row.messages,
        ...(await writeCategoriesForRow(
          pending.fields,
          pending.source,
          pending.candidate,
          categoryCaches,
        )),
      ];
    } catch (error) {
      row.messages = [
        ...row.messages,
        `Category backup records failed: ${
          error instanceof Error ? error.message : "unknown error"
        } — the Matrix dates are saved.`,
      ];
    }
  }
  categoriesPhase.end({
    rows: categoryBacklog.length,
    pending: pendingBackups,
  });

  return results;
}
