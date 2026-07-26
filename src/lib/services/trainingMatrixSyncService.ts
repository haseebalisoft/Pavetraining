import "server-only";

import {
  createAdminMatrix,
  listAdminCompanies,
  listAdminWorkforce,
  updateAdminMatrix,
  type AdminMatrixRecord,
  type AdminRegisterKey,
  type AdminTrainingRecord,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import { computeMatrixStatusFromDates } from "@/lib/services/expiryStatusService";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  getListItemByKey,
  getListItemsByKey,
  listHasColumn,
  updateListItemFieldsByKey,
} from "@/lib/services/sharePointListService";
import {
  listAllNormalizedRegisters,
  listNormalizedRegisters,
  normalizeRegisterFromAdminRecord,
  type NormalizedRegisterRecord,
  type RegisterSource,
} from "@/lib/services/trainingRegisterService";
import type {
  MatrixSyncResult,
  MatrixSyncResultItem,
} from "@/types/matrixSync";
import type { Company } from "@/types/models";

export type { MatrixSyncResult, MatrixSyncResultItem };

const NPORS_COLUMN_BY_CODE: Record<
  string,
  keyof Pick<
    AdminMatrixRecord,
    | "n001Expiry"
    | "n003Expiry"
    | "n004Expiry"
    | "n010Expiry"
    | "n020Expiry"
    | "n021Expiry"
    | "n027Expiry"
    | "n100Expiry"
  >
> = {
  N001: "n001Expiry",
  N003: "n003Expiry",
  N004: "n004Expiry",
  N010: "n010Expiry",
  N020: "n020Expiry",
  N021: "n021Expiry",
  N027: "n027Expiry",
  N100: "n100Expiry",
};

export interface MatrixSyncOptions {
  dryRun?: boolean;
  userEmail?: string | null;
  /** When set, load path may seed registers; sync still uses full candidate set. */
  focusRecords?: NormalizedRegisterRecord[];
}

function nameKey(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizeCompanyKey(value: string | null | undefined): string {
  return nameKey(value)
    .replace(/\bltd\b\.?/g, "")
    .replace(/\blimited\b/g, "")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDateMs(value: string | null | undefined): number | null {
  if (!value?.trim()) return null;
  const ms = new Date(value).getTime();
  return Number.isNaN(ms) ? null : ms;
}

/**
 * Apply register expiry onto an existing matrix/workforce date only when:
 * - outcome is Pass
 * - incoming date exists
 * - existing is empty OR incoming is clearly newer (later or equal)
 * Fail never extends. Missing outcome does not extend.
 */
export function shouldApplyPassExpiry(
  existing: string | null | undefined,
  incoming: string | null | undefined,
  outcome: NormalizedRegisterRecord["trainingOutcome"],
): boolean {
  if (outcome !== "Pass") return false;
  if (!incoming?.trim()) return false;
  const incomingMs = parseDateMs(incoming);
  if (incomingMs === null) return false;
  const existingMs = parseDateMs(existing);
  if (existingMs === null) return true;
  return incomingMs >= existingMs;
}

function extractLookupIdFromFields(
  fields: Record<string, unknown>,
  fieldName: string,
): string | null {
  const nested = fields[fieldName];
  if (nested && typeof nested === "object" && "LookupId" in nested) {
    const id = (nested as { LookupId?: unknown }).LookupId;
    if (typeof id === "number" || typeof id === "string") {
      return String(id);
    }
  }
  const lookupId = fields[`${fieldName}LookupId`];
  if (typeof lookupId === "number" || typeof lookupId === "string") {
    const text = String(lookupId).trim();
    return text || null;
  }
  return null;
}

type MatrixRowWithLookups = AdminMatrixRecord & {
  candidateLookupId: string | null;
  companyLookupId: string | null;
  workforceNumber: string | null;
};

async function loadMatrixRowsWithLookups(): Promise<MatrixRowWithLookups[]> {
  const f = getSharePointFields("trainingMatrix");
  const items = await getListItemsByKey("trainingMatrix", { top: 5000 });
  const rows: MatrixRowWithLookups[] = [];

  for (const item of items) {
    const candidateName =
      asLookupOrString(item.fields[f.candidateName]) ??
      asString(item.fields[f.candidateName]);
    if (!candidateName) continue;

    rows.push({
      id: item.id,
      candidateName,
      companyName:
        asLookupOrString(item.fields[f.companyName]) ??
        asLookupOrString(item.fields[f.matrixCompany]),
      department: asNullableString(item.fields[f.department]),
      overallStatus: asNullableString(item.fields[f.overallStatus]),
      needsReview: asBoolean(item.fields[f.needsReview]),
      matrixNotes: asNullableString(item.fields[f.matrixNotes]),
      nextExpiryDate: asNullableString(item.fields[f.nextExpiryDate]),
      n001Expiry: asNullableString(item.fields[f.n001Expiry]),
      n003Expiry: asNullableString(item.fields[f.n003Expiry]),
      n004Expiry: asNullableString(item.fields[f.n004Expiry]),
      n010Expiry: asNullableString(item.fields[f.n010Expiry]),
      n020Expiry: asNullableString(item.fields[f.n020Expiry]),
      n021Expiry: asNullableString(item.fields[f.n021Expiry]),
      n027Expiry: asNullableString(item.fields[f.n027Expiry]),
      n100Expiry: asNullableString(item.fields[f.n100Expiry]),
      candidateLookupId:
        extractLookupIdFromFields(item.fields, f.candidateName) ??
        extractLookupIdFromFields(item.fields, "CandidateName"),
      companyLookupId:
        extractLookupIdFromFields(item.fields, f.matrixCompany) ??
        extractLookupIdFromFields(item.fields, f.companyName),
      workforceNumber:
        asLookupOrString(
          item.fields.Candidate_x0020_Name_x003a__x002,
        ) ??
        asNullableString(item.fields.WorkforceNumber),
    });
  }

  return rows;
}

type ResolvedCandidate = {
  workforce: AdminWorkforceRecord;
  company: Company;
};

type SyncContext = {
  companies: Company[];
  workforce: AdminWorkforceRecord[];
  matrixRows: MatrixRowWithLookups[];
  registers: NormalizedRegisterRecord[];
  dryRun: boolean;
  optionalColumns: { matrixSyncedAt: boolean; automationStatus: boolean };
};

function findCompany(
  companies: Company[],
  hints: {
    companyLookupId?: string | null;
    companyNumber?: string | null;
    companyName?: string | null;
  },
): Company | null {
  if (hints.companyLookupId) {
    const byId = companies.find((c) => c.id === hints.companyLookupId);
    if (byId) return byId;
  }
  if (hints.companyNumber?.trim()) {
    const numberKey = nameKey(hints.companyNumber);
    const byNumber = companies.find(
      (c) => nameKey(c.companyNumber) === numberKey,
    );
    if (byNumber) return byNumber;
  }
  if (hints.companyName?.trim()) {
    const exact = companies.find(
      (c) => nameKey(c.companyName) === nameKey(hints.companyName),
    );
    if (exact) return exact;
    const normalized = normalizeCompanyKey(hints.companyName);
    return (
      companies.find(
        (c) => normalizeCompanyKey(c.companyName) === normalized,
      ) ?? null
    );
  }
  return null;
}

function findWorkforce(
  workforce: AdminWorkforceRecord[],
  hints: {
    candidateLookupId?: string | null;
    workforceNumber?: string | null;
    candidateName?: string | null;
    companyName?: string | null;
  },
): AdminWorkforceRecord | null {
  if (hints.candidateLookupId) {
    const byId = workforce.find((w) => w.id === hints.candidateLookupId);
    if (byId) return byId;
  }
  if (hints.workforceNumber?.trim()) {
    const numberKey = nameKey(hints.workforceNumber);
    const byNumber = workforce.find(
      (w) => nameKey(w.workforceNumber) === numberKey,
    );
    if (byNumber) return byNumber;
  }
  if (hints.candidateName?.trim() && hints.companyName?.trim()) {
    const cName = nameKey(hints.candidateName);
    const coName = nameKey(hints.companyName);
    const coNorm = normalizeCompanyKey(hints.companyName);
    return (
      workforce.find(
        (w) =>
          nameKey(w.candidateName) === cName &&
          (nameKey(w.companyName) === coName ||
            normalizeCompanyKey(w.companyName) === coNorm),
      ) ?? null
    );
  }
  if (hints.candidateName?.trim()) {
    const cName = nameKey(hints.candidateName);
    const matches = workforce.filter((w) => nameKey(w.candidateName) === cName);
    return matches.length === 1 ? matches[0]! : null;
  }
  return null;
}

function findMatrixRow(
  rows: SyncContext["matrixRows"],
  workforce: AdminWorkforceRecord,
  company: Company,
) {
  const byCandidateId = rows.filter(
    (row) => row.candidateLookupId && row.candidateLookupId === workforce.id,
  );
  if (byCandidateId.length === 1) return byCandidateId[0]!;
  if (byCandidateId.length > 1) {
    const forCompany = byCandidateId.filter(
      (row) =>
        row.companyLookupId === company.id ||
        nameKey(row.companyName) === nameKey(company.companyName),
    );
    if (forCompany.length === 1) return forCompany[0]!;
    if (forCompany.length > 1) return forCompany[0]!;
    return byCandidateId[0]!;
  }

  if (workforce.workforceNumber?.trim()) {
    const byNumber = rows.filter(
      (row) =>
        nameKey(row.workforceNumber) === nameKey(workforce.workforceNumber),
    );
    if (byNumber.length === 1) return byNumber[0]!;
    if (byNumber.length > 1) {
      const forCompany = byNumber.find(
        (row) =>
          row.companyLookupId === company.id ||
          nameKey(row.companyName) === nameKey(company.companyName),
      );
      if (forCompany) return forCompany;
    }
  }

  const byNameCompany = rows.filter(
    (row) =>
      nameKey(row.candidateName) === nameKey(workforce.candidateName) &&
      (nameKey(row.companyName) === nameKey(company.companyName) ||
        normalizeCompanyKey(row.companyName) ===
          normalizeCompanyKey(company.companyName)),
  );
  if (byNameCompany.length >= 1) return byNameCompany[0]!;
  return null;
}

function recordsForCandidate(
  registers: NormalizedRegisterRecord[],
  workforce: AdminWorkforceRecord,
  company: Company,
): NormalizedRegisterRecord[] {
  return registers.filter((record) => {
    if (record.candidateLookupId && record.candidateLookupId === workforce.id) {
      return true;
    }
    if (
      nameKey(record.candidateName) === nameKey(workforce.candidateName) &&
      (nameKey(record.companyName) === nameKey(company.companyName) ||
        normalizeCompanyKey(record.companyName) ===
          normalizeCompanyKey(company.companyName) ||
        (record.companyLookupId && record.companyLookupId === company.id))
    ) {
      return true;
    }
    return false;
  });
}

function emptyResultItem(
  partial: Partial<MatrixSyncResultItem> & {
    candidate: string;
    company: string;
  },
): MatrixSyncResultItem {
  return {
    candidate: partial.candidate,
    company: partial.company,
    candidateId: partial.candidateId ?? null,
    companyId: partial.companyId ?? null,
    registerSources: partial.registerSources ?? [],
    matrixRowId: partial.matrixRowId ?? null,
    matrixRowFound: partial.matrixRowFound ?? false,
    matrixRowCreated: partial.matrixRowCreated ?? false,
    fieldsUpdated: partial.fieldsUpdated ?? [],
    warnings: partial.warnings ?? [],
    errors: partial.errors ?? [],
    skipped: partial.skipped ?? false,
    skipReason: partial.skipReason,
  };
}

function buildSummary(items: MatrixSyncResultItem[]) {
  return {
    processed: items.length,
    updated: items.filter((i) => i.fieldsUpdated.length > 0 && !i.skipped)
      .length,
    created: items.filter((i) => i.matrixRowCreated).length,
    skipped: items.filter((i) => i.skipped).length,
    errors: items.filter((i) => i.errors.length > 0).length,
    warnings: items.reduce((sum, i) => sum + i.warnings.length, 0),
  };
}

async function loadSyncContext(
  options: MatrixSyncOptions,
): Promise<SyncContext> {
  const [companies, workforce, matrixRows, registers] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    loadMatrixRowsWithLookups(),
    options.focusRecords?.length
      ? Promise.resolve(options.focusRecords)
      : listAllNormalizedRegisters(),
  ]);

  const [matrixSyncedAt, automationStatus] = await Promise.all([
    listHasColumn("trainingMatrix", "MatrixSyncedAt"),
    listHasColumn("trainingMatrix", "AutomationStatus"),
  ]);

  return {
    companies,
    workforce,
    matrixRows,
    registers,
    dryRun: Boolean(options.dryRun),
    optionalColumns: { matrixSyncedAt, automationStatus },
  };
}

async function writeOptionalSyncMeta(
  matrixRowId: string,
  optionalColumns: SyncContext["optionalColumns"],
  dryRun: boolean,
): Promise<string[]> {
  const updated: string[] = [];
  if (dryRun) {
    if (optionalColumns.matrixSyncedAt) updated.push("MatrixSyncedAt");
    if (optionalColumns.automationStatus) updated.push("AutomationStatus");
    return updated;
  }

  const payload: Record<string, unknown> = {};
  if (optionalColumns.matrixSyncedAt) {
    payload.MatrixSyncedAt = new Date().toISOString();
    updated.push("MatrixSyncedAt");
  }
  if (optionalColumns.automationStatus) {
    payload.AutomationStatus = "Synced";
    updated.push("AutomationStatus");
  }
  if (Object.keys(payload).length > 0) {
    await updateListItemFieldsByKey("trainingMatrix", matrixRowId, payload);
  }
  return updated;
}

async function syncOneCandidate(
  ctx: SyncContext,
  resolved: ResolvedCandidate,
): Promise<MatrixSyncResultItem> {
  const { workforce, company } = resolved;
  const warnings: string[] = [];
  const errors: string[] = [];
  const fieldsUpdated: string[] = [];

  const registersForMapping = recordsForCandidate(
    ctx.registers,
    workforce,
    company,
  );

  const sources = [
    ...new Set(registersForMapping.map((r) => r.source)),
  ] as RegisterSource[];

  let matrixRow = findMatrixRow(ctx.matrixRows, workforce, company);
  let matrixRowCreated = false;

  if (!matrixRow) {
    if (!workforce.candidateName.trim() || !company.companyName.trim()) {
      return emptyResultItem({
        candidate: workforce.candidateName,
        company: company.companyName,
        candidateId: workforce.id,
        companyId: company.id,
        registerSources: sources,
        skipped: true,
        skipReason:
          "Not enough candidate/company information to create matrix row.",
        warnings: ["Matrix row missing and cannot be created safely."],
      });
    }

    if (ctx.dryRun) {
      matrixRowCreated = true;
      fieldsUpdated.push("candidateName", "companyName", "(create)");
    } else {
      try {
        const created = await createAdminMatrix({
          candidateName: workforce.candidateName,
          companyName: company.companyName,
          department: workforce.department,
          needsReview: true,
          overallStatus: "Missing Data",
        });
        try {
          await updateListItemFieldsByKey("trainingMatrix", created.id, {
            CandidateNameLookupId: Number(workforce.id),
            MatrixCompanyLookupId: Number(company.id),
          });
        } catch {
          warnings.push(
            "Matrix row created; lookup IDs could not be set (display names kept).",
          );
        }
        matrixRow = {
          ...created,
          candidateLookupId: workforce.id,
          companyLookupId: company.id,
          workforceNumber: workforce.workforceNumber,
        };
        ctx.matrixRows.push(matrixRow);
        matrixRowCreated = true;
        fieldsUpdated.push("candidateName", "companyName");
      } catch (error) {
        errors.push(
          error instanceof Error
            ? error.message
            : "Failed to create Training Matrix row.",
        );
        return emptyResultItem({
          candidate: workforce.candidateName,
          company: company.companyName,
          candidateId: workforce.id,
          companyId: company.id,
          registerSources: sources,
          errors,
          skipped: true,
          skipReason: "Create matrix row failed.",
        });
      }
    }
  }

  const draft: AdminMatrixRecord = matrixRow
    ? { ...matrixRow }
    : {
        id: "dry-run-new",
        candidateName: workforce.candidateName,
        companyName: company.companyName,
        department: workforce.department,
        overallStatus: null,
        needsReview: true,
        matrixNotes: null,
        nextExpiryDate: null,
        n001Expiry: null,
        n003Expiry: null,
        n004Expiry: null,
        n010Expiry: null,
        n020Expiry: null,
        n021Expiry: null,
        n027Expiry: null,
        n100Expiry: null,
      };

  const patch: Record<string, unknown> = {};
  let workforceEusrExpiry: string | null | undefined;
  let workforceSwqrExpiry: string | null | undefined;
  let needsReviewForced = false;

  for (const record of registersForMapping) {
    if (record.trainingOutcome === null) {
      needsReviewForced = true;
      warnings.push(
        `${record.source} #${record.id}: missing TrainingOutcome → Records to Review.`,
      );
      continue;
    }

    if (record.trainingOutcome === "Fail") {
      warnings.push(
        `${record.source} #${record.id}: Fail — expiry not extended.`,
      );
      continue;
    }

    if (record.source === "NPORS") {
      if (record.nporsCategories.length === 0) {
        warnings.push(
          `NPORS #${record.id}: Pass but no mapped NPORS category (N001–N100).`,
        );
        continue;
      }
      if (!record.expiry) {
        needsReviewForced = true;
        warnings.push(`NPORS #${record.id}: Pass but missing Expiry.`);
        continue;
      }
      for (const code of record.nporsCategories) {
        const field = NPORS_COLUMN_BY_CODE[code];
        if (!field) continue;
        const existing = draft[field];
        if (shouldApplyPassExpiry(existing, record.expiry, "Pass")) {
          draft[field] = record.expiry;
          patch[field] = record.expiry;
          if (!fieldsUpdated.includes(field)) fieldsUpdated.push(field);
        } else if (existing) {
          warnings.push(
            `NPORS ${code}: kept existing matrix date (register not newer).`,
          );
        }
      }
    }

    if (record.source === "EUSR") {
      if (!record.expiry) {
        needsReviewForced = true;
        warnings.push(`EUSR #${record.id}: Pass but missing Expiry.`);
        continue;
      }
      if (
        workforceEusrExpiry === undefined ||
        shouldApplyPassExpiry(workforceEusrExpiry, record.expiry, "Pass")
      ) {
        workforceEusrExpiry = record.expiry;
      }
    }

    if (record.source === "NRSWA") {
      if (!record.expiry) {
        needsReviewForced = true;
        warnings.push(`NRSWA #${record.id}: Pass but missing Expirydate.`);
        continue;
      }
      if (
        workforceSwqrExpiry === undefined ||
        shouldApplyPassExpiry(workforceSwqrExpiry, record.expiry, "Pass")
      ) {
        workforceSwqrExpiry = record.expiry;
      }
    }

    if (record.source === "In-House") {
      if (!record.expiry) {
        needsReviewForced = true;
        warnings.push(`In-House #${record.id}: Pass but missing ExpiryDate.`);
      }
    }
  }

  const inHousePassExpiries = registersForMapping
    .filter(
      (r) =>
        r.source === "In-House" && r.trainingOutcome === "Pass" && r.expiry,
    )
    .map((r) => r.expiry);

  const status = computeMatrixStatusFromDates([
    draft.n001Expiry,
    draft.n003Expiry,
    draft.n004Expiry,
    draft.n010Expiry,
    draft.n020Expiry,
    draft.n021Expiry,
    draft.n027Expiry,
    draft.n100Expiry,
    workforceEusrExpiry,
    workforceSwqrExpiry,
    ...inHousePassExpiries,
  ]);

  if (draft.nextExpiryDate !== status.nextExpiryDate) {
    patch.nextExpiryDate = status.nextExpiryDate;
    fieldsUpdated.push("nextExpiryDate");
  }
  if (draft.overallStatus !== status.overallStatus) {
    patch.overallStatus = status.overallStatus;
    fieldsUpdated.push("overallStatus");
  }

  const needsReview = needsReviewForced || status.needsReview;
  if (draft.needsReview !== needsReview) {
    patch.needsReview = needsReview;
    fieldsUpdated.push("needsReview");
  }

  if (!draft.matrixNotes?.trim() && sources.length > 0) {
    patch.matrixNotes = `Synced from ${sources.join(", ")} registers.`;
    fieldsUpdated.push("matrixNotes");
  }

  if (ctx.dryRun) {
    for (const key of Object.keys(patch)) {
      if (!fieldsUpdated.includes(key)) fieldsUpdated.push(key);
    }
  }

  if (!ctx.dryRun && matrixRow && Object.keys(patch).length > 0) {
    try {
      const updated = await updateAdminMatrix(matrixRow.id, patch);
      Object.assign(matrixRow, updated);
    } catch (error) {
      errors.push(
        error instanceof Error
          ? error.message
          : "Failed to update Training Matrix row.",
      );
    }
  }

  if (!ctx.dryRun && workforceEusrExpiry) {
    try {
      await updateListItemFieldsByKey("workforce", workforce.id, {
        EusrExpiry: workforceEusrExpiry,
      });
      fieldsUpdated.push("Workforce.EusrExpiry");
    } catch (error) {
      warnings.push(
        error instanceof Error
          ? `Could not update Workforce EUSR expiry: ${error.message}`
          : "Could not update Workforce EUSR expiry.",
      );
    }
  } else if (ctx.dryRun && workforceEusrExpiry) {
    fieldsUpdated.push("Workforce.EusrExpiry");
  }

  if (!ctx.dryRun && workforceSwqrExpiry) {
    try {
      await updateListItemFieldsByKey("workforce", workforce.id, {
        SwqrExpiry: workforceSwqrExpiry,
      });
      fieldsUpdated.push("Workforce.SwqrExpiry");
    } catch (error) {
      warnings.push(
        error instanceof Error
          ? `Could not update Workforce SWQR expiry: ${error.message}`
          : "Could not update Workforce SWQR expiry.",
      );
    }
  } else if (ctx.dryRun && workforceSwqrExpiry) {
    fieldsUpdated.push("Workforce.SwqrExpiry");
  }

  if (matrixRow && (fieldsUpdated.length > 0 || matrixRowCreated)) {
    try {
      const meta = await writeOptionalSyncMeta(
        matrixRow.id,
        ctx.optionalColumns,
        ctx.dryRun,
      );
      fieldsUpdated.push(...meta.filter((f) => !fieldsUpdated.includes(f)));
    } catch (error) {
      warnings.push(
        error instanceof Error
          ? `Optional sync meta skipped: ${error.message}`
          : "Optional sync meta skipped.",
      );
    }
  }

  const skipped =
    fieldsUpdated.length === 0 && !matrixRowCreated && errors.length === 0;

  return emptyResultItem({
    candidate: workforce.candidateName,
    company: company.companyName,
    candidateId: workforce.id,
    companyId: company.id,
    registerSources: sources,
    matrixRowId: matrixRow?.id ?? null,
    matrixRowFound: Boolean(matrixRow) && !matrixRowCreated,
    matrixRowCreated,
    fieldsUpdated: [...new Set(fieldsUpdated)],
    warnings,
    errors,
    skipped,
    skipReason: skipped ? "No matrix field changes required." : undefined,
  });
}

function resolveCandidateFromId(
  ctx: SyncContext,
  candidateId: string,
): ResolvedCandidate | null {
  const workforce = ctx.workforce.find((w) => w.id === candidateId);
  if (!workforce) return null;
  const matched =
    findCompany(ctx.companies, { companyName: workforce.companyName }) ??
    ctx.companies.find(
      (c) =>
        normalizeCompanyKey(c.companyName) ===
        normalizeCompanyKey(workforce.companyName),
    );
  if (!matched) return null;
  return { workforce, company: matched };
}

async function auditSync(
  userEmail: string | null | undefined,
  scope: string,
  result: MatrixSyncResult,
  phase: "started" | "completed" | "failed" = "completed",
) {
  const { logMatrixSync } = await import("@/lib/services/auditLogService");
  await logMatrixSync({
    userEmail,
    phase:
      phase === "started"
        ? "started"
        : result.summary.errors > 0
          ? "failed"
          : "completed",
    scope,
    success: result.summary.errors === 0,
    errorMessage:
      result.summary.errors > 0
        ? `${result.summary.errors} sync item(s) had errors`
        : null,
    metadata: {
      processed: result.summary.processed,
      updated: result.summary.updated,
      created: result.summary.created,
      skipped: result.summary.skipped,
      dryRun: result.dryRun,
    },
  });
}

export async function syncCandidateMatrix(
  candidateId: string,
  options: MatrixSyncOptions = {},
): Promise<MatrixSyncResult> {
  const { logMatrixSync } = await import("@/lib/services/auditLogService");
  await logMatrixSync({
    userEmail: options.userEmail,
    phase: "started",
    scope: `candidate/${candidateId}`,
  });

  const ctx = await loadSyncContext(options);
  const resolved = resolveCandidateFromId(ctx, candidateId);
  const items: MatrixSyncResultItem[] = [];

  if (!resolved) {
    items.push(
      emptyResultItem({
        candidate: candidateId,
        company: "—",
        candidateId,
        errors: ["Candidate not found in Workforce List."],
        skipped: true,
        skipReason: "Candidate not found.",
      }),
    );
  } else {
    items.push(await syncOneCandidate(ctx, resolved));
  }

  const result: MatrixSyncResult = {
    dryRun: Boolean(options.dryRun),
    scope: "candidate",
    items,
    summary: buildSummary(items),
  };
  await auditSync(options.userEmail, `candidate/${candidateId}`, result);
  return result;
}

export async function syncCompanyMatrix(
  companyId: string,
  options: MatrixSyncOptions = {},
): Promise<MatrixSyncResult> {
  const { logMatrixSync } = await import("@/lib/services/auditLogService");
  await logMatrixSync({
    userEmail: options.userEmail,
    phase: "started",
    scope: `company/${companyId}`,
  });

  const ctx = await loadSyncContext(options);
  const company = ctx.companies.find((c) => c.id === companyId);
  const items: MatrixSyncResultItem[] = [];

  if (!company) {
    items.push(
      emptyResultItem({
        candidate: "—",
        company: companyId,
        companyId,
        errors: ["Company not found."],
        skipped: true,
        skipReason: "Company not found.",
      }),
    );
  } else {
    const candidates = ctx.workforce.filter(
      (w) =>
        nameKey(w.companyName) === nameKey(company.companyName) ||
        normalizeCompanyKey(w.companyName) ===
          normalizeCompanyKey(company.companyName),
    );
    if (candidates.length === 0) {
      items.push(
        emptyResultItem({
          candidate: "—",
          company: company.companyName,
          companyId: company.id,
          skipped: true,
          skipReason: "No workforce candidates for this company.",
          warnings: ["No candidates to sync."],
        }),
      );
    } else {
      for (const workforce of candidates) {
        items.push(await syncOneCandidate(ctx, { workforce, company }));
      }
    }
  }

  const result: MatrixSyncResult = {
    dryRun: Boolean(options.dryRun),
    scope: "company",
    items,
    summary: buildSummary(items),
  };
  await auditSync(options.userEmail, `company/${companyId}`, result);
  return result;
}

export async function syncAllMatrix(
  options: MatrixSyncOptions = {},
): Promise<MatrixSyncResult> {
  const { logMatrixSync } = await import("@/lib/services/auditLogService");
  await logMatrixSync({
    userEmail: options.userEmail,
    phase: "started",
    scope: "all",
  });

  const ctx = await loadSyncContext(options);
  const items: MatrixSyncResultItem[] = [];

  // Sync candidates that appear in registers or already have matrix rows.
  const candidateIds = new Set<string>();
  for (const row of ctx.matrixRows) {
    if (row.candidateLookupId) candidateIds.add(row.candidateLookupId);
  }
  for (const record of ctx.registers) {
    if (record.candidateLookupId) candidateIds.add(record.candidateLookupId);
  }

  // Also include name-matched workforce for register rows without lookup IDs.
  for (const record of ctx.registers) {
    const company = findCompany(ctx.companies, {
      companyLookupId: record.companyLookupId,
      companyName: record.companyName,
    });
    const workforce = findWorkforce(ctx.workforce, {
      candidateLookupId: record.candidateLookupId,
      candidateName: record.candidateName,
      companyName: company?.companyName ?? record.companyName,
    });
    if (workforce) candidateIds.add(workforce.id);
  }

  for (const candidateId of candidateIds) {
    const resolved = resolveCandidateFromId(ctx, candidateId);
    if (!resolved) {
      // Try resolve from register name if id was a lookup that isn't in workforce list load
      continue;
    }
    items.push(await syncOneCandidate(ctx, resolved));
  }

  // Name-only register rows where workforce id wasn't captured
  for (const record of ctx.registers) {
    const company = findCompany(ctx.companies, {
      companyLookupId: record.companyLookupId,
      companyName: record.companyName,
    });
    if (!company) continue;
    const workforce = findWorkforce(ctx.workforce, {
      candidateLookupId: record.candidateLookupId,
      candidateName: record.candidateName,
      companyName: company.companyName,
    });
    if (!workforce || candidateIds.has(workforce.id)) continue;
    candidateIds.add(workforce.id);
    items.push(await syncOneCandidate(ctx, { workforce, company }));
  }

  const result: MatrixSyncResult = {
    dryRun: Boolean(options.dryRun),
    scope: "all",
    items,
    summary: buildSummary(items),
  };
  await auditSync(options.userEmail, "all", result);
  return result;
}

/**
 * After a register create/update: sync that candidate only (not full matrix).
 */
export async function syncAfterRegisterSave(
  registerKey: AdminRegisterKey,
  record: AdminTrainingRecord,
  options: MatrixSyncOptions = {},
): Promise<MatrixSyncResult> {
  const focus = normalizeRegisterFromAdminRecord(registerKey, record);

  // Enrich lookup IDs from the saved SharePoint item when possible.
  try {
    const item = await getListItemByKey(registerKey, record.id);
    if (item) {
      const fCandidate =
        registerKey === "nporsRegister"
          ? "CandidateName"
          : registerKey === "eusrRegister"
            ? "CandidateName"
            : registerKey === "nrswaRegister"
              ? "CandidateName"
              : "CandidateName";
      focus.candidateLookupId =
        extractLookupIdFromFields(item.fields, fCandidate) ??
        focus.candidateLookupId;
      focus.companyLookupId =
        extractLookupIdFromFields(item.fields, "CompanyName") ??
        focus.companyLookupId;
    }
  } catch {
    // continue with name-based matching
  }

  const ctx = await loadSyncContext({ ...options, focusRecords: [focus] });
  // Ensure we have full register set for status — reload all for this candidate path.
  ctx.registers = await listAllNormalizedRegisters();

  const company = findCompany(ctx.companies, {
    companyLookupId: focus.companyLookupId,
    companyName: focus.companyName,
  });
  const workforce = findWorkforce(ctx.workforce, {
    candidateLookupId: focus.candidateLookupId,
    candidateName: focus.candidateName,
    companyName: company?.companyName ?? focus.companyName,
  });

  const items: MatrixSyncResultItem[] = [];
  if (!workforce || !company) {
    items.push(
      emptyResultItem({
        candidate: focus.candidateName,
        company: focus.companyName,
        registerSources: [focus.source],
        skipped: true,
        skipReason: "Could not resolve candidate/company for matrix sync.",
        warnings: [
          !workforce ? "Candidate not matched in Workforce." : "",
          !company ? "Company not matched in Company List." : "",
        ].filter(Boolean),
      }),
    );
  } else {
    items.push(await syncOneCandidate(ctx, { workforce, company }));
  }

  const result: MatrixSyncResult = {
    dryRun: Boolean(options.dryRun),
    scope: "register",
    items,
    summary: buildSummary(items),
  };
  await auditSync(
    options.userEmail,
    `register/${registerKey}/${record.id}`,
    result,
  );
  return result;
}

/** Load registers for one key — used by tests / diagnostics. */
export { listNormalizedRegisters };
