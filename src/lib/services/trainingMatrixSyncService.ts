import "server-only";

import {
  listAdminCompanies,
  listAdminWorkforce,
  type AdminMatrixRecord,
  type AdminRegisterKey,
  type AdminTrainingRecord,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import { computeMatrixStatusFromDates } from "@/lib/services/expiryStatusService";
import {
  MATRIX_CATEGORY_EXPIRY_COLUMNS,
} from "@/lib/services/bulkUpload/matrixExpiryFieldMap";
import {
  listTrainingMatrixExampleRows,
  stripExampleMatrixId,
  upsertTrainingMatrixExampleRow,
} from "@/lib/services/bulkUpload/trainingMatrixExampleService";
import {
  buildWorkforceMatrixSource,
  findMatrixRowByWorkforce,
} from "@/lib/services/bulkUpload/workforceMatrixSync";
import {
  computeMatrixFieldAfterRemoval,
  latestPassExpiry,
  shouldApplyPassExpiry,
} from "@/lib/services/bulkUpload/registerMatrixFieldSync";
import { addCalendarYearsIso } from "@/lib/utils/formatDate";
import {
  listAllNormalizedRegisters,
  listNormalizedRegisters,
  normalizeRegisterFromAdminRecord,
  registerKeyToSource,
  type NormalizedRegisterRecord,
  type RegisterSource,
} from "@/lib/services/trainingRegisterService";
import {
  ASBESTOS_MATRIX_HEADER,
  isAsbestosAwarenessCategory,
  isManualOverrideHeader,
} from "@/lib/training/matrixManualOverrides";
import {
  EUSR_MATRIX_CATEGORY_COLUMNS,
  eusrMatrixHeaderForCategory,
  parseEusrCategories,
} from "@/lib/training/eusrOptions";
import { matrixStoredExpiryKey } from "@/lib/training/matrixEditorFields";
import type {
  MatrixSyncResult,
  MatrixSyncResultItem,
} from "@/types/matrixSync";
import type { Company } from "@/types/models";

export type { MatrixSyncResult, MatrixSyncResultItem };

/** NPORS category code → Excel / Training Matrix column header. */
const NPORS_HEADER_BY_CODE: Record<string, string> = Object.fromEntries(
  MATRIX_CATEGORY_EXPIRY_COLUMNS.map((column) => [
    column.code.toUpperCase(),
    column.header,
  ]),
);

function eusrHeadersForRecord(record: NormalizedRegisterRecord): string[] {
  const categories =
    record.eusrCategories.length > 0
      ? record.eusrCategories
      : parseEusrCategories(record.eusrCategory);
  const headers: string[] = [];
  const seen = new Set<string>();
  for (const category of categories) {
    const header = eusrMatrixHeaderForCategory(category);
    if (!header || seen.has(header)) continue;
    seen.add(header);
    headers.push(header);
  }
  return headers;
}

function stampEusrCategoryDates(
  header: string,
  expiry: string | null,
  trainingDate: string | null | undefined,
  columnValues: Record<string, string | null>,
  categoryTrainingDates: Record<string, string | null>,
  fieldsUpdated: string[],
) {
  columnValues[header] = expiry;
  categoryTrainingDates[matrixStoredExpiryKey(header)] = expiry;
  if (trainingDate?.trim()) {
    categoryTrainingDates[header] = trainingDate.trim().slice(0, 10);
  } else if (!expiry) {
    categoryTrainingDates[header] = null;
  }
  if (!fieldsUpdated.includes(header)) fieldsUpdated.push(header);
}

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

type MatrixRowWithLookups = AdminMatrixRecord & {
  /** Strong link keys — mirrors MatrixRowRef so this can be passed straight
   * into findMatrixRowByWorkforce with no translation step. */
  workforceItemId: string | null;
  companyItemId: string | null;
  workforceNumber: string | null;
  /** Raw SharePoint item id on Training Matrix Update (no `example:` prefix). */
  exampleItemId: string | null;
};

function exampleRowToMatrix(
  example: Awaited<ReturnType<typeof listTrainingMatrixExampleRows>>[number],
  workforce: AdminWorkforceRecord | null,
  company: Company | null,
): MatrixRowWithLookups {
  return {
    id: `example:${example.id}`,
    candidateName: example.candidateName,
    companyName: workforce?.companyName ?? company?.companyName ?? null,
    department: workforce?.department ?? null,
    dateOfBirth: example.dateOfBirth ?? workforce?.dateOfBirth ?? null,
    overallStatus: null,
    needsReview: !example.nextExpiryDate,
    matrixNotes: null,
    nextExpiryDate: example.nextExpiryDate,
    n001Expiry: example.columnValues["N001 - Ind FLT"] ?? null,
    n003Expiry: example.columnValues["N003 - Reach Lift Truck"] ?? null,
    n004Expiry: example.columnValues["N004 - Lorry Mounted Lift Truck"] ?? null,
    n010Expiry: example.columnValues["N010 - Telescopic Handler"] ?? null,
    n020Expiry: example.columnValues["N020 - Tiltrotator System"] ?? null,
    n021Expiry: example.columnValues["N021 - Suction Excavator"] ?? null,
    n027Expiry:
      example.columnValues["N027 - Excavation Marshal - Banksperson"] ?? null,
    n100Expiry: example.columnValues["N100 - Exc Crane"] ?? null,
    n031Expiry: example.columnValues[ASBESTOS_MATRIX_HEADER] ?? null,
    columnValues: { ...example.columnValues },
    categoryTrainingDates: example.categoryTrainingDates ?? {},
    manualOverrideHeaders: example.manualOverrides ?? [],
    workforceId: workforce?.id ?? example.workforceItemId ?? null,
    // Prefer the row's own stored link — never re-derive it from name once set.
    workforceItemId: example.workforceItemId ?? workforce?.id ?? null,
    companyItemId: example.companyItemId ?? company?.id ?? null,
    workforceNumber:
      example.workforceNumber ?? workforce?.workforceNumber ?? null,
    exampleItemId: example.id,
  };
}

async function loadMatrixRowsWithLookups(
  workforce: AdminWorkforceRecord[],
  companies: Company[],
): Promise<MatrixRowWithLookups[]> {
  const exampleRows = await listTrainingMatrixExampleRows();
  const workforceById = new Map(workforce.map((row) => [row.id, row] as const));
  const workforceByName = new Map<string, AdminWorkforceRecord>();
  for (const row of workforce) {
    const key = nameKey(row.candidateName);
    if (key && !workforceByName.has(key)) workforceByName.set(key, row);
  }
  const companyById = new Map(companies.map((c) => [c.id, c] as const));
  const companyByName = new Map(
    companies.map((c) => [nameKey(c.companyName), c] as const),
  );

  return exampleRows.map((example) => {
    // A stored WorkforceItemId is authoritative — only fall back to a
    // name-derived guess for legacy/unlinked rows, so a linked row is never
    // reattached to a different same-name candidate.
    const linkedWorkforce = example.workforceItemId
      ? (workforceById.get(example.workforceItemId) ?? null)
      : null;
    const resolvedWorkforce =
      linkedWorkforce ??
      workforceByName.get(nameKey(example.candidateName)) ??
      null;
    const linkedCompany = example.companyItemId
      ? (companyById.get(example.companyItemId) ?? null)
      : null;
    const resolvedCompany =
      linkedCompany ??
      (resolvedWorkforce?.companyName
        ? (companyByName.get(nameKey(resolvedWorkforce.companyName)) ?? null)
        : null);
    return exampleRowToMatrix(example, resolvedWorkforce, resolvedCompany);
  });
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

/**
 * Delegates to the shared, tested findMatrixRowByWorkforce (id → legacy
 * WorkforceNumber+CompanyItemId → unambiguous unlinked name) instead of
 * reinventing name-first matching here — this is what makes register-driven
 * sync agree with the Workforce-profile-save and delete paths on which row
 * belongs to which candidate, and never adopt an ambiguous same-name row.
 */
function findMatrixRow(
  rows: MatrixRowWithLookups[],
  workforce: AdminWorkforceRecord,
  company: Company,
): MatrixRowWithLookups | null {
  const { row } = findMatrixRowByWorkforce(rows, {
    id: workforce.id,
    workforceItemId: workforce.id,
    candidateName: workforce.candidateName,
    dateOfBirth: workforce.dateOfBirth,
    workforceNumber: workforce.workforceNumber,
    companyItemId: company.id,
  });
  return row;
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
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const [matrixRows, registers] = await Promise.all([
    loadMatrixRowsWithLookups(workforce, companies),
    options.focusRecords?.length
      ? Promise.resolve(options.focusRecords)
      : listAllNormalizedRegisters(),
  ]);

  const [matrixSyncedAt, automationStatus] = await Promise.all([
    listHasColumn("trainingMatrixExample", "MatrixSyncedAt").catch(() => false),
    listHasColumn("trainingMatrixExample", "AutomationStatus").catch(
      () => false,
    ),
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
  exampleItemId: string,
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
    await updateListItemFieldsByKey(
      "trainingMatrixExample",
      exampleItemId,
      payload,
    );
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

  // Merge-safe column map starts from existing example row (or blank Name/DOB).
  const columnValues: Record<string, string | null> = {
    ...(matrixRow?.columnValues ?? {}),
    Name: workforce.candidateName,
    DOB:
      matrixRow?.columnValues?.DOB ??
      matrixRow?.dateOfBirth ??
      workforce.dateOfBirth ??
      null,
  };

  if (
    !matrixRow &&
    (!workforce.candidateName.trim() || !company.companyName.trim())
  ) {
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

  let workforceEusrExpiry: string | null | undefined;
  let workforceSwqrExpiry: string | null | undefined;
  let needsReviewForced = false;
  const manualOverrides = matrixRow?.manualOverrideHeaders ?? [];
  const categoryTrainingDates: Record<string, string | null> = {
    ...(matrixRow?.categoryTrainingDates ?? {}),
  };
  let trainingDatesDirty = false;

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
        continue;
      }
      const nporsExpiry =
        record.expiry ?? addCalendarYearsIso(record.trainingDate, 3);
      if (!nporsExpiry) {
        needsReviewForced = true;
        continue;
      }
      for (const code of record.nporsCategories) {
        const header = NPORS_HEADER_BY_CODE[code.toUpperCase()];
        if (!header) {
          warnings.push(
            `NPORS #${record.id}: category ${code} has no Training Matrix column.`,
          );
          continue;
        }
        if (isManualOverrideHeader(header, manualOverrides)) {
          warnings.push(
            `NPORS ${code}: skipped sync (manual override on matrix).`,
          );
          continue;
        }
        const existing = columnValues[header];
        if (shouldApplyPassExpiry(existing, nporsExpiry, "Pass")) {
          columnValues[header] = nporsExpiry;
          if (!fieldsUpdated.includes(header)) fieldsUpdated.push(header);
          if (record.trainingDate?.trim()) {
            categoryTrainingDates[header] = record.trainingDate
              .trim()
              .slice(0, 10);
            trainingDatesDirty = true;
          }
        } else if (existing) {
          warnings.push(
            `NPORS ${code}: kept existing matrix date (register not newer).`,
          );
        }
      }
    }

    if (record.source === "EUSR") {
      const headers = eusrHeadersForRecord(record);
      if (headers.length === 0) {
        warnings.push(
          `EUSR #${record.id}: Pass but no EUSR category mapped to a matrix column.`,
        );
      }
      const eusrExpiry =
        record.expiry ?? addCalendarYearsIso(record.trainingDate, 3);
      if (!eusrExpiry) {
        needsReviewForced = true;
        warnings.push(`EUSR #${record.id}: Pass but missing Expiry.`);
        continue;
      }
      for (const header of headers) {
        if (isManualOverrideHeader(header, manualOverrides)) {
          warnings.push(
            `${header}: skipped sync (manual override on matrix).`,
          );
          continue;
        }
        const existing = columnValues[header];
        if (shouldApplyPassExpiry(existing, eusrExpiry, "Pass")) {
          stampEusrCategoryDates(
            header,
            eusrExpiry,
            record.trainingDate,
            columnValues,
            categoryTrainingDates,
            fieldsUpdated,
          );
          trainingDatesDirty = true;
        } else if (existing) {
          warnings.push(
            `${header}: kept existing matrix date (register not newer).`,
          );
        }
      }
      if (
        workforceEusrExpiry === undefined ||
        shouldApplyPassExpiry(workforceEusrExpiry, eusrExpiry, "Pass")
      ) {
        workforceEusrExpiry = eusrExpiry;
      }
    }

    if (record.source === "NRSWA") {
      const nrswaExpiry =
        record.expiry ?? addCalendarYearsIso(record.trainingDate, 5);
      if (!nrswaExpiry) {
        continue;
      }
      if (
        workforceSwqrExpiry === undefined ||
        shouldApplyPassExpiry(workforceSwqrExpiry, nrswaExpiry, "Pass")
      ) {
        workforceSwqrExpiry = nrswaExpiry;
      }
    }

    if (record.source === "In-House") {
      const category =
        record.certificateCategory || record.courseCategory || null;
      if (!isAsbestosAwarenessCategory(category)) {
        continue;
      }
      if (!record.expiry) {
        needsReviewForced = true;
        warnings.push(
          `In-House #${record.id}: Asbestos Awareness Pass but missing Expiry.`,
        );
        continue;
      }
      if (isManualOverrideHeader(ASBESTOS_MATRIX_HEADER, manualOverrides)) {
        warnings.push(
          `N031 Asbestos Awareness: skipped sync (manual override on matrix).`,
        );
        continue;
      }
      const existing = columnValues[ASBESTOS_MATRIX_HEADER];
      if (shouldApplyPassExpiry(existing, record.expiry, "Pass")) {
        columnValues[ASBESTOS_MATRIX_HEADER] = record.expiry;
        if (!fieldsUpdated.includes(ASBESTOS_MATRIX_HEADER)) {
          fieldsUpdated.push(ASBESTOS_MATRIX_HEADER);
        }
      } else if (existing) {
        warnings.push(
          `N031 Asbestos Awareness: kept existing matrix date (register not newer).`,
        );
      }
    }
  }

  if (workforceEusrExpiry) {
    if (isManualOverrideHeader("EUSR Expiry", manualOverrides)) {
      warnings.push("EUSR Expiry: skipped sync (manual override on matrix).");
    } else {
      const existing = columnValues["EUSR Expiry"];
      if (shouldApplyPassExpiry(existing, workforceEusrExpiry, "Pass")) {
        columnValues["EUSR Expiry"] = workforceEusrExpiry;
        if (!fieldsUpdated.includes("EUSR Expiry")) {
          fieldsUpdated.push("EUSR Expiry");
        }
      }
    }
  }

  if (workforceSwqrExpiry) {
    if (isManualOverrideHeader("NRSWA Expiry", manualOverrides)) {
      warnings.push("NRSWA Expiry: skipped sync (manual override on matrix).");
    } else {
      const existing = columnValues["NRSWA Expiry"];
      if (shouldApplyPassExpiry(existing, workforceSwqrExpiry, "Pass")) {
        columnValues["NRSWA Expiry"] = workforceSwqrExpiry;
        if (!fieldsUpdated.includes("NRSWA Expiry")) {
          fieldsUpdated.push("NRSWA Expiry");
        }
      }
    }
  }

  const statusDates = [
    ...Object.entries(columnValues)
      .filter(([key]) => key !== "Name" && key !== "DOB")
      .map(([, value]) => value),
    workforceEusrExpiry,
    workforceSwqrExpiry,
  ];
  const status = computeMatrixStatusFromDates(statusDates);
  const needsReview = needsReviewForced || status.needsReview;

  if (ctx.dryRun) {
    if (needsReview) fieldsUpdated.push("needsReview");
    if (status.nextExpiryDate) fieldsUpdated.push("nextExpiryDate");
  }

  const willCreate = !matrixRow && fieldsUpdated.length > 0;
  if (willCreate) {
    matrixRowCreated = true;
    if (ctx.dryRun && !fieldsUpdated.includes("(create)")) {
      fieldsUpdated.push("Name", "(create)");
    }
  }

  if (!ctx.dryRun && fieldsUpdated.length > 0) {
    try {
      // Stamp the strong link columns on every register-driven write too, not
      // just Workforce-profile saves — so a row this sync creates or touches
      // is never left name-only (and thus vulnerable to a same-name mismatch
      // on the next sync).
      const { linkFields } = buildWorkforceMatrixSource({
        id: workforce.id,
        workforceItemId: workforce.id,
        candidateName: workforce.candidateName,
        companyName: company.companyName,
        companyItemId: company.id,
        companyNumber: company.companyNumber,
        workforceNumber: workforce.workforceNumber,
        dateOfBirth: workforce.dateOfBirth,
      });
      const upserted = await upsertTrainingMatrixExampleRow({
        candidateName: workforce.candidateName,
        existingItemId:
          matrixRow?.exampleItemId ??
          stripExampleMatrixId(matrixRow?.id) ??
          null,
        source: columnValues,
        linkFields,
        // Do not touch ManualOverrides — sync never clears admin flags.
        ...(trainingDatesDirty ? { categoryTrainingDates } : {}),
      });
      const nextRow: MatrixRowWithLookups = {
        id: `example:${upserted.id}`,
        candidateName: workforce.candidateName,
        companyName: company.companyName,
        department: workforce.department,
        dateOfBirth: workforce.dateOfBirth ?? columnValues.DOB ?? null,
        overallStatus: status.overallStatus,
        needsReview,
        matrixNotes: null,
        nextExpiryDate: status.nextExpiryDate,
        n001Expiry: columnValues["N001 - Ind FLT"] ?? null,
        n003Expiry: columnValues["N003 - Reach Lift Truck"] ?? null,
        n004Expiry: columnValues["N004 - Lorry Mounted Lift Truck"] ?? null,
        n010Expiry: columnValues["N010 - Telescopic Handler"] ?? null,
        n020Expiry: columnValues["N020 - Tiltrotator System"] ?? null,
        n021Expiry: columnValues["N021 - Suction Excavator"] ?? null,
        n027Expiry:
          columnValues["N027 - Excavation Marshal - Banksperson"] ?? null,
        n100Expiry: columnValues["N100 - Exc Crane"] ?? null,
        n031Expiry: columnValues[ASBESTOS_MATRIX_HEADER] ?? null,
        columnValues: { ...columnValues },
        categoryTrainingDates: { ...categoryTrainingDates },
        manualOverrideHeaders: manualOverrides,
        workforceId: workforce.id,
        workforceItemId: workforce.id,
        companyItemId: company.id,
        workforceNumber: workforce.workforceNumber,
        exampleItemId: upserted.id,
      };
      if (matrixRow) {
        Object.assign(matrixRow, nextRow);
      } else {
        matrixRow = nextRow;
        ctx.matrixRows.push(matrixRow);
      }
      if (upserted.created) {
        matrixRowCreated = true;
        if (!fieldsUpdated.includes("Name")) fieldsUpdated.push("Name");
      }
    } catch (error) {
      errors.push(
        error instanceof Error
          ? error.message
          : "Failed to update Training Matrix row.",
      );
    }
  } else if (ctx.dryRun && willCreate) {
    matrixRow = {
      id: "dry-run-new",
      candidateName: workforce.candidateName,
      companyName: company.companyName,
      department: workforce.department,
      dateOfBirth: workforce.dateOfBirth,
      overallStatus: status.overallStatus,
      needsReview,
      matrixNotes: null,
      nextExpiryDate: status.nextExpiryDate,
      n001Expiry: columnValues["N001 - Ind FLT"] ?? null,
      n003Expiry: columnValues["N003 - Reach Lift Truck"] ?? null,
      n004Expiry: columnValues["N004 - Lorry Mounted Lift Truck"] ?? null,
      n010Expiry: columnValues["N010 - Telescopic Handler"] ?? null,
      n020Expiry: columnValues["N020 - Tiltrotator System"] ?? null,
      n021Expiry: columnValues["N021 - Suction Excavator"] ?? null,
      n027Expiry:
        columnValues["N027 - Excavation Marshal - Banksperson"] ?? null,
      n100Expiry: columnValues["N100 - Exc Crane"] ?? null,
      columnValues: { ...columnValues },
      categoryTrainingDates: { ...categoryTrainingDates },
      workforceItemId: workforce.id,
      companyItemId: company.id,
      workforceNumber: workforce.workforceNumber,
      exampleItemId: null,
    };
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

  if (
    matrixRow?.exampleItemId &&
    (fieldsUpdated.length > 0 || matrixRowCreated)
  ) {
    try {
      const meta = await writeOptionalSyncMeta(
        matrixRow.exampleItemId,
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
    if (row.workforceItemId) candidateIds.add(row.workforceItemId);
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
  // In-House only syncs Asbestos Awareness → N031; other courses stay standalone.
  if (registerKey === "inHouseCertificates") {
    const asbestos =
      isAsbestosAwarenessCategory(record.certificateCategory) ||
      isAsbestosAwarenessCategory(record.courseCategory) ||
      isAsbestosAwarenessCategory(record.course);
    if (!asbestos) {
      const items = [
        emptyResultItem({
          candidate: record.candidateName,
          company: record.companyName,
          registerSources: [],
          skipped: true,
        }),
      ];
      return {
        dryRun: Boolean(options.dryRun),
        scope: "register",
        items,
        summary: buildSummary(items),
      };
    }
  }

  // mapRegister already resolved the strong lookup ids from the saved
  // SharePoint item (see adminCrudService.ts resolveRegisterPeople) — no need
  // to re-read the item here just to re-extract them.
  const focus = normalizeRegisterFromAdminRecord(registerKey, record, {
    candidateLookupId: record.candidateLookupId,
    companyLookupId: record.companyLookupId,
  });
  if (focus.source === "NRSWA" && !focus.expiry && focus.trainingDate) {
    focus.expiry = addCalendarYearsIso(focus.trainingDate, 5);
  }
  if (
    (focus.source === "NPORS" || focus.source === "EUSR") &&
    !focus.expiry &&
    focus.trainingDate
  ) {
    focus.expiry = addCalendarYearsIso(focus.trainingDate, 3);
  }

  const ctx = await loadSyncContext({ ...options, focusRecords: [focus] });
  // Ensure we have full register set for status — reload all for this candidate path.
  ctx.registers = await listAllNormalizedRegisters();
  // Prefer the just-saved form values (e.g. NPORS category) when SharePoint
  // MultiChoice writes are blocked for app-only Graph.
  ctx.registers = [
    ...ctx.registers.filter(
      (row) => !(row.registerKey === registerKey && row.id === record.id),
    ),
    focus,
  ];

  const company = findCompany(ctx.companies, {
    companyLookupId: focus.companyLookupId,
    companyName: focus.companyName,
  });
  const workforce = findWorkforce(ctx.workforce, {
    candidateLookupId: focus.candidateLookupId ?? record.workforceId,
    workforceNumber: record.workforceNumber,
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
  // Register save already succeeded. Category/expiry notes from sibling
  // rows must not surface as error toasts on Streetworks / NPORS / EUSR.
  for (const item of result.items) {
    item.warnings = [];
    if (!item.errors.length) item.skipReason = undefined;
  }
  result.summary.warnings = 0;
  return result;
}

/**
 * Recompute the Training Matrix field(s) fed by `deletedRegisterKey` for one
 * candidate, from whatever valid (Pass) source records remain — per the
 * source-aware refresh rules in registerMatrixFieldSync.ts: manual overrides
 * are preserved and flagged, never blindly blanked; a system-synced value is
 * cleared and flagged Needs Review only when no source remains to support it.
 */
async function recomputeCandidateMatrixAfterDelete(
  ctx: SyncContext,
  resolved: ResolvedCandidate,
  deletedRegisterKey: AdminRegisterKey,
): Promise<MatrixSyncResultItem> {
  const { workforce, company } = resolved;
  const warnings: string[] = [];
  const errors: string[] = [];
  const fieldsUpdated: string[] = [];

  const matrixRow = findMatrixRow(ctx.matrixRows, workforce, company);
  if (!matrixRow) {
    return emptyResultItem({
      candidate: workforce.candidateName,
      company: company.companyName,
      candidateId: workforce.id,
      companyId: company.id,
      registerSources: [registerKeyToSource(deletedRegisterKey)],
      skipped: true,
      skipReason:
        "No Training Matrix row for this candidate — nothing to recompute.",
    });
  }

  // ctx.registers is loaded fresh from SharePoint after the delete completed
  // (the API route awaits the delete before calling this), so the deleted
  // record is already absent — this is the authoritative remaining set.
  const remainingForCandidate = recordsForCandidate(
    ctx.registers,
    workforce,
    company,
  );
  const manualOverrides = matrixRow.manualOverrideHeaders ?? [];
  const columnValues: Record<string, string | null> = {
    ...matrixRow.columnValues,
  };
  const categoryTrainingDates: Record<string, string | null> = {
    ...(matrixRow.categoryTrainingDates ?? {}),
  };
  let trainingDatesDirty = false;

  const applyHeaderRecompute = (
    header: string,
    remainingRecords: NormalizedRegisterRecord[],
  ) => {
    const outcome = computeMatrixFieldAfterRemoval({
      header,
      currentValue: columnValues[header] ?? null,
      isManualOverride: isManualOverrideHeader(header, manualOverrides),
      remainingRecords,
    });
    if (outcome.action !== "unchanged") warnings.push(outcome.note);
    if (outcome.action === "recomputed" || outcome.action === "cleared") {
      columnValues[header] = outcome.nextValue;
      if (!fieldsUpdated.includes(header)) fieldsUpdated.push(header);
    }
  };

  let workforceEusrExpiry: string | null | undefined;
  let workforceSwqrExpiry: string | null | undefined;

  if (deletedRegisterKey === "nporsRegister") {
    const nporsRecords = remainingForCandidate.filter(
      (r) => r.source === "NPORS",
    );
    // Recompute every NPORS header the candidate still has data for, plus any
    // header the row currently holds — covers the deleted record being the
    // last one backing that category.
    const headers = new Set<string>();
    for (const record of nporsRecords) {
      for (const code of record.nporsCategories) {
        const header = NPORS_HEADER_BY_CODE[code.toUpperCase()];
        if (header) headers.add(header);
      }
    }
    for (const header of Object.values(NPORS_HEADER_BY_CODE)) {
      if (columnValues[header]) headers.add(header);
    }
    for (const header of headers) {
      const remaining = nporsRecords.filter((r) =>
        r.nporsCategories.some(
          (code) => NPORS_HEADER_BY_CODE[code.toUpperCase()] === header,
        ),
      );
      applyHeaderRecompute(header, remaining);
    }
  }

  if (deletedRegisterKey === "eusrRegister") {
    const eusrRecords = remainingForCandidate.filter(
      (r) => r.source === "EUSR",
    );
    const headers = new Set<string>();
    for (const column of EUSR_MATRIX_CATEGORY_COLUMNS) {
      if (
        columnValues[column.header] ||
        categoryTrainingDates[column.header] ||
        categoryTrainingDates[matrixStoredExpiryKey(column.header)]
      ) {
        headers.add(column.header);
      }
    }
    for (const record of eusrRecords) {
      for (const header of eusrHeadersForRecord(record)) {
        headers.add(header);
      }
    }
    for (const key of Object.keys(columnValues)) {
      if (key.startsWith("EUSR - ") && columnValues[key]) headers.add(key);
    }
    for (const header of headers) {
      const remaining = eusrRecords.filter((record) =>
        eusrHeadersForRecord(record).includes(header),
      );
      const previousExpiry = columnValues[header] ?? null;
      const previousTraining = categoryTrainingDates[header] ?? null;
      applyHeaderRecompute(header, remaining);
      const latest = remaining
        .filter((record) => record.trainingOutcome === "Pass" && record.expiry)
        .sort((a, b) => {
          const aMs = new Date(a.expiry ?? "").getTime();
          const bMs = new Date(b.expiry ?? "").getTime();
          return bMs - aMs;
        })[0];
      const nextExpiry = columnValues[header] ?? null;
      const nextTraining = latest?.trainingDate ?? null;
      if (previousExpiry !== nextExpiry || previousTraining !== nextTraining) {
        stampEusrCategoryDates(
          header,
          nextExpiry,
          nextTraining,
          columnValues,
          categoryTrainingDates,
          fieldsUpdated,
        );
        trainingDatesDirty = true;
      }
    }
    applyHeaderRecompute("EUSR Expiry", eusrRecords);
    if (fieldsUpdated.includes("EUSR Expiry")) {
      workforceEusrExpiry = latestPassExpiry(eusrRecords);
    }
  }

  if (deletedRegisterKey === "nrswaRegister") {
    const nrswaRecords = remainingForCandidate.filter(
      (r) => r.source === "NRSWA",
    );
    applyHeaderRecompute("NRSWA Expiry", nrswaRecords);
    if (fieldsUpdated.includes("NRSWA Expiry")) {
      workforceSwqrExpiry = latestPassExpiry(nrswaRecords);
    }
  }

  if (deletedRegisterKey === "inHouseCertificates") {
    const asbestosRecords = remainingForCandidate.filter(
      (r) =>
        r.source === "In-House" &&
        isAsbestosAwarenessCategory(r.certificateCategory || r.courseCategory),
    );
    applyHeaderRecompute(ASBESTOS_MATRIX_HEADER, asbestosRecords);
  }

  if (fieldsUpdated.length === 0) {
    return emptyResultItem({
      candidate: workforce.candidateName,
      company: company.companyName,
      candidateId: workforce.id,
      companyId: company.id,
      matrixRowId: matrixRow.id,
      matrixRowFound: true,
      registerSources: [registerKeyToSource(deletedRegisterKey)],
      warnings,
      skipped: true,
      skipReason: "No matrix field changes required after delete.",
    });
  }

  if (ctx.dryRun) {
    return emptyResultItem({
      candidate: workforce.candidateName,
      company: company.companyName,
      candidateId: workforce.id,
      companyId: company.id,
      matrixRowId: matrixRow.id,
      matrixRowFound: true,
      registerSources: [registerKeyToSource(deletedRegisterKey)],
      fieldsUpdated,
      warnings,
    });
  }

  try {
    await upsertTrainingMatrixExampleRow({
      candidateName: workforce.candidateName,
      existingItemId:
        matrixRow.exampleItemId ?? stripExampleMatrixId(matrixRow.id),
      source: columnValues,
      // Do not touch ManualOverrides — preserved overrides stay flagged, not cleared.
      ...(trainingDatesDirty ? { categoryTrainingDates } : {}),
    });
  } catch (error) {
    errors.push(
      error instanceof Error
        ? error.message
        : "Failed to update Training Matrix row after delete.",
    );
  }

  if (workforceEusrExpiry !== undefined) {
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
  }

  if (workforceSwqrExpiry !== undefined) {
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
  }

  return emptyResultItem({
    candidate: workforce.candidateName,
    company: company.companyName,
    candidateId: workforce.id,
    companyId: company.id,
    matrixRowId: matrixRow.id,
    matrixRowFound: true,
    registerSources: [registerKeyToSource(deletedRegisterKey)],
    fieldsUpdated: [...new Set(fieldsUpdated)],
    warnings,
    errors,
    skipped: false,
  });
}

/**
 * After a register record is deleted: recompute (or clear) the Matrix
 * field(s) it fed, rather than leaving a stale system-synced expiry with no
 * backing record. In-House non-Asbestos deletes have no Matrix target and
 * are reported skipped, matching the create/update behavior for the same
 * courses.
 */
export async function syncAfterRegisterDelete(
  registerKey: AdminRegisterKey,
  deletedRecord: AdminTrainingRecord,
  options: MatrixSyncOptions = {},
): Promise<MatrixSyncResult> {
  if (registerKey === "inHouseCertificates") {
    const asbestos =
      isAsbestosAwarenessCategory(deletedRecord.certificateCategory) ||
      isAsbestosAwarenessCategory(deletedRecord.courseCategory) ||
      isAsbestosAwarenessCategory(deletedRecord.course);
    if (!asbestos) {
      const items = [
        emptyResultItem({
          candidate: deletedRecord.candidateName,
          company: deletedRecord.companyName,
          registerSources: [],
          skipped: true,
          skipReason:
            "Deleted In-House course was standalone (only Asbestos Awareness affects N031).",
        }),
      ];
      return {
        dryRun: Boolean(options.dryRun),
        scope: "register-delete",
        items,
        summary: buildSummary(items),
      };
    }
  }

  const ctx = await loadSyncContext(options);
  const company = findCompany(ctx.companies, {
    companyLookupId: deletedRecord.companyLookupId,
    companyName: deletedRecord.companyName,
  });
  const workforce = findWorkforce(ctx.workforce, {
    candidateLookupId: deletedRecord.candidateLookupId,
    candidateName: deletedRecord.candidateName,
    companyName: company?.companyName ?? deletedRecord.companyName,
  });

  const items: MatrixSyncResultItem[] = [];
  if (!workforce || !company) {
    items.push(
      emptyResultItem({
        candidate: deletedRecord.candidateName,
        company: deletedRecord.companyName,
        registerSources: [registerKeyToSource(registerKey)],
        skipped: true,
        skipReason:
          "Could not resolve candidate/company for matrix recompute after delete.",
      }),
    );
  } else {
    items.push(
      await recomputeCandidateMatrixAfterDelete(
        ctx,
        { workforce, company },
        registerKey,
      ),
    );
  }

  const result: MatrixSyncResult = {
    dryRun: Boolean(options.dryRun),
    scope: "register-delete",
    items,
    summary: buildSummary(items),
  };
  await auditSync(
    options.userEmail,
    `register-delete/${registerKey}/${deletedRecord.id}`,
    result,
  );
  return result;
}

/** Load registers for one key — used by tests / diagnostics. */
export { listNormalizedRegisters };
