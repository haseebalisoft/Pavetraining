import "server-only";

import { allocateNextCompanyNumber } from "@/lib/companyNumber";
import { allocateNextWorkforceNumber } from "@/lib/workforceNumber";
import {
  createAdminCompany,
  createAdminWorkforce,
  ensurePermissionPerson,
  findPermissionPerson,
  listAdminCompanies,
  listAdminWorkforce,
  loadPermissionPeople,
  toMatrixProfile,
  updateAdminWorkforce,
  type AdminWorkforceRecord,
  type PermissionPersonRef,
} from "@/lib/services/adminCrudService";
import {
  withBulkSharePointWrites,
} from "@/lib/services/sharePointListService";
import {
  findCandidateDuplicate,
} from "@/lib/services/bulkUpload/matching";
import { resolveWorkforceCompanyMatch } from "@/lib/services/bulkUpload/companyMatch";
import {
  createBulkLogger,
  type BulkLogger,
} from "@/lib/services/bulkUpload/bulkUploadLog";
import {
  normalizeDateValue,
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkImportSummary,
  BulkPreviewRow,
} from "@/types/bulkUpload";
import type { Company } from "@/types/models";

const NAME_ALIASES = [
  "Candidate Name",
  "CandidateName",
  "Name",
  "Full Name",
];
const COMPANY_ALIASES = ["Company", "Company Name", "CompanyName"];
const COMPANY_NUMBER_ALIASES = [
  "Company Number",
  "CompanyNumber",
  "Company No",
];
const DEPT_ALIASES = ["Department", "Dept", " Department"];
const DOB_ALIASES = [
  "DOB",
  "Date of Birth",
  "Date of birth",
  "DateOfBirth",
  "Birth Date",
];
const WF_ALIASES = [
  "Workforce Number",
  "WorkforceNumber",
  "Workforce No",
  "Workforce No.",
];
const STATUS_ALIASES = ["Status"];
const TM_ALIASES = ["Training Manager", "Training manager", "TrainingManager"];
const SUPERVISOR_ALIASES = ["Supervisor"];

function emptySummary(): BulkImportSummary {
  return {
    totalRows: 0,
    readyRows: 0,
    warningRows: 0,
    duplicateRows: 0,
    errorRows: 0,
    skippedRows: 0,
    importedRows: 0,
  };
}

export function summarizeBulkRows(rows: BulkPreviewRow[]): BulkImportSummary {
  const summary = emptySummary();
  summary.totalRows = rows.length;
  for (const row of rows) {
    switch (row.status) {
      case "Ready":
        summary.readyRows += 1;
        break;
      case "Warning":
        summary.warningRows += 1;
        break;
      case "Duplicate":
        summary.duplicateRows += 1;
        break;
      case "Error":
        summary.errorRows += 1;
        break;
      case "Skipped":
        summary.skippedRows += 1;
        break;
      case "Imported":
        summary.importedRows += 1;
        break;
      default:
        break;
    }
  }
  return summary;
}

function mapCandidateFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  return {
    candidateName: pickField(raw, NAME_ALIASES),
    company: pickField(raw, COMPANY_ALIASES),
    companyNumber: pickField(raw, COMPANY_NUMBER_ALIASES),
    department: pickField(raw, DEPT_ALIASES),
    dateOfBirth: normalizeDateValue(pickField(raw, DOB_ALIASES)),
    workforceNumber: pickField(raw, WF_ALIASES),
    status: pickField(raw, STATUS_ALIASES),
    trainingManager: pickField(raw, TM_ALIASES),
    supervisor: pickField(raw, SUPERVISOR_ALIASES),
    candidateAddress: pickField(raw, [
      "Candidate Address",
      "CandidateAddress",
      "Address",
    ]),
    email: pickField(raw, ["Email", "E-mail"]),
    contactNumber: pickField(raw, [
      "Contact number",
      "Contact Number",
      "Phone",
      "Mobile",
    ]),
    niNumber: pickField(raw, ["Ni Number", "NI Number", "NiNumber", "NI"]),
    nporsNumbers: pickField(raw, [
      "NPORS Number",
      "NPORS Numbers",
      "NPORSNumbers",
    ]),
    cscsNumber: pickField(raw, ["CSCS Number", "CSCSNumber"]),
    cscsExpiry: normalizeDateValue(
      pickField(raw, ["Cscs Expiry", "CSCS Expiry", "CscsExpiry"]),
    ),
    swqrNumber: pickField(raw, ["SWQR Number", "SWQRNumber"]),
    swqrExpiry: normalizeDateValue(
      pickField(raw, ["Swqr Expiry", "SWQR Expiry", "SwqrExpiry"]),
    ),
    eusrNumber: pickField(raw, ["EUSR Number", "EUSRNumber"]),
    eusrExpiry: normalizeDateValue(
      pickField(raw, ["Eusr Expiry", "EUSR Expiry", "EusrExpiry"]),
    ),
    inHouseCertificationNumber: pickField(raw, [
      "In House Certification Number",
      "In-House Certification Number",
      "InHouseCertificationNumber",
    ]),
    notes: pickField(raw, ["Notes"]),
  };
}

function workforceWritePayload(
  fields: Record<string, string | null>,
  resolvedCompanyName: string,
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    candidateName: fields.candidateName,
    companyName: resolvedCompanyName,
  };
  const assign = (key: string, value: string | null | undefined) => {
    if (value?.trim()) payload[key] = value.trim();
  };
  assign("workforceNumber", fields.workforceNumber);
  assign("dateOfBirth", fields.dateOfBirth);
  // Department is a Lookup → Departments list (resolved in create/update).
  assign("department", fields.department);
  assign("departmentText", fields.department);
  // SharePoint Workforce Status choices: Active | inactive (blank Excel → Active).
  payload.status = fields.status?.trim() || "Active";
  assign("trainingManager", fields.trainingManager);
  assign("supervisor", fields.supervisor);
  assign("candidateAddress", fields.candidateAddress);
  assign("email", fields.email);
  assign("contactNumber", fields.contactNumber);
  assign("niNumber", fields.niNumber);
  assign("nporsNumbers", fields.nporsNumbers);
  assign("cscsNumber", fields.cscsNumber);
  assign("cscsExpiry", fields.cscsExpiry);
  assign("swqrNumber", fields.swqrNumber);
  assign("swqrExpiry", fields.swqrExpiry);
  assign("eusrNumber", fields.eusrNumber);
  assign("eusrExpiry", fields.eusrExpiry);
  assign("inHouseCertificationNumber", fields.inHouseCertificationNumber);
  assign("notes", fields.notes);
  return payload;
}

function buildNonBlankUpdate(
  fields: Record<string, string | null>,
  resolvedCompanyName: string,
): Record<string, unknown> {
  const full = workforceWritePayload(fields, resolvedCompanyName);
  const payload: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(full)) {
    if (value === null || value === undefined) continue;
    if (typeof value === "string" && !value.trim()) continue;
    payload[key] = value;
  }
  return payload;
}

function validateCandidateRow(
  rowNumber: number,
  fields: Record<string, string | null>,
  companies: Company[],
  workforce: AdminWorkforceRecord[],
  people: PermissionPersonRef[],
  allocatedWorkforceNumbers: string[] = [],
  options: { autoCreateMissing: boolean } = { autoCreateMissing: false },
): BulkPreviewRow {
  const messages: string[] = [];
  const candidateName = fields.candidateName?.trim() ?? "";
  const companyInput = fields.company?.trim() ?? "";
  const companyNumberInput = fields.companyNumber?.trim() ?? "";
  let workforceNumber = fields.workforceNumber?.trim() ?? "";
  let autoAssignedWorkforceNumber = false;

  if (!candidateName) {
    messages.push("Candidate Name is required.");
  }
  if (!fields.email?.trim()) {
    messages.push(
      "Email is required in SharePoint Workforce List — row will fail without it.",
    );
  }

  // Company matching: Company Number first, then a UNIQUE Company Name. Never
  // assign by fuzzy name when a number is present; never auto-match ambiguous
  // names into a silent duplicate company.
  const match = resolveWorkforceCompanyMatch(
    companies,
    { companyNumber: companyNumberInput, companyName: companyInput },
    options,
  );
  const report = match.report;
  const matchedCompany = match.kind === "matched" ? match.company : null;
  const resolvedCompanyName = matchedCompany?.companyName ?? companyInput;
  const companyMatchedBy: BulkPreviewRow["companyMatchedBy"] =
    match.kind === "matched"
      ? match.matchedBy
      : match.kind === "create"
        ? "create"
        : null;
  const reportFields = {
    matchedCompanyId: report.matchedCompanyId,
    matchedCompanyNumber:
      report.matchedCompanyNumber ??
      matchedCompany?.companyNumber ??
      (companyNumberInput || null),
    matchedCompanyName: report.matchedCompanyName ?? resolvedCompanyName ?? null,
    companyMatchedBy,
  };
  if (report.warning) messages.push(report.warning);

  const department = fields.department?.trim();
  if (department) {
    messages.push(
      `Department "${department}" will be linked (created for the company if missing).`,
    );
  }

  const tm = fields.trainingManager?.trim();
  if (tm && !findPermissionPerson(people, tm)) {
    messages.push(
      `Training manager "${tm}" not in Permissions — will be created on import.`,
    );
  }
  const supervisor = fields.supervisor?.trim();
  if (supervisor && !findPermissionPerson(people, supervisor)) {
    messages.push(
      `Supervisor "${supervisor}" not in Permissions — will be created on import.`,
    );
  }

  const hasCompanyInput = Boolean(companyInput || companyNumberInput);
  if (!workforceNumber && candidateName && hasCompanyInput) {
    workforceNumber = allocateNextWorkforceNumber(workforce, [
      ...allocatedWorkforceNumbers,
    ]);
    autoAssignedWorkforceNumber = true;
    messages.push(
      `Workforce Number will be auto-assigned: ${workforceNumber}.`,
    );
  }
  if (
    workforceNumber &&
    !allocatedWorkforceNumbers.some(
      (value) => value.toLowerCase() === workforceNumber.toLowerCase(),
    )
  ) {
    allocatedWorkforceNumbers.push(workforceNumber);
  }

  const fieldsWithNumber = {
    ...fields,
    workforceNumber: workforceNumber || fields.workforceNumber,
    company: resolvedCompanyName,
  };

  // Hard failures: missing candidate/email, or an unresolved company (wrong
  // number, ambiguous name, or missing company without auto-create).
  if (!candidateName || !fields.email?.trim() || match.kind === "error") {
    if (match.kind === "error" && report.error) messages.push(report.error);
    return {
      rowNumber,
      status: "Error",
      messages,
      fields: fieldsWithNumber,
      resolvedCompanyName: matchedCompany?.companyName ?? null,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
      ...reportFields,
    };
  }

  const duplicate = findCandidateDuplicate(workforce, {
    candidateName,
    companyName: resolvedCompanyName,
    workforceNumber: workforceNumber || null,
    dateOfBirth: fields.dateOfBirth,
  });

  if (duplicate?.kind === "workforceNumber") {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate: workforce number matches existing candidate "${duplicate.record.candidateName}".`,
      ],
      fields: fieldsWithNumber,
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "workforceNumber",
      ...reportFields,
    };
  }

  if (duplicate?.kind === "nameDobCompany") {
    return {
      rowNumber,
      status: "Duplicate",
      messages: [
        `Duplicate: name + DOB + company matches existing candidate "${duplicate.record.candidateName}".`,
      ],
      fields: fieldsWithNumber,
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "nameDobCompany",
      ...reportFields,
    };
  }

  if (duplicate?.kind === "nameCompany") {
    return {
      rowNumber,
      status: "Warning",
      messages: [
        `Possible match: name + company matches existing candidate "${duplicate.record.candidateName}". Import will create a new record unless you choose Update existing.`,
      ],
      fields: fieldsWithNumber,
      resolvedCompanyName,
      matchedEntityId: duplicate.record.id,
      matchedEntityName: duplicate.record.candidateName,
      duplicateMatch: "nameCompany",
      ...reportFields,
    };
  }

  if (match.kind === "create") {
    // Auto-create mode: company will be created on commit.
    return {
      rowNumber,
      status: "Warning",
      messages,
      fields: fieldsWithNumber,
      resolvedCompanyName,
      matchedEntityId: null,
      matchedEntityName: null,
      duplicateMatch: null,
      ...reportFields,
    };
  }

  // Informative notes (not blockers): keep status Ready unless a real risk
  // was already flagged above (soft name match / company create).
  if (!autoAssignedWorkforceNumber && !workforceNumber) {
    messages.push("Workforce Number is missing (optional) — can be auto-assigned.");
  }
  if (!fields.dateOfBirth) {
    messages.push("DOB is missing (optional).");
  }

  return {
    rowNumber,
    status: "Ready",
    messages,
    fields: fieldsWithNumber,
    resolvedCompanyName,
    matchedEntityId: null,
    matchedEntityName: null,
    duplicateMatch: null,
    ...reportFields,
  };
}

/**
 * Resolve the row's company by Number-first / unique-Name matching. Returns the
 * matched Company List item, or (only when auto-create is enabled) creates it.
 * Throws with the matcher's error message when the company cannot be resolved,
 * so the row is reported as an Error instead of silently mis-assigned.
 */
async function resolveOrCreateCompany(
  companies: Company[],
  companyName: string,
  companyNumber: string | null,
  options: { autoCreateMissing: boolean },
): Promise<{ companies: Company[]; company: Company }> {
  const match = resolveWorkforceCompanyMatch(
    companies,
    { companyNumber, companyName },
    options,
  );

  if (match.kind === "matched") {
    const existing =
      companies.find((c) => c.id === match.company.id) ?? null;
    if (existing) return { companies, company: existing };
    // Matched a ref that is not in the local array (shouldn't happen) — fall
    // through to a fresh lookup by id below.
    const byId = companies.find((c) => c.id === match.company.id);
    if (byId) return { companies, company: byId };
  }

  if (match.kind === "error") {
    throw new ValidationErrorLike(
      match.report.error ?? "Company could not be resolved.",
    );
  }

  // match.kind === "create" — auto-create the missing company. Preserve a
  // supplied Company Number; otherwise allocate a fresh one.
  const created = await createAdminCompany({
    companyName,
    companyNumber:
      companyNumber?.trim() || allocateNextCompanyNumber(companies),
    status: "Active",
    existingCompanies: companies,
    // Bulk: skip OneDrive/SharePoint folder provisioning per new company.
    bulkMode: true,
  });
  return { companies: [...companies, created], company: created };
}

/** Local error carrier so the commit loop reports the matcher message per row. */
class ValidationErrorLike extends Error {}

/** Bound parallel Graph POSTs — ~5× faster than serial for workforce creates. */
const BULK_CREATE_CONCURRENCY = 5;

async function mapPool<T, R>(
  items: T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  if (items.length === 0) return [];
  const results = new Array<R>(items.length);
  let nextIndex = 0;
  const workerCount = Math.min(concurrency, items.length);

  await Promise.all(
    Array.from({ length: workerCount }, async () => {
      while (true) {
        const index = nextIndex;
        nextIndex += 1;
        if (index >= items.length) return;
        results[index] = await mapper(items[index], index);
      }
    }),
  );

  return results;
}

export async function previewCandidateImport(
  spreadsheet: ParsedSpreadsheet,
  options: { autoCreateMissingCompanies?: boolean } = {},
): Promise<BulkPreviewRow[]> {
  const autoCreateMissing = options.autoCreateMissingCompanies ?? false;
  const [companies, workforce, people] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
    loadPermissionPeople(),
  ]);
  const allocatedWorkforceNumbers: string[] = [];

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapCandidateFields(raw);
    const validated = validateCandidateRow(
      index + 2,
      fields,
      companies,
      workforce,
      people,
      allocatedWorkforceNumbers,
      { autoCreateMissing },
    );
    return {
      ...validated,
      source: raw,
    };
  });
}

export async function commitCandidateImport(input: {
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
  autoCreateMissingCompanies?: boolean;
  log?: BulkLogger;
}): Promise<BulkPreviewRow[]> {
  const autoCreateMissing = input.autoCreateMissingCompanies ?? false;
  const log = input.log ?? createBulkLogger("commit:workforce");
  // Load all reference data in one parallel round-trip (companies used to be a
  // separate serial read before this Promise.all).
  const loadPhase = log.phase("load");
  const [companiesInitial, workforce, initialPeople, initialDepartments] =
    await Promise.all([
      listAdminCompanies(),
      listAdminWorkforce(),
      loadPermissionPeople(),
      import("@/lib/services/departmentService").then((mod) =>
        mod.listAdminDepartments(),
      ),
    ]);
  let companies = companiesInitial;
  loadPhase.end({
    companies: companies.length,
    workforce: workforce.length,
    people: initialPeople.length,
    departments: initialDepartments.length,
  });
  let people = initialPeople;
  let departmentRecords = initialDepartments.map((row) => ({
    id: row.id,
    name: row.name,
    companyId: row.companyId,
    companyName: row.companyName,
  }));
  const liveWorkforce = [...workforce];
  const knownWorkforceNumbers = new Set(
    liveWorkforce
      .map((row) => row.workforceNumber?.trim().toLowerCase())
      .filter((value): value is string => Boolean(value)),
  );
  const results: BulkPreviewRow[] = [];
  const allocatedWorkforceNumbers: string[] = [];

  type PendingCreate = {
    index: number;
    validated: BulkPreviewRow;
    company: Company;
    companyName: string;
    extraMessages: string[];
  };
  type PendingUpdate = {
    index: number;
    validated: BulkPreviewRow;
    company: Company;
    companyName: string;
    entityId: string;
    extraMessages: string[];
  };

  const pendingCreates: PendingCreate[] = [];
  const pendingUpdates: PendingUpdate[] = [];
  // Imported workforce records (created + updated). Drives the two post-import
  // passes: matrix row sync (Phase 3c) and document folders (Phase 3d).
  const importedTargets: Array<{ index: number; wf: AdminWorkforceRecord }> =
    [];

  // Phase 1: validate + ensure companies (serial — unique company names).
  const phase1 = log.phase("phase1:validate+companies");
  for (const row of input.rows) {
    const fields = mapCandidateFields(row.fields);
    for (const [key, value] of Object.entries(row.fields)) {
      if (value != null && value !== "") {
        fields[key] = value;
      }
    }
    fields.dateOfBirth = normalizeDateValue(fields.dateOfBirth);
    fields.cscsExpiry = normalizeDateValue(fields.cscsExpiry);
    fields.swqrExpiry = normalizeDateValue(fields.swqrExpiry);
    fields.eusrExpiry = normalizeDateValue(fields.eusrExpiry);

    if (!fields.workforceNumber?.trim()) {
      fields.workforceNumber = null;
    }

    const validated = validateCandidateRow(
      row.rowNumber,
      fields,
      companies,
      liveWorkforce,
      people,
      allocatedWorkforceNumbers,
      { autoCreateMissing },
    );

    if (validated.status === "Error") {
      log.debug("row rejected", {
        row: validated.rowNumber,
        reason: validated.messages[0] ?? "validation error",
      });
      results.push(validated);
      continue;
    }

    const assignedNumber = validated.fields.workforceNumber?.trim();
    if (assignedNumber && !allocatedWorkforceNumbers.includes(assignedNumber)) {
      allocatedWorkforceNumbers.push(assignedNumber);
    }

    try {
      const ensured = await resolveOrCreateCompany(
        companies,
        validated.resolvedCompanyName ?? validated.fields.company ?? "",
        validated.fields.companyNumber,
        { autoCreateMissing },
      );
      companies = ensured.companies;
      const companyName = ensured.company.companyName;
      const resultIndex = results.length;
      // Placeholder so result order matches input order for non-error rows.
      results.push({ ...validated, status: validated.status });

      if (
        validated.status === "Warning" &&
        validated.duplicateMatch === "nameCompany"
      ) {
        if (input.duplicateMode === "update" && validated.matchedEntityId) {
          pendingUpdates.push({
            index: resultIndex,
            validated,
            company: ensured.company,
            companyName,
            entityId: validated.matchedEntityId,
            extraMessages: [
              "Updated existing candidate (name + company match).",
            ],
          });
          continue;
        }
        pendingCreates.push({
          index: resultIndex,
          validated,
          company: ensured.company,
          companyName,
          extraMessages: [
            "Created new candidate (soft name match was a warning only).",
          ],
        });
        continue;
      }

      if (validated.status === "Duplicate") {
        if (input.duplicateMode === "skip") {
          results[resultIndex] = {
            ...validated,
            status: "Skipped",
            messages: [...validated.messages, "Skipped duplicate (default)."],
          };
          continue;
        }
        if (input.duplicateMode === "update" && validated.matchedEntityId) {
          pendingUpdates.push({
            index: resultIndex,
            validated,
            company: ensured.company,
            companyName,
            entityId: validated.matchedEntityId,
            extraMessages: ["Updated existing candidate."],
          });
          continue;
        }
        if (input.duplicateMode === "create") {
          pendingCreates.push({
            index: resultIndex,
            validated,
            company: ensured.company,
            companyName,
            extraMessages: [
              "Created new candidate (admin confirmed create despite duplicate).",
            ],
          });
          continue;
        }
        results[resultIndex] = {
          ...validated,
          status: "Skipped",
          messages: [...validated.messages, "Skipped duplicate."],
        };
        continue;
      }

      pendingCreates.push({
        index: resultIndex,
        validated,
        company: ensured.company,
        companyName,
        extraMessages: ["Imported successfully."],
      });
    } catch (error) {
      log.warn("company resolve failed", {
        row: validated.rowNumber,
        company: validated.fields.company ?? null,
        error: error instanceof Error ? error.message : String(error),
      });
      results.push({
        ...validated,
        status: "Error",
        messages: [
          ...validated.messages,
          error instanceof Error
            ? error.message
            : "Failed to import candidate.",
        ],
      });
    }
  }

  phase1.end({
    rows: input.rows.length,
    creates: pendingCreates.length,
    updates: pendingUpdates.length,
    errors: results.filter((r) => r.status === "Error").length,
    companies: companies.length,
  });

  await withBulkSharePointWrites(async () => {
    // Phase 2: pre-create unique TMs / supervisors / departments (serial).
    const phase2 = log.phase("phase2:permissions+departments");
    const { createAdminDepartment } = await import(
      "@/lib/services/departmentService"
    );
    const deptKey = (value: string) =>
      value.trim().toLowerCase().replace(/\s+/g, " ");

    for (const job of [...pendingCreates, ...pendingUpdates]) {
      const tm = job.validated.fields.trainingManager?.trim();
      if (tm && !findPermissionPerson(people, tm)) {
        const ensured = await ensurePermissionPerson({
          displayName: tm,
          roleType: "Admin",
          companyId: job.company.id,
          people,
        });
        people = ensured.people;
      }
      const supervisor = job.validated.fields.supervisor?.trim();
      if (supervisor && !findPermissionPerson(people, supervisor)) {
        const ensured = await ensurePermissionPerson({
          displayName: supervisor,
          roleType: "Customer",
          companyId: job.company.id,
          people,
        });
        people = ensured.people;
      }
      const department = job.validated.fields.department?.trim();
      if (department) {
        const hit = departmentRecords.find(
          (row) =>
            deptKey(row.name) === deptKey(department) &&
            (!row.companyId || row.companyId === job.company.id),
        );
        if (!hit) {
          const created = await createAdminDepartment({
            name: department,
            companyId: job.company.id,
            companyName: job.companyName,
            skipDuplicateScan: true,
          });
          departmentRecords = [
            ...departmentRecords,
            {
              id: created.id,
              name: created.name,
              companyId: created.companyId,
              companyName: created.companyName,
            },
          ];
        }
      }
    }

    phase2.end({
      people: people.length,
      departments: departmentRecords.length,
    });

    // Phase 3a: updates stay sequential (rarer path).
    const phase3a = log.phase("phase3a:updates");
    for (const job of pendingUpdates) {
      try {
        const updated = await updateAdminWorkforce(job.entityId, {
          ...buildNonBlankUpdate(job.validated.fields, job.companyName),
          companyId: job.company.id,
          createMissingPermissionPeople: true,
          permissionPeople: people,
          departmentRecords,
          // Bulk syncs the matrix once in Phase 3c — skip per-row sync here.
          bulkMode: true,
        });
        const idx = liveWorkforce.findIndex((w) => w.id === updated.id);
        if (idx >= 0) liveWorkforce[idx] = updated;
        importedTargets.push({ index: job.index, wf: updated });
        results[job.index] = {
          ...job.validated,
          status: "Imported",
          resolvedCompanyName: job.companyName,
          messages: [...job.validated.messages, ...job.extraMessages],
        };
        log.debug("row updated", {
          row: job.validated.rowNumber,
          wf: updated.workforceNumber ?? null,
          name: updated.candidateName,
        });
      } catch (error) {
        log.warn("update failed", {
          row: job.validated.rowNumber,
          name: job.validated.fields.candidateName ?? null,
          error: error instanceof Error ? error.message : String(error),
        });
        results[job.index] = {
          ...job.validated,
          status: "Error",
          messages: [
            ...job.validated.messages,
            error instanceof Error
              ? error.message
              : "Failed to import candidate.",
          ],
        };
      }
    }
    phase3a.end({ updates: pendingUpdates.length });

    // Phase 3b: parallel workforce creates (main speed win).
    const phase3b = log.phase("phase3b:creates");
    let createdCount = 0;
    const knownSnapshot = Array.from(knownWorkforceNumbers);
    await mapPool(pendingCreates, BULK_CREATE_CONCURRENCY, async (job) => {
      try {
        const created = await createAdminWorkforce({
          ...workforceWritePayload(job.validated.fields, job.companyName),
          companyId: job.company.id,
          companyNumber:
            job.company.companyNumber ??
            job.validated.fields.companyNumber ??
            null,
          createMissingPermissionPeople: true,
          permissionPeople: people,
          departmentRecords,
          knownWorkforceNumbers: knownSnapshot,
          bulkMode: true,
        });
        liveWorkforce.push(created);
        const wfNum = created.workforceNumber?.trim().toLowerCase();
        if (wfNum) knownWorkforceNumbers.add(wfNum);
        importedTargets.push({ index: job.index, wf: created });
        results[job.index] = {
          ...job.validated,
          status: "Imported",
          matchedEntityId: created.id,
          matchedEntityName: created.candidateName,
          resolvedCompanyName: job.companyName,
          messages: [...job.validated.messages, ...job.extraMessages],
        };
        createdCount += 1;
        log.debug("row created", {
          row: job.validated.rowNumber,
          wf: created.workforceNumber ?? null,
          name: created.candidateName,
        });
        if (createdCount % 10 === 0) {
          log.info("progress", {
            done: createdCount,
            total: pendingCreates.length,
          });
        }
      } catch (error) {
        log.warn("create failed", {
          row: job.validated.rowNumber,
          name: job.validated.fields.candidateName ?? null,
          error: error instanceof Error ? error.message : String(error),
        });
        results[job.index] = {
          ...job.validated,
          status: "Error",
          messages: [
            ...job.validated.messages,
            error instanceof Error
              ? error.message
              : "Failed to import candidate.",
          ],
        };
      }
    });
    phase3b.end({
      created: createdCount,
      errors: pendingCreates.length - createdCount,
      concurrency: BULK_CREATE_CONCURRENCY,
    });

    // Phase 3c: create/update REAL Training Matrix Update rows for every
    // imported candidate, so the matrix no longer relies on synthetic
    // `workforce-only:` ids and later matrix spreadsheet imports can PATCH
    // real ids. One matrix read up front; sequential writes keep the dedupe
    // cache consistent and never wipe existing expiry data with blanks.
    if (importedTargets.length) {
      const phase3c = log.phase("phase3c:trainingMatrix");
      const { listTrainingMatrixExampleRows, syncWorkforceToTrainingMatrix } =
        await import(
          "@/lib/services/bulkUpload/trainingMatrixExampleService"
        );
      const { candidateNameKey } = await import(
        "@/lib/services/bulkUpload/workforceMatrixSync"
      );
      const matrixRows = await listTrainingMatrixExampleRows();
      log.info("matrix read", { rows: matrixRows.length });

      // Group by the dedupe key (candidate name). Distinct names never share a
      // matrix row, so groups run in parallel; same-name candidates stay
      // sequential inside a group so the first creates the row and the rest
      // update it — identical to the old sequential result, still no duplicates.
      const groups = new Map<
        string,
        Array<{ index: number; wf: AdminWorkforceRecord }>
      >();
      for (const target of importedTargets) {
        const key = candidateNameKey(target.wf.candidateName);
        const bucket = groups.get(key);
        if (bucket) bucket.push(target);
        else groups.set(key, [target]);
      }

      let synced = 0;
      await mapPool(
        [...groups.entries()],
        BULK_CREATE_CONCURRENCY,
        async ([groupKey, group]) => {
          // Seed a local rows view from the up-front snapshot for this name, so
          // a create by the first same-name candidate is visible to the rest.
          const localRows = matrixRows.filter(
            (row) => candidateNameKey(row.candidateName) === groupKey,
          );
          for (const { index, wf } of group) {
            const current = results[index];
            if (!current) continue;
            try {
              // Thread the strong link keys (WorkforceItemId, CompanyItemId) so
              // bulk-created rows link by id like the manual create/update path.
              const sync = await syncWorkforceToTrainingMatrix(
                toMatrixProfile(wf),
                { existingRows: localRows },
              );
              if (sync.created) localRows.push(sync.row);
              results[index] = {
                ...current,
                messages: [
                  ...current.messages,
                  sync.created
                    ? `Training Matrix row created (#${sync.id}).`
                    : `Training Matrix row updated (#${sync.id}).`,
                ],
              };
              synced += 1;
              log.debug("matrix synced", {
                row: current.rowNumber,
                id: sync.id,
                created: sync.created,
              });
              if (synced % 10 === 0) {
                log.info("progress", {
                  done: synced,
                  total: importedTargets.length,
                });
              }
            } catch (error) {
              log.warn("matrix sync failed", {
                row: current.rowNumber,
                error:
                  error instanceof Error ? error.message : String(error),
              });
              results[index] = {
                ...current,
                messages: [
                  ...current.messages,
                  `Candidate imported, but Training Matrix row sync failed: ${
                    error instanceof Error ? error.message : "unknown error"
                  }.`,
                ],
              };
            }
          }
        },
      );
      phase3c.end({ synced, targets: importedTargets.length });
    }

    // Phase 3d: ensure the Customer Documents folder structure for every
    // imported candidate — company folder + Company Documents + Candidates +
    // {Workforce Number - Candidate Name} + Certificates / Card Scans /
    // NVQ Documents / Other Documents. Reuses ensureCandidateDocumentFolders,
    // which is idempotent (resolve-first, number-prefix match) and never
    // throws — folder issues are reported per row, never blocking the import.
    // Bounded concurrency keeps 50 candidates well under the route timeout.
    if (importedTargets.length) {
      const phase3d = log.phase("phase3d:documentFolders");
      let foldersDone = 0;
      const { ensureCandidateDocumentFolders, createFolderEnsureCache } =
        await import("@/lib/services/customerDocumentsFolderService");
      // One cache for the whole batch: the drive id and each company's
      // Company Documents / Candidates folders resolve once and are shared
      // (as in-flight promises) across all candidates instead of per-row.
      const folderCache = createFolderEnsureCache();
      await mapPool(
        importedTargets,
        BULK_CREATE_CONCURRENCY,
        async ({ index, wf }) => {
          const current = results[index];
          if (!current) return;
          try {
            const folders = await ensureCandidateDocumentFolders(
              {
                companyName: wf.companyName,
                companyNumber: wf.companyNumber ?? null,
                candidateName: wf.candidateName,
                workforceNumber: wf.workforceNumber ?? null,
              },
              folderCache,
            );
            results[index] = {
              ...current,
              messages: [
                ...current.messages,
                folders.ok
                  ? "Document folders created or already existed under the company folder."
                  : `Folder creation pending/failed: ${folders.warning}`,
              ],
            };
            if (!folders.ok) {
              log.warn("folders warning", {
                row: current.rowNumber,
                warning: folders.warning,
              });
            } else {
              log.debug("folders ensured", {
                row: current.rowNumber,
                name: wf.candidateName,
              });
            }
          } catch (error) {
            log.warn("folders failed", {
              row: current.rowNumber,
              error: error instanceof Error ? error.message : String(error),
            });
            results[index] = {
              ...current,
              messages: [
                ...current.messages,
                `Candidate imported, but document folders were not created: ${
                  error instanceof Error ? error.message : "unknown error"
                }.`,
              ],
            };
          } finally {
            foldersDone += 1;
            if (foldersDone % 10 === 0) {
              log.info("progress", {
                done: foldersDone,
                total: importedTargets.length,
              });
            }
          }
        },
      );
      phase3d.end({ processed: foldersDone, targets: importedTargets.length });
    }
  });

  return results;
}

export function buildValidationReportCsv(rows: BulkPreviewRow[]): string {
  const header = [
    "Row",
    "Status",
    "Candidate Name",
    "Company",
    "Workforce Number",
    "DOB",
    "Department",
    "Matched Company Id",
    "Matched Company Number",
    "Matched Company Name",
    "Company Matched By",
    "Matched Candidate Id",
    "Warning/Error",
    "Messages",
  ];
  const lines = [header.map(escapeCsv).join(",")];
  for (const row of rows) {
    const isProblem = row.status === "Error" || row.status === "Warning";
    lines.push(
      [
        String(row.rowNumber),
        row.status,
        row.fields.candidateName ?? "",
        row.fields.company ?? "",
        row.fields.workforceNumber ?? "",
        row.fields.dateOfBirth ?? "",
        row.fields.department ?? "",
        row.matchedCompanyId ?? "",
        row.matchedCompanyNumber ?? "",
        row.matchedCompanyName ?? "",
        row.companyMatchedBy ?? "",
        row.matchedEntityId ?? "",
        isProblem ? row.messages.join("; ") : "",
        row.messages.join("; "),
      ]
        .map(escapeCsv)
        .join(","),
    );
  }
  return `${lines.join("\r\n")}\r\n`;
}

function escapeCsv(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}
