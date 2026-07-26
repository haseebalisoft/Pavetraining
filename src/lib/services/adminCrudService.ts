import "server-only";

import { NotFoundError } from "@/lib/services/errorHandler";
import { ValidationError } from "@/lib/services/validationService";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asLookupOrString,
  asNullableString,
  asString,
  createListItemByKey,
  extractLookupId,
  getListItemByKey,
  getListItemsByKey,
  listHasColumn,
  toSharePointFields,
  updateListItemFieldsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import {
  normalizePermissionRoleType,
  toSharePointRoleType,
} from "@/lib/services/permissionService";
import { mapCompanyFields, getCompanyById } from "@/lib/services/companyService";
import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import type { Company, RoleType } from "@/types/models";
import type {
  AdminDocumentRecord,
  DocumentMetadataStatus,
} from "@/types/adminDocuments";

export type {
  AdminDocumentRecord,
  DocumentMetadataStatus,
  CustomerDocumentType,
} from "@/types/adminDocuments";
export { CUSTOMER_DOCUMENT_TYPES } from "@/types/adminDocuments";

export { mapCompanyFields as mapAdminCompany } from "@/lib/services/companyService";

function requireText(value: unknown, label: string): string {
  const text = asString(value);
  if (!text) {
    throw new ValidationError(`${label} is required.`);
  }
  return text;
}

function optionalText(value: unknown): string | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  return asString(value) ?? String(value);
}

function optionalBool(value: unknown): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (["true", "yes", "1"].includes(normalized)) return true;
    if (["false", "no", "0"].includes(normalized)) return false;
  }
  throw new ValidationError("Boolean fields must be true/false or Yes/No.");
}

function asDateInput(value: unknown): string | null | undefined {
  if (value === undefined) return undefined;
  if (value === null || value === "") return null;
  const text = String(value).trim();
  return text || null;
}

function matchesCompany(
  value: string | null | undefined,
  companyName: string | null | undefined,
): boolean {
  if (!companyName) return true;
  return (value ?? "").trim().toLowerCase() === companyName.trim().toLowerCase();
}

/* ───────────────── Companies ───────────────── */

function optionalEmail(value: unknown, label: string): string | null {
  const text = optionalText(value);
  if (!text) return null;
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) {
    throw new ValidationError(`${label} must be a valid email address.`);
  }
  return text;
}

function companyWritePayload(input: Record<string, unknown>, partial: boolean) {
  const companyName = partial
    ? optionalText(input.companyName)
    : requireText(input.companyName, "Company name");
  const companyNumber = partial
    ? input.companyNumber === undefined
      ? undefined
      : optionalText(input.companyNumber)
    : requireText(input.companyNumber, "Company number");

  if (input.email !== undefined && input.email !== null && input.email !== "") {
    optionalEmail(input.email, "Email");
  }
  if (
    input.accountsEmail !== undefined &&
    input.accountsEmail !== null &&
    input.accountsEmail !== ""
  ) {
    optionalEmail(input.accountsEmail, "Accounts email");
  }

  return toSharePointFields("company", {
    title: companyName ?? undefined,
    companyName: companyName ?? undefined,
    companyNumber,
    companySize:
      input.companySize === undefined
        ? undefined
        : optionalText(input.companySize),
    registeredAddress:
      input.registeredAddress === undefined
        ? undefined
        : optionalText(input.registeredAddress),
    companyRegNumber:
      input.companyRegNumber === undefined
        ? undefined
        : optionalText(input.companyRegNumber),
    vatNo: input.vatNo === undefined ? undefined : optionalText(input.vatNo),
    telNo: input.telNo === undefined ? undefined : optionalText(input.telNo),
    email: input.email === undefined ? undefined : optionalText(input.email),
    mainContact:
      input.mainContact === undefined
        ? undefined
        : optionalText(input.mainContact),
    accountsContactName:
      input.accountsContactName === undefined
        ? undefined
        : optionalText(input.accountsContactName),
    accountsAddress:
      input.accountsAddress === undefined
        ? undefined
        : optionalText(input.accountsAddress),
    accountsContactNumber:
      input.accountsContactNumber === undefined
        ? undefined
        : optionalText(input.accountsContactNumber),
    accountsEmail:
      input.accountsEmail === undefined
        ? undefined
        : optionalText(input.accountsEmail),
    notesPricesAgreed:
      input.notesPricesAgreed === undefined
        ? undefined
        : optionalText(input.notesPricesAgreed),
    companyLogo:
      input.companyLogo === undefined
        ? undefined
        : optionalText(input.companyLogo),
    status: optionalText(input.status) ?? (partial ? undefined : "Active"),
  });
}

export async function listAdminCompanies(): Promise<Company[]> {
  const items = await getListItemsByKey("company", { top: 5000 });
  return items
    .map((item) => mapCompanyFields(item.id, item.fields))
    .filter((row): row is Company => row !== null)
    .sort((a, b) => a.companyName.localeCompare(b.companyName));
}

export async function createAdminCompany(input: Record<string, unknown>) {
  const payload = companyWritePayload(input, false);
  const item = await createListItemByKey("company", payload);
  const mapped = mapCompanyFields(item.id, item.fields);
  if (!mapped) throw new Error("Created company could not be mapped.");

  const { ensureCompanyDocumentFolders } = await import(
    "@/lib/services/customerDocumentsFolderService"
  );
  const folderResult = await ensureCompanyDocumentFolders({
    companyName: mapped.companyName,
    companyNumber: mapped.companyNumber,
  });

  return {
    ...mapped,
    folderWarning: folderResult.ok ? undefined : folderResult.warning,
  };
}

export async function updateAdminCompany(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("company", id);
  if (!existing) throw new NotFoundError("Company not found.");

  if (input.companyName !== undefined) {
    requireText(input.companyName, "Company name");
  }
  if (input.companyNumber !== undefined) {
    requireText(input.companyNumber, "Company number");
  }

  const payload = companyWritePayload(input, true);
  const item = await updateListItemFieldsByKey("company", id, payload);
  const mapped = mapCompanyFields(item.id, item.fields);
  if (!mapped) throw new Error("Updated company could not be mapped.");
  return mapped;
}

/* ───────────────── Workforce ───────────────── */

const workforceFields = getSharePointFields("workforce");

export interface AdminWorkforceRecord {
  id: string;
  candidateName: string;
  companyName: string;
  workforceNumber: string | null;
  dateOfBirth: string | null;
  department: string | null;
  status: string | null;
  trainingManager: string | null;
  supervisor: string | null;
}

function mapWorkforceDepartment(value: unknown): string | null {
  if (Array.isArray(value)) {
    const parts = value
      .map((entry) => {
        if (typeof entry === "string") return entry.trim();
        if (entry && typeof entry === "object") {
          const record = entry as { LookupValue?: unknown; Label?: unknown };
          return (
            asString(record.LookupValue) ?? asString(record.Label) ?? ""
          ).trim();
        }
        return "";
      })
      .filter(Boolean);
    return parts.length ? parts.join(", ") : null;
  }
  return asLookupOrString(value) ?? asNullableString(value);
}

function mapWorkforce(
  item: SharePointListItem,
  companyNameById?: Map<string, string>,
): AdminWorkforceRecord | null {
  // Candidate Name is text; Company Name is a Lookup (Graph often returns LookupId only).
  const candidateName =
    asLookupOrString(item.fields[workforceFields.candidateName]) ??
    asString(item.fields[workforceFields.candidateName]);
  const companyLookupId = extractLookupId(
    item.fields,
    workforceFields.companyName,
  );
  const companyName =
    asLookupOrString(item.fields[workforceFields.companyName]) ??
    asString(item.fields[workforceFields.companyName]) ??
    (companyLookupId && companyNameById
      ? (companyNameById.get(companyLookupId) ?? null)
      : null);
  if (!candidateName || !companyName) return null;
  return {
    id: item.id,
    candidateName,
    companyName,
    workforceNumber: asNullableString(
      item.fields[workforceFields.workforceNumber],
    ),
    dateOfBirth: asNullableString(item.fields[workforceFields.dateOfBirth]),
    department: mapWorkforceDepartment(item.fields[workforceFields.department]),
    status: asNullableString(item.fields[workforceFields.status]),
    trainingManager:
      asLookupOrString(item.fields[workforceFields.trainingManager]) ??
      asNullableString(item.fields[workforceFields.trainingManager]),
    supervisor:
      asLookupOrString(item.fields[workforceFields.supervisor]) ??
      asNullableString(item.fields[workforceFields.supervisor]),
  };
}

export async function listAdminWorkforce(companyName?: string | null) {
  const [items, companies] = await Promise.all([
    getListItemsByKey("workforce", { top: 5000 }),
    listAdminCompanies(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  return items
    .map((item) => mapWorkforce(item, companyNameById))
    .filter((row): row is AdminWorkforceRecord => {
      if (!row) return false;
      return matchesCompany(row.companyName, companyName);
    })
    .sort((a, b) => a.candidateName.localeCompare(b.candidateName));
}

export async function createAdminWorkforce(input: Record<string, unknown>) {
  const candidateName = requireText(input.candidateName, "Candidate name");
  const companyName = requireText(input.companyName, "Company");
  const workforceNumber = optionalText(input.workforceNumber);

  const companies = await listAdminCompanies();
  const company =
    companies.find(
      (row) =>
        row.companyName.trim().toLowerCase() ===
        companyName.trim().toLowerCase(),
    ) ??
    companies.find((row) => {
      const normalize = (value: string) =>
        value
          .trim()
          .toLowerCase()
          .replace(/\bltd\b\.?/g, "")
          .replace(/\blimited\b/g, "")
          .replace(/[^\w\s]/g, " ")
          .replace(/\s+/g, " ")
          .trim();
      return normalize(row.companyName) === normalize(companyName);
    }) ??
    null;
  if (!company) {
    throw new ValidationError(`Company "${companyName}" was not found.`);
  }

  // CompanyName is a Lookup — write LookupId, not free text.
  const payload = toSharePointFields("workforce", {
    candidateName,
    workforceNumber,
    dateOfBirth: asDateInput(input.dateOfBirth),
    department: optionalText(input.department),
    status: optionalText(input.status) ?? "Active",
  });
  payload.CompanyNameLookupId = Number(company.id);

  const item = await createListItemByKey("workforce", payload);
  const mapped =
    mapWorkforce(
      item,
      new Map([[company.id, company.companyName]]),
    ) ??
    ({
      id: item.id,
      candidateName,
      companyName: company.companyName,
      workforceNumber,
      dateOfBirth: asNullableString(payload[workforceFields.dateOfBirth]),
      department: optionalText(input.department),
      status: optionalText(input.status) ?? "Active",
      trainingManager: null,
      supervisor: null,
    } satisfies AdminWorkforceRecord);

  const { ensureCandidateDocumentFolders } = await import(
    "@/lib/services/customerDocumentsFolderService"
  );
  const folderResult = await ensureCandidateDocumentFolders({
    companyName: company.companyName,
    companyNumber: company.companyNumber ?? null,
    candidateName,
    workforceNumber,
  });

  return {
    ...mapped,
    folderWarning: folderResult.ok ? undefined : folderResult.warning,
  };
}

export async function updateAdminWorkforce(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("workforce", id);
  if (!existing) throw new NotFoundError("Candidate not found.");

  const payload = toSharePointFields("workforce", {
    candidateName: optionalText(input.candidateName) ?? undefined,
    workforceNumber:
      input.workforceNumber === undefined
        ? undefined
        : optionalText(input.workforceNumber),
    dateOfBirth:
      input.dateOfBirth === undefined
        ? undefined
        : asDateInput(input.dateOfBirth),
    department:
      input.department === undefined
        ? undefined
        : optionalText(input.department),
    status: optionalText(input.status) ?? undefined,
  });

  if (input.companyName !== undefined) {
    const companyName = optionalText(input.companyName);
    if (companyName) {
      const companies = await listAdminCompanies();
      const company = companies.find(
        (row) =>
          row.companyName.trim().toLowerCase() ===
          companyName.trim().toLowerCase(),
      );
      if (!company) {
        throw new ValidationError(`Company "${companyName}" was not found.`);
      }
      payload.CompanyNameLookupId = Number(company.id);
    }
  }

  const item = await updateListItemFieldsByKey("workforce", id, payload);
  const companies = await listAdminCompanies();
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const mapped = mapWorkforce(item, companyNameById);
  if (!mapped) throw new Error("Updated candidate could not be mapped.");
  return mapped;
}

/* ───────────────── Training Matrix ───────────────── */

const matrixFields = getSharePointFields("trainingMatrix");

export interface AdminMatrixRecord {
  id: string;
  candidateName: string;
  companyName: string | null;
  department: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  matrixNotes: string | null;
  nextExpiryDate: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
}

function mapMatrix(
  item: SharePointListItem,
  lookups?: {
    companyNameById?: Map<string, string>;
    workforceById?: Map<string, AdminWorkforceRecord>;
  },
): AdminMatrixRecord | null {
  // Candidate Name + company fields are Lookups — Graph often returns LookupId only.
  const candidateLookupId = extractLookupId(
    item.fields,
    matrixFields.candidateName,
  );
  const companyLookupId =
    extractLookupId(item.fields, matrixFields.matrixCompany) ??
    extractLookupId(item.fields, matrixFields.companyName);

  const workforceHit =
    candidateLookupId && lookups?.workforceById
      ? lookups.workforceById.get(candidateLookupId)
      : undefined;

  const candidateName =
    asLookupOrString(item.fields[matrixFields.candidateName]) ??
    asString(item.fields[matrixFields.candidateName]) ??
    workforceHit?.candidateName ??
    null;

  if (!candidateName) return null;

  const companyName =
    asLookupOrString(item.fields[matrixFields.matrixCompany]) ??
    asLookupOrString(item.fields[matrixFields.companyName]) ??
    (companyLookupId && lookups?.companyNameById
      ? (lookups.companyNameById.get(companyLookupId) ?? null)
      : null) ??
    workforceHit?.companyName ??
    null;

  return {
    id: item.id,
    candidateName,
    companyName,
    department: asNullableString(item.fields[matrixFields.department]),
    overallStatus: asNullableString(item.fields[matrixFields.overallStatus]),
    needsReview: asBoolean(item.fields[matrixFields.needsReview]),
    matrixNotes: asNullableString(item.fields[matrixFields.matrixNotes]),
    nextExpiryDate: asNullableString(item.fields[matrixFields.nextExpiryDate]),
    n001Expiry: asNullableString(item.fields[matrixFields.n001Expiry]),
    n003Expiry: asNullableString(item.fields[matrixFields.n003Expiry]),
    n004Expiry: asNullableString(item.fields[matrixFields.n004Expiry]),
    n010Expiry: asNullableString(item.fields[matrixFields.n010Expiry]),
    n020Expiry: asNullableString(item.fields[matrixFields.n020Expiry]),
    n021Expiry: asNullableString(item.fields[matrixFields.n021Expiry]),
    n027Expiry: asNullableString(item.fields[matrixFields.n027Expiry]),
    n100Expiry: asNullableString(item.fields[matrixFields.n100Expiry]),
  };
}

export async function listAdminMatrix(companyName?: string | null) {
  const [items, companies, workforce] = await Promise.all([
    getListItemsByKey("trainingMatrix", { top: 5000 }),
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const workforceById = new Map(
    workforce.map((row) => [row.id, row] as const),
  );
  return items
    .map((item) => mapMatrix(item, { companyNameById, workforceById }))
    .filter((row): row is AdminMatrixRecord => {
      if (!row) return false;
      return matchesCompany(row.companyName, companyName);
    });
}

export async function createAdminMatrix(input: Record<string, unknown>) {
  const candidateName = requireText(input.candidateName, "Candidate name");
  const companyName = requireText(input.companyName, "Company");

  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const company =
    companies.find(
      (row) =>
        row.companyName.trim().toLowerCase() ===
        companyName.trim().toLowerCase(),
    ) ?? null;
  if (!company) {
    throw new ValidationError(`Company "${companyName}" was not found.`);
  }

  const candidate =
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
          candidateName.trim().toLowerCase() &&
        row.companyName.trim().toLowerCase() ===
          company.companyName.trim().toLowerCase(),
    ) ??
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
        candidateName.trim().toLowerCase(),
    ) ??
    null;
  if (!candidate) {
    throw new ValidationError(
      `Candidate "${candidateName}" was not found in Workforce. Import or create the candidate first.`,
    );
  }

  // CandidateName + MatrixCompany are Lookups — write LookupIds only.
  const payload = toSharePointFields("trainingMatrix", {
    department: optionalText(input.department) ?? candidate.department,
    overallStatus: optionalText(input.overallStatus),
    needsReview: optionalBool(input.needsReview) ?? false,
    matrixNotes: optionalText(input.matrixNotes),
    nextExpiryDate: asDateInput(input.nextExpiryDate),
    n001Expiry: asDateInput(input.n001Expiry),
    n003Expiry: asDateInput(input.n003Expiry),
    n004Expiry: asDateInput(input.n004Expiry),
    n010Expiry: asDateInput(input.n010Expiry),
    n020Expiry: asDateInput(input.n020Expiry),
    n021Expiry: asDateInput(input.n021Expiry),
    n027Expiry: asDateInput(input.n027Expiry),
    n100Expiry: asDateInput(input.n100Expiry),
  });
  payload.CandidateNameLookupId = Number(candidate.id);
  payload.MatrixCompanyLookupId = Number(company.id);

  const item = await createListItemByKey("trainingMatrix", payload);
  const mapped =
    mapMatrix(item, {
      companyNameById: new Map([[company.id, company.companyName]]),
      workforceById: new Map([[candidate.id, candidate]]),
    }) ??
    ({
      id: item.id,
      candidateName: candidate.candidateName,
      companyName: company.companyName,
      department: optionalText(input.department) ?? candidate.department,
      overallStatus: optionalText(input.overallStatus),
      needsReview: optionalBool(input.needsReview) ?? false,
      matrixNotes: optionalText(input.matrixNotes),
      nextExpiryDate: asNullableString(
        asDateInput(input.nextExpiryDate) ?? null,
      ),
      n001Expiry: asNullableString(asDateInput(input.n001Expiry) ?? null),
      n003Expiry: asNullableString(asDateInput(input.n003Expiry) ?? null),
      n004Expiry: asNullableString(asDateInput(input.n004Expiry) ?? null),
      n010Expiry: asNullableString(asDateInput(input.n010Expiry) ?? null),
      n020Expiry: asNullableString(asDateInput(input.n020Expiry) ?? null),
      n021Expiry: asNullableString(asDateInput(input.n021Expiry) ?? null),
      n027Expiry: asNullableString(asDateInput(input.n027Expiry) ?? null),
      n100Expiry: asNullableString(asDateInput(input.n100Expiry) ?? null),
    } satisfies AdminMatrixRecord);
  return mapped;
}

export async function updateAdminMatrix(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("trainingMatrix", id);
  if (!existing) throw new NotFoundError("Matrix record not found.");

  const payload = toSharePointFields("trainingMatrix", {
    department:
      input.department === undefined
        ? undefined
        : optionalText(input.department),
    overallStatus:
      input.overallStatus === undefined
        ? undefined
        : optionalText(input.overallStatus),
    needsReview: optionalBool(input.needsReview),
    matrixNotes:
      input.matrixNotes === undefined
        ? undefined
        : optionalText(input.matrixNotes),
    nextExpiryDate:
      input.nextExpiryDate === undefined
        ? undefined
        : asDateInput(input.nextExpiryDate),
    n001Expiry:
      input.n001Expiry === undefined
        ? undefined
        : asDateInput(input.n001Expiry),
    n003Expiry:
      input.n003Expiry === undefined
        ? undefined
        : asDateInput(input.n003Expiry),
    n004Expiry:
      input.n004Expiry === undefined
        ? undefined
        : asDateInput(input.n004Expiry),
    n010Expiry:
      input.n010Expiry === undefined
        ? undefined
        : asDateInput(input.n010Expiry),
    n020Expiry:
      input.n020Expiry === undefined
        ? undefined
        : asDateInput(input.n020Expiry),
    n021Expiry:
      input.n021Expiry === undefined
        ? undefined
        : asDateInput(input.n021Expiry),
    n027Expiry:
      input.n027Expiry === undefined
        ? undefined
        : asDateInput(input.n027Expiry),
    n100Expiry:
      input.n100Expiry === undefined
        ? undefined
        : asDateInput(input.n100Expiry),
  });

  const wantsCandidate = input.candidateName !== undefined;
  const wantsCompany = input.companyName !== undefined;
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const workforceById = new Map(
    workforce.map((row) => [row.id, row] as const),
  );
  const mappedExisting = mapMatrix(existing, {
    companyNameById,
    workforceById,
  });

  if (wantsCandidate || wantsCompany) {
    const candidateName =
      optionalText(input.candidateName) ?? mappedExisting?.candidateName ?? "";
    const companyName =
      optionalText(input.companyName) ?? mappedExisting?.companyName ?? "";

    if (wantsCompany || !mappedExisting?.companyName) {
      const company = companies.find(
        (row) =>
          row.companyName.trim().toLowerCase() ===
          companyName.trim().toLowerCase(),
      );
      if (!company) {
        throw new ValidationError(
          `Company "${companyName || "(empty)"}" was not found.`,
        );
      }
      payload.MatrixCompanyLookupId = Number(company.id);
    }

    if (wantsCandidate || !mappedExisting?.candidateName) {
      const companyKey = (
        optionalText(input.companyName) ??
        mappedExisting?.companyName ??
        ""
      )
        .trim()
        .toLowerCase();
      const candidate =
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
              candidateName.trim().toLowerCase() &&
            row.companyName.trim().toLowerCase() === companyKey,
        ) ??
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
            candidateName.trim().toLowerCase(),
        );
      if (!candidate) {
        throw new ValidationError(
          `Candidate "${candidateName || "(empty)"}" was not found in Workforce.`,
        );
      }
      payload.CandidateNameLookupId = Number(candidate.id);
      if (!payload.MatrixCompanyLookupId) {
        const company = companies.find(
          (row) =>
            row.companyName.trim().toLowerCase() ===
            candidate.companyName.trim().toLowerCase(),
        );
        if (company) {
          payload.MatrixCompanyLookupId = Number(company.id);
        }
      }
    }
  }

  const item = await updateListItemFieldsByKey("trainingMatrix", id, payload);
  const mapped = mapMatrix(item, { companyNameById, workforceById });
  if (!mapped) throw new Error("Updated matrix row could not be mapped.");
  return mapped;
}

/* ───────────────── Training registers ───────────────── */

export type AdminRegisterKey =
  | "nporsRegister"
  | "eusrRegister"
  | "nrswaRegister"
  | "inHouseCertificates";

export interface AdminTrainingRecord {
  id: string;
  candidateName: string;
  companyName: string;
  trainingDate: string | null;
  trainingAddress: string | null;
  trainingOutcome: string | null;
  customerVisible: boolean;
  expiry: string | null;
  // register-specific
  nporsNumber?: string | null;
  noviceOrEwt?: string | null;
  nporsCategory?: string | null;
  eusrNumber?: string | null;
  eusrCategory?: string | null;
  swqrNumber?: string | null;
  course?: string | null;
  streetworksCategory?: string | null;
  certificateCategory?: string | null;
  courseCategory?: string | null;
}

type RegisterLookupMaps = {
  companyNameById: Map<string, string>;
  workforceNameById: Map<string, string>;
  workforceCompanyById: Map<string, string>;
};

async function loadRegisterLookupMaps(): Promise<RegisterLookupMaps> {
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const workforceNameById = new Map(
    workforce.map((row) => [row.id, row.candidateName] as const),
  );
  const workforceCompanyById = new Map(
    workforce.map((row) => [row.id, row.companyName] as const),
  );
  return { companyNameById, workforceNameById, workforceCompanyById };
}

function resolveRegisterPeople(
  fields: SharePointFields,
  candidateField: string,
  companyField: string,
  lookups: RegisterLookupMaps,
): { candidateName: string; companyName: string } | null {
  const candidateLookupId = extractLookupId(fields, candidateField);
  const companyLookupId = extractLookupId(fields, companyField);

  const candidateName =
    asLookupOrString(fields[candidateField]) ??
    asString(fields[candidateField]) ??
    (candidateLookupId
      ? (lookups.workforceNameById.get(candidateLookupId) ?? null)
      : null);

  if (!candidateName) return null;

  const companyName =
    asLookupOrString(fields[companyField]) ??
    asString(fields[companyField]) ??
    (companyLookupId
      ? (lookups.companyNameById.get(companyLookupId) ?? null)
      : null) ??
    (candidateLookupId
      ? (lookups.workforceCompanyById.get(candidateLookupId) ?? null)
      : null);

  if (!companyName) return null;
  return { candidateName, companyName };
}

function asMultiChoiceText(value: unknown): string | null {
  if (Array.isArray(value)) {
    const parts = value
      .map((entry) => {
        if (typeof entry === "string") return entry.trim();
        if (entry && typeof entry === "object") {
          const record = entry as { LookupValue?: unknown; Label?: unknown };
          return (
            asString(record.LookupValue) ?? asString(record.Label) ?? ""
          ).trim();
        }
        return String(entry ?? "").trim();
      })
      .filter(Boolean);
    return parts.length ? parts.join(", ") : null;
  }
  return asLookupOrString(value) ?? asNullableString(value);
}

function mapRegister(
  key: AdminRegisterKey,
  item: SharePointListItem,
  lookups: RegisterLookupMaps,
): AdminTrainingRecord | null {
  if (key === "nporsRegister") {
    const f = getSharePointFields("nporsRegister");
    const people = resolveRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
      lookups,
    );
    if (!people) return null;
    return {
      id: item.id,
      ...people,
      trainingDate: asNullableString(item.fields[f.trainingDate]),
      trainingAddress: asNullableString(item.fields[f.trainingAddress]),
      trainingOutcome: asNullableString(item.fields[f.trainingOutcome]),
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      expiry: asNullableString(item.fields[f.expiry]),
      nporsNumber:
        asNullableString(item.fields[f.nporsNumber]) ??
        asLookupOrString(item.fields.Candidate_x0020_Name_x003a__x0020) ??
        asNullableString(item.fields.On_x002f_Number),
      noviceOrEwt: asNullableString(item.fields[f.noviceOrEwt]),
      nporsCategory: asMultiChoiceText(item.fields[f.nporsCategory]),
    };
  }

  if (key === "eusrRegister") {
    const f = getSharePointFields("eusrRegister");
    const people = resolveRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
      lookups,
    );
    if (!people) return null;
    return {
      id: item.id,
      ...people,
      trainingDate: asNullableString(item.fields[f.trainingDate]),
      trainingAddress: asNullableString(item.fields[f.trainingAddress]),
      trainingOutcome: asNullableString(item.fields[f.trainingOutcome]),
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      expiry: asNullableString(item.fields[f.expiry]),
      eusrNumber: asNullableString(item.fields[f.eusrNumber]),
      eusrCategory: asNullableString(item.fields[f.eusrCategory]),
    };
  }

  if (key === "nrswaRegister") {
    const f = getSharePointFields("nrswaRegister");
    const people = resolveRegisterPeople(
      item.fields,
      f.candidateName,
      f.companyName,
      lookups,
    );
    if (!people) return null;
    return {
      id: item.id,
      ...people,
      trainingDate: asNullableString(item.fields[f.trainingDate]),
      trainingAddress: asNullableString(item.fields[f.trainingAddress]),
      trainingOutcome: asNullableString(item.fields[f.trainingOutcome]),
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      expiry: asNullableString(item.fields[f.expiryDate]),
      swqrNumber: asNullableString(item.fields[f.swqrNumber]),
      course: asNullableString(item.fields[f.course]),
      streetworksCategory: asNullableString(item.fields[f.streetworksCategory]),
    };
  }

  const f = getSharePointFields("inHouseCertificates");
  const people = resolveRegisterPeople(
    item.fields,
    f.candidateName,
    f.companyName,
    lookups,
  );
  if (!people) return null;
  return {
    id: item.id,
    ...people,
    trainingDate: asNullableString(item.fields[f.courseDate]),
    trainingAddress: asNullableString(item.fields[f.trainingAddress]),
    trainingOutcome: asNullableString(item.fields[f.trainingOutcome]),
    customerVisible: asBoolean(item.fields[f.customerVisible]),
    expiry: asNullableString(item.fields[f.expiryDate]),
    course: asNullableString(item.fields[f.courseCategory]),
    certificateCategory: asNullableString(item.fields[f.certificateCategory]),
    courseCategory: asNullableString(item.fields[f.courseCategory]),
  };
}

function registerWritePayload(
  key: AdminRegisterKey,
  input: Record<string, unknown>,
  mode: "create" | "update",
): SharePointFields {
  const outcome = optionalText(input.trainingOutcome);
  if (
    outcome &&
    !["Pass", "Fail", "pass", "fail"].includes(outcome)
  ) {
    throw new ValidationError("Training outcome must be Pass or Fail.");
  }
  const normalizedOutcome = outcome
    ? outcome.toLowerCase().startsWith("p")
      ? "Pass"
      : "Fail"
    : mode === "create"
      ? null
      : undefined;

  // CandidateName/CompanyName are Lookups — set LookupIds separately.
  if (key === "nporsRegister") {
    const values: Record<string, unknown> = {
      nporsNumber:
        input.nporsNumber === undefined
          ? undefined
          : optionalText(input.nporsNumber),
      trainingDate:
        input.trainingDate === undefined
          ? undefined
          : asDateInput(input.trainingDate),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      noviceOrEwt:
        input.noviceOrEwt === undefined
          ? undefined
          : optionalText(input.noviceOrEwt),
      nporsCategory:
        input.nporsCategory === undefined
          ? undefined
          : optionalText(input.nporsCategory),
      trainingOutcome: normalizedOutcome,
      expiry:
        input.expiry === undefined ? undefined : asDateInput(input.expiry),
      customerVisible: optionalBool(input.customerVisible),
    };
    if (mode === "create" && values.customerVisible === undefined) {
      values.customerVisible = true;
    }
    return toSharePointFields(key, values);
  }

  if (key === "eusrRegister") {
    const values: Record<string, unknown> = {
      eusrNumber:
        input.eusrNumber === undefined
          ? undefined
          : optionalText(input.eusrNumber),
      eusrCategory:
        input.eusrCategory === undefined
          ? undefined
          : optionalText(input.eusrCategory),
      trainingDate:
        input.trainingDate === undefined
          ? undefined
          : asDateInput(input.trainingDate),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      trainingOutcome: normalizedOutcome,
      expiry:
        input.expiry === undefined ? undefined : asDateInput(input.expiry),
      customerVisible: optionalBool(input.customerVisible),
    };
    if (mode === "create" && values.customerVisible === undefined) {
      values.customerVisible = true;
    }
    return toSharePointFields(key, values);
  }

  if (key === "nrswaRegister") {
    const values: Record<string, unknown> = {
      swqrNumber:
        input.swqrNumber === undefined
          ? undefined
          : optionalText(input.swqrNumber),
      course:
        input.course === undefined ? undefined : optionalText(input.course),
      streetworksCategory:
        input.streetworksCategory === undefined
          ? undefined
          : optionalText(input.streetworksCategory),
      trainingDate:
        input.trainingDate === undefined
          ? undefined
          : asDateInput(input.trainingDate),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      trainingOutcome: normalizedOutcome,
      expiryDate:
        input.expiry === undefined ? undefined : asDateInput(input.expiry),
      customerVisible: optionalBool(input.customerVisible),
    };
    if (mode === "create" && values.customerVisible === undefined) {
      values.customerVisible = true;
    }
    return toSharePointFields(key, values);
  }

  const values: Record<string, unknown> = {
    courseCategory:
      input.course === undefined && input.courseCategory === undefined
        ? undefined
        : optionalText(input.course ?? input.courseCategory),
    certificateCategory:
      input.certificateCategory === undefined
        ? undefined
        : optionalText(input.certificateCategory),
    courseDate:
      input.trainingDate === undefined
        ? undefined
        : asDateInput(input.trainingDate),
    trainingAddress:
      input.trainingAddress === undefined
        ? undefined
        : optionalText(input.trainingAddress),
    trainingOutcome: normalizedOutcome,
    expiryDate:
      input.expiry === undefined ? undefined : asDateInput(input.expiry),
    customerVisible: optionalBool(input.customerVisible),
  };
  if (mode === "create" && values.customerVisible === undefined) {
    values.customerVisible = true;
  }
  return toSharePointFields(key, values);
}

async function applyRegisterLookupIds(
  payload: SharePointFields,
  input: Record<string, unknown>,
  mode: "create" | "update",
): Promise<void> {
  const candidateName =
    mode === "create"
      ? requireText(input.candidateName, "Candidate name")
      : optionalText(input.candidateName);
  const companyName =
    mode === "create"
      ? requireText(input.companyName, "Company")
      : optionalText(input.companyName);

  if (!candidateName && !companyName) return;

  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);

  let resolvedCompany =
    companyName
      ? companies.find(
          (row) =>
            row.companyName.trim().toLowerCase() ===
            companyName.trim().toLowerCase(),
        )
      : null;

  let resolvedCandidate =
    candidateName
      ? workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
              candidateName.trim().toLowerCase() &&
            (!resolvedCompany ||
              row.companyName.trim().toLowerCase() ===
                resolvedCompany.companyName.trim().toLowerCase()),
        ) ??
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
            candidateName.trim().toLowerCase(),
        )
      : null;

  if (mode === "create" && !resolvedCandidate) {
    throw new ValidationError(
      `Candidate "${candidateName}" was not found in Workforce. Create the candidate first.`,
    );
  }
  if (mode === "create" && !resolvedCompany) {
    resolvedCompany =
      companies.find(
        (row) =>
          row.companyName.trim().toLowerCase() ===
          (resolvedCandidate?.companyName ?? "").trim().toLowerCase(),
      ) ?? null;
  }
  if (mode === "create" && !resolvedCompany) {
    throw new ValidationError(`Company "${companyName}" was not found.`);
  }

  if (resolvedCandidate) {
    payload.CandidateNameLookupId = Number(resolvedCandidate.id);
  }
  if (resolvedCompany) {
    payload.CompanyNameLookupId = Number(resolvedCompany.id);
  } else if (resolvedCandidate) {
    const fromWorkforce = companies.find(
      (row) =>
        row.companyName.trim().toLowerCase() ===
        resolvedCandidate!.companyName.trim().toLowerCase(),
    );
    if (fromWorkforce) {
      payload.CompanyNameLookupId = Number(fromWorkforce.id);
    }
  }
}

export async function listAdminRegister(
  key: AdminRegisterKey,
  companyName?: string | null,
) {
  const [items, lookups] = await Promise.all([
    getListItemsByKey(key, { top: 5000 }),
    loadRegisterLookupMaps(),
  ]);
  return items
    .map((item) => mapRegister(key, item, lookups))
    .filter((row): row is AdminTrainingRecord => {
      if (!row) return false;
      return matchesCompany(row.companyName, companyName);
    });
}

export async function createAdminRegister(
  key: AdminRegisterKey,
  input: Record<string, unknown>,
) {
  const payload = registerWritePayload(key, input, "create");
  await applyRegisterLookupIds(payload, input, "create");
  const item = await createListItemByKey(key, payload);
  const lookups = await loadRegisterLookupMaps();
  const mapped = mapRegister(key, item, lookups);
  if (!mapped) throw new Error("Created training record could not be mapped.");
  return mapped;
}

export async function updateAdminRegister(
  key: AdminRegisterKey,
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey(key, id);
  if (!existing) throw new NotFoundError("Training record not found.");
  const payload = registerWritePayload(key, input, "update");
  await applyRegisterLookupIds(payload, input, "update");
  const item = await updateListItemFieldsByKey(key, id, payload);
  const lookups = await loadRegisterLookupMaps();
  const mapped = mapRegister(key, item, lookups);
  if (!mapped) throw new Error("Updated training record could not be mapped.");
  return mapped;
}

/* ───────────────── NVQ ───────────────── */

const nvqFields = getSharePointFields("nvqRegister");

export interface AdminNvqRecord {
  id: string;
  candidateName: string;
  companyName: string | null;
  nvqTitle: string | null;
  boltOn: string | null;
  dateRegistered: string | null;
  inductionDate: string | null;
  stageOfNvq: string | null;
  notes: string | null;
  completedDate: string | null;
  customerVisible: boolean;
  status: "Active" | "Completed";
}

function mapNvq(
  item: SharePointListItem,
  lookups: RegisterLookupMaps,
): AdminNvqRecord | null {
  const candidateLookupId = extractLookupId(
    item.fields,
    nvqFields.candidateName,
  );
  const candidateName =
    asLookupOrString(item.fields[nvqFields.candidateName]) ??
    asString(item.fields[nvqFields.candidateName]) ??
    (candidateLookupId
      ? (lookups.workforceNameById.get(candidateLookupId) ?? null)
      : null);
  if (!candidateName) return null;

  const companyLookupId =
    extractLookupId(item.fields, nvqFields.nvqCompany) ??
    extractLookupId(item.fields, nvqFields.companyName);
  const companyName =
    asLookupOrString(item.fields[nvqFields.nvqCompany]) ??
    asLookupOrString(item.fields[nvqFields.companyName]) ??
    (companyLookupId
      ? (lookups.companyNameById.get(companyLookupId) ?? null)
      : null) ??
    (candidateLookupId
      ? (lookups.workforceCompanyById.get(candidateLookupId) ?? null)
      : null);

  const completedDate = asNullableString(item.fields[nvqFields.completedDate]);
  return {
    id: item.id,
    candidateName,
    companyName,
    nvqTitle: asNullableString(item.fields[nvqFields.nvqTitle]),
    boltOn: asNullableString(item.fields[nvqFields.boltonNvq]),
    dateRegistered: asNullableString(item.fields[nvqFields.dateRegistered]),
    inductionDate: asNullableString(item.fields[nvqFields.dateInductionBooked]),
    stageOfNvq: asNullableString(item.fields[nvqFields.stageOfNvq]),
    notes: asNullableString(item.fields[nvqFields.customerUpdateNotes]),
    completedDate,
    customerVisible: asBoolean(item.fields[nvqFields.customerVisible]),
    status: completedDate?.trim() ? "Completed" : "Active",
  };
}

export async function listAdminNvq(companyName?: string | null) {
  const [items, lookups] = await Promise.all([
    getListItemsByKey("nvqRegister", { top: 5000 }),
    loadRegisterLookupMaps(),
  ]);
  return items
    .map((item) => mapNvq(item, lookups))
    .filter((row): row is AdminNvqRecord => {
      if (!row) return false;
      return matchesCompany(row.companyName, companyName);
    });
}

export async function createAdminNvq(input: Record<string, unknown>) {
  const candidateName = requireText(input.candidateName, "Candidate name");
  const companyName = requireText(input.companyName, "Company");
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const company =
    companies.find(
      (row) =>
        row.companyName.trim().toLowerCase() ===
        companyName.trim().toLowerCase(),
    ) ?? null;
  if (!company) {
    throw new ValidationError(`Company "${companyName}" was not found.`);
  }
  const candidate =
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
          candidateName.trim().toLowerCase() &&
        row.companyName.trim().toLowerCase() ===
          company.companyName.trim().toLowerCase(),
    ) ??
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
        candidateName.trim().toLowerCase(),
    );
  if (!candidate) {
    throw new ValidationError(
      `Candidate "${candidateName}" was not found in Workforce. Create the candidate first.`,
    );
  }

  const payload = toSharePointFields("nvqRegister", {
    nvqTitle: optionalText(input.nvqTitle),
    boltonNvq: optionalText(input.boltOn),
    dateRegistered: asDateInput(input.dateRegistered),
    dateInductionBooked: asDateInput(input.inductionDate),
    stageOfNvq: optionalText(input.stageOfNvq),
    customerUpdateNotes: optionalText(input.notes),
    completedDate: asDateInput(input.completedDate),
    customerVisible: optionalBool(input.customerVisible) ?? true,
  });
  payload.CandidateNameLookupId = Number(candidate.id);
  payload.NVQCompanyLookupId = Number(company.id);

  const item = await createListItemByKey("nvqRegister", payload);
  const lookups = await loadRegisterLookupMaps();
  const mapped = mapNvq(item, lookups);
  if (!mapped) throw new Error("Created NVQ could not be mapped.");
  return mapped;
}

export async function updateAdminNvq(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("nvqRegister", id);
  if (!existing) throw new NotFoundError("NVQ record not found.");
  const payload = toSharePointFields("nvqRegister", {
    nvqTitle:
      input.nvqTitle === undefined ? undefined : optionalText(input.nvqTitle),
    boltonNvq:
      input.boltOn === undefined ? undefined : optionalText(input.boltOn),
    dateRegistered:
      input.dateRegistered === undefined
        ? undefined
        : asDateInput(input.dateRegistered),
    dateInductionBooked:
      input.inductionDate === undefined
        ? undefined
        : asDateInput(input.inductionDate),
    stageOfNvq:
      input.stageOfNvq === undefined
        ? undefined
        : optionalText(input.stageOfNvq),
    customerUpdateNotes:
      input.notes === undefined ? undefined : optionalText(input.notes),
    completedDate:
      input.completedDate === undefined
        ? undefined
        : asDateInput(input.completedDate),
    customerVisible: optionalBool(input.customerVisible),
  });

  if (input.candidateName !== undefined || input.companyName !== undefined) {
    const [companies, workforce] = await Promise.all([
      listAdminCompanies(),
      listAdminWorkforce(),
    ]);
    const lookups = {
      companyNameById: new Map(
        companies.map((row) => [row.id, row.companyName] as const),
      ),
      workforceNameById: new Map(
        workforce.map((row) => [row.id, row.candidateName] as const),
      ),
      workforceCompanyById: new Map(
        workforce.map((row) => [row.id, row.companyName] as const),
      ),
    };
    const existingMapped = mapNvq(existing, lookups);
    const candidateName =
      optionalText(input.candidateName) ?? existingMapped?.candidateName ?? "";
    const companyName =
      optionalText(input.companyName) ?? existingMapped?.companyName ?? "";

    if (companyName) {
      const company = companies.find(
        (row) =>
          row.companyName.trim().toLowerCase() ===
          companyName.trim().toLowerCase(),
      );
      if (!company) {
        throw new ValidationError(`Company "${companyName}" was not found.`);
      }
      payload.NVQCompanyLookupId = Number(company.id);
    }

    if (candidateName) {
      const companyKey = companyName.trim().toLowerCase();
      const candidate =
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
              candidateName.trim().toLowerCase() &&
            row.companyName.trim().toLowerCase() === companyKey,
        ) ??
        workforce.find(
          (row) =>
            row.candidateName.trim().toLowerCase() ===
            candidateName.trim().toLowerCase(),
        );
      if (!candidate) {
        throw new ValidationError(
          `Candidate "${candidateName}" was not found in Workforce.`,
        );
      }
      payload.CandidateNameLookupId = Number(candidate.id);
    }
  }

  const item = await updateListItemFieldsByKey("nvqRegister", id, payload);
  const lookups = await loadRegisterLookupMaps();
  const mapped = mapNvq(item, lookups);
  if (!mapped) throw new Error("Updated NVQ could not be mapped.");
  return mapped;
}

/* ───────────────── Documents ───────────────── */

const documentFields = getSharePointFields("customerDocuments");

export interface AdminDocumentListFilters {
  companyName?: string | null;
  candidate?: string | null;
  documentType?: string | null;
  /** true = visible only, false = hidden only, null/undefined = all */
  customerVisible?: boolean | null;
  /** Inclusive lower bound on modified date (ISO date or datetime). */
  modifiedFrom?: string | null;
  /** Inclusive upper bound on modified date (ISO date or datetime). */
  modifiedTo?: string | null;
}

function isSharePointFolder(fields: SharePointFields): boolean {
  const fs = fields[documentFields.fsObjType];
  return fs === 1 || fs === "1";
}

function decodePathSegment(segment: string): string {
  try {
    return decodeURIComponent(segment.replace(/\+/g, " "));
  } catch {
    return segment;
  }
}

/** Parts under the Customer Documents library from a FileRef / FileDirRef. */
export function customerDocumentsPathParts(
  pathValue: string | null | undefined,
): string[] {
  if (!pathValue?.trim()) return [];
  const parts = pathValue
    .split("/")
    .map((part) => decodePathSegment(part))
    .filter(Boolean);
  const libraryIdx = parts.findIndex(
    (part) => part.toLowerCase() === "customer documents",
  );
  return libraryIdx >= 0 ? parts.slice(libraryIdx + 1) : parts;
}

function resolveDocumentMetadataStatus(input: {
  company: string | null;
  documentType: string | null;
  customerVisible: boolean;
}): DocumentMetadataStatus {
  if (!input.company?.trim()) {
    return "Missing Company";
  }
  if (!input.documentType?.trim()) {
    return "Missing Document Type";
  }
  if (!input.customerVisible) {
    return "Hidden from Customer";
  }
  return "Complete";
}

export function mapDocument(item: SharePointListItem): AdminDocumentRecord | null {
  const isFolder = isSharePointFolder(item.fields);
  const name =
    asString(item.fields[documentFields.fileLeafRef]) ??
    asString(item.fields[documentFields.title]);
  if (!name) return null;

  const fileRef = asNullableString(item.fields[documentFields.fileRef]);
  const fileDirRef = asNullableString(item.fields[documentFields.fileDirRef]);
  const pathParts = customerDocumentsPathParts(fileRef ?? fileDirRef);
  const parentPath = pathParts.length > 0 ? pathParts.slice(0, -1) : [];

  const hasFile =
    !isFolder &&
    Boolean(
      fileRef || asString(item.fields[documentFields.fileLeafRef]),
    );

  const company = asLookupOrString(item.fields[documentFields.company]);
  const companyId =
    asString(item.fields[documentFields.companyLookupId]) ?? null;
  const candidate = asLookupOrString(item.fields[documentFields.candidate]);
  const candidateId =
    asString(item.fields[documentFields.candidateLookupId]) ?? null;
  const documentType = asNullableString(
    item.fields[documentFields.documentType],
  );
  const customerVisible = asBoolean(
    item.fields[documentFields.customerVisible],
  );
  const modifiedDate =
    item.lastModifiedDateTime ??
    asNullableString(item.fields[documentFields.modified]) ??
    item.createdDateTime ??
    null;

  return {
    id: item.id,
    name,
    company,
    companyId,
    candidate,
    candidateId,
    documentType,
    customerVisible,
    canDownload: hasFile,
    notificationSent: asBoolean(
      item.fields[documentFields.notificationSent],
    ),
    notifyCustomer: asBoolean(item.fields[documentFields.notifyCustomer]),
    modifiedDate,
    modifiedBy: asLookupOrString(item.fields[documentFields.editor]),
    metadataStatus: resolveDocumentMetadataStatus({
      company,
      documentType,
      customerVisible,
    }),
    isFolder,
    fileRef,
    parentPath,
    uploadedDate: modifiedDate,
    previewPath: hasFile
      ? `/api/admin/documents/${item.id}/download?disposition=inline`
      : null,
    downloadPath: hasFile
      ? `/api/admin/documents/${item.id}/download`
      : null,
  };
}

function matchesOptionalText(
  value: string | null | undefined,
  filter: string | null | undefined,
): boolean {
  if (!filter?.trim()) return true;
  return (value ?? "").trim().toLowerCase() === filter.trim().toLowerCase();
}

export async function listAdminDocuments(
  filters?: AdminDocumentListFilters | string | null,
) {
  // Back-compat: older callers passed companyName as a string.
  const normalized: AdminDocumentListFilters =
    typeof filters === "string" || filters === null || filters === undefined
      ? { companyName: filters }
      : filters;

  const items = await getListItemsByKey("customerDocuments", { top: 5000 });
  return items
    .map(mapDocument)
    .filter((row): row is AdminDocumentRecord => {
      if (!row) return false;
      // Keep folders for SharePoint-style library browse; filter files by metadata.
      if (!row.isFolder) {
        if (!matchesCompany(row.company, normalized.companyName)) return false;
        if (!matchesOptionalText(row.candidate, normalized.candidate)) {
          return false;
        }
        if (!matchesOptionalText(row.documentType, normalized.documentType)) {
          return false;
        }
        if (
          typeof normalized.customerVisible === "boolean" &&
          row.customerVisible !== normalized.customerVisible
        ) {
          return false;
        }
      }
      if (normalized.modifiedFrom || normalized.modifiedTo) {
        const modifiedMs = row.modifiedDate
          ? new Date(row.modifiedDate).getTime()
          : NaN;
        if (Number.isNaN(modifiedMs)) {
          if (!row.isFolder) return false;
        } else {
          if (normalized.modifiedFrom) {
            const from = new Date(normalized.modifiedFrom).getTime();
            if (!Number.isNaN(from) && modifiedMs < from) return false;
          }
          if (normalized.modifiedTo) {
            const to = new Date(normalized.modifiedTo);
            if (!Number.isNaN(to.getTime())) {
              to.setHours(23, 59, 59, 999);
              if (modifiedMs > to.getTime()) return false;
            }
          }
        }
      }
      return true;
    });
}

function buildDocumentWritePayload(
  input: Record<string, unknown>,
): SharePointFields {
  const payload: SharePointFields = {
    ...toSharePointFields("customerDocuments", {
      title: optionalText(input.name) ?? optionalText(input.title) ?? undefined,
      documentType:
        input.documentType === undefined
          ? undefined
          : optionalText(input.documentType),
      customerVisible: optionalBool(input.customerVisible),
      notificationSent: optionalBool(input.notificationSent),
      notifyCustomer: optionalBool(input.notifyCustomer),
    }),
  };

  if (input.companyId !== undefined) {
    const companyId = optionalText(input.companyId);
    payload[documentFields.companyLookupId] = companyId
      ? Number(companyId)
      : null;
  } else if (input.company !== undefined) {
    const companyName = optionalText(input.company);
    if (companyName) {
      payload[documentFields.company] = companyName;
    }
  }

  if (input.candidateId !== undefined) {
    const candidateId = optionalText(input.candidateId);
    payload[documentFields.candidateLookupId] = candidateId
      ? Number(candidateId)
      : null;
  } else if (input.candidate !== undefined) {
    const candidateName = optionalText(input.candidate);
    if (candidateName) {
      payload[documentFields.candidate] = candidateName;
    } else {
      payload[documentFields.candidateLookupId] = null;
    }
  }

  return payload;
}

export async function updateAdminDocument(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("customerDocuments", id);
  if (!existing) throw new NotFoundError("Document not found.");
  const payload = buildDocumentWritePayload(input);
  const item = await updateListItemFieldsByKey(
    "customerDocuments",
    id,
    payload,
  );
  const mapped = mapDocument(item);
  if (!mapped) throw new Error("Updated document could not be mapped.");
  return mapped;
}

export async function bulkUpdateAdminDocuments(
  ids: string[],
  input: Record<string, unknown>,
): Promise<{ updated: AdminDocumentRecord[]; failed: string[] }> {
  const uniqueIds = Array.from(
    new Set(ids.map((id) => String(id).trim()).filter(Boolean)),
  );
  if (uniqueIds.length === 0) {
    throw new ValidationError("Select at least one document.");
  }

  const updated: AdminDocumentRecord[] = [];
  const failed: string[] = [];

  for (const id of uniqueIds) {
    try {
      updated.push(await updateAdminDocument(id, input));
    } catch {
      failed.push(id);
    }
  }

  return { updated, failed };
}

function resolveDocumentMetadataStatusFromFields(input: {
  company: string | null;
  documentType: string | null;
  customerVisible: boolean;
}): DocumentMetadataStatus {
  if (!input.company?.trim()) return "Missing Company";
  if (!input.documentType?.trim()) return "Missing Document Type";
  if (!input.customerVisible) return "Hidden from Customer";
  return "Complete";
}

/**
 * Browse one folder level in Customer Documents (drive children).
 * pathSegments follow:
 * Company Number - Company Name / Company Documents|Candidates /
 * Candidate Number - Name / Certificates|Card Scans|…
 *
 * When opening a company or candidate folder, ensures the expected
 * subfolders exist (does not recreate the library or move files).
 */
export async function listAdminDocumentsAtPath(
  pathSegments: string[],
): Promise<AdminDocumentRecord[]> {
  const {
    browseCustomerDocumentsFolder,
    ensureCompanyDocumentFolders,
    ensureCandidateDocumentFolders,
  } = await import("@/lib/services/customerDocumentsFolderService");

  const parentPath = pathSegments
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (parentPath.length === 1) {
    const company = parseNumberNameFolder(parentPath[0]);
    await ensureCompanyDocumentFolders({
      companyNumber: company.number,
      companyName: company.name,
    });
  } else if (
    parentPath.length === 3 &&
    parentPath[1].toLowerCase() === "candidates"
  ) {
    const company = parseNumberNameFolder(parentPath[0]);
    const candidate = parseNumberNameFolder(parentPath[2]);
    await ensureCandidateDocumentFolders({
      companyNumber: company.number,
      companyName: company.name,
      workforceNumber: candidate.number,
      candidateName: candidate.name,
    });
  }

  const children = await browseCustomerDocumentsFolder(parentPath);

  return children.map((child): AdminDocumentRecord => {
    const fields = child.fields;
    const company = asLookupOrString(fields[documentFields.company]);
    const companyId =
      asString(fields[documentFields.companyLookupId]) ?? null;
    const candidate = asLookupOrString(fields[documentFields.candidate]);
    const candidateId =
      asString(fields[documentFields.candidateLookupId]) ?? null;
    const documentType = asNullableString(fields[documentFields.documentType]);
    const customerVisible = asBoolean(fields[documentFields.customerVisible]);
    const listId = child.listItemId ?? child.driveItemId;
    const hasFile = !child.isFolder;

    return {
      id: listId,
      name: child.name,
      company,
      companyId,
      candidate,
      candidateId,
      documentType,
      customerVisible,
      canDownload: hasFile,
      notificationSent: asBoolean(fields[documentFields.notificationSent]),
      notifyCustomer: asBoolean(fields[documentFields.notifyCustomer]),
      modifiedDate: child.lastModifiedDateTime,
      modifiedBy: asLookupOrString(fields[documentFields.editor]),
      metadataStatus: resolveDocumentMetadataStatusFromFields({
        company,
        documentType,
        customerVisible,
      }),
      isFolder: child.isFolder,
      fileRef: child.webUrl,
      parentPath,
      uploadedDate: child.lastModifiedDateTime,
      previewPath: hasFile
        ? `/api/admin/documents/${listId}/download?disposition=inline`
        : null,
      downloadPath: hasFile
        ? `/api/admin/documents/${listId}/download`
        : null,
    };
  });
}

function parseNumberNameFolder(segment: string): {
  number: string | null;
  name: string;
} {
  const trimmed = segment.trim();
  const match = trimmed.match(/^(.+?)\s+-\s+(.+)$/);
  if (!match) {
    return { number: null, name: trimmed };
  }
  return {
    number: match[1].trim() || null,
    name: match[2].trim() || trimmed,
  };
}

/* ───────────────── Events ───────────────── */

const eventFields = getSharePointFields("events");

export const EVENT_COMPANY_MISSING_WARNING =
  "Events list needs EventCompany lookup to Company List.";

export interface AdminEventRecord {
  id: string;
  title: string;
  /** Display name from EventCompany lookup only (never legacy Company). */
  company: string | null;
  companyId: string | null;
  customerVisible: boolean;
  trainingAddress: string | null;
  location: string | null;
  eventDate: string | null;
  endDate: string | null;
  description: string | null;
  doNotSync: boolean;
  syncStatus: string | null;
  syncDirection: string | null;
  lastSyncedAt: string | null;
  lastSyncSource: string | null;
  syncError: string | null;
  outlookEventId: string | null;
  outlookCalendarId: string | null;
  outlookICalUid: string | null;
  syncHash: string | null;
}

function asDateTimeInput(value: unknown): string | null | undefined {
  if (value === undefined) return undefined;
  if (value === null || value === "") return null;
  const text = String(value).trim();
  if (!text) return null;

  // datetime-local: 2026-07-25T10:00
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(text)) {
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? text : parsed.toISOString();
  }

  // date-only: keep as date midnight UTC-ish for SharePoint
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return `${text}T09:00:00.000Z`;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? text : parsed.toISOString();
}

function resolveEventCompanyId(item: SharePointListItem): string | null {
  return (
    asString(item.fields[eventFields.eventCompanyLookupId]) ??
    asString(item.fields.EventCompanyId) ??
    (typeof item.fields[eventFields.eventCompany] === "object"
      ? asString(
          (item.fields[eventFields.eventCompany] as { LookupId?: unknown })
            .LookupId,
        )
      : null) ??
    null
  );
}

function mapEvent(item: SharePointListItem): AdminEventRecord | null {
  const title = asString(item.fields[eventFields.title]);
  if (!title) return null;

  // Prefer LookupValue; Graph often only returns EventCompanyLookupId.
  const companyFromLookup = asLookupOrString(
    item.fields[eventFields.eventCompany],
  );
  const companyId = resolveEventCompanyId(item);

  return {
    id: item.id,
    title,
    company: companyFromLookup,
    companyId,
    customerVisible:
      asBoolean(item.fields[eventFields.customerVisible]) ||
      asBoolean(item.fields.CustomerVisible),
    trainingAddress: stripSharePointHtml(
      asNullableString(item.fields[eventFields.trainingAddress]),
    ),
    location: asNullableString(item.fields[eventFields.location]),
    eventDate: asNullableString(item.fields[eventFields.eventDate]),
    endDate: asNullableString(item.fields[eventFields.endDate]),
    description: stripSharePointHtml(
      asNullableString(item.fields[eventFields.description]),
    ),
    doNotSync: asBoolean(item.fields[eventFields.doNotSync]),
    syncStatus: asNullableString(item.fields[eventFields.syncStatus]),
    syncDirection: asNullableString(item.fields[eventFields.syncDirection]),
    lastSyncedAt: asNullableString(item.fields[eventFields.lastSyncedAt]),
    lastSyncSource: asNullableString(item.fields[eventFields.lastSyncSource]),
    syncError: asNullableString(item.fields[eventFields.syncError]),
    outlookEventId: asNullableString(item.fields[eventFields.outlookEventId]),
    outlookCalendarId: asNullableString(
      item.fields[eventFields.outlookCalendarId],
    ),
    outlookICalUid: asNullableString(item.fields[eventFields.outlookICalUid]),
    syncHash: asNullableString(item.fields[eventFields.syncHash]),
  };
}

async function attachEventCompanyNames(
  rows: AdminEventRecord[],
): Promise<AdminEventRecord[]> {
  const missingIds = Array.from(
    new Set(
      rows
        .filter((row) => !row.company?.trim() && row.companyId)
        .map((row) => row.companyId as string),
    ),
  );
  if (missingIds.length === 0) {
    return rows;
  }

  const companies = await listAdminCompanies();
  const byId = new Map(
    companies.map((company) => [company.id, company.companyName]),
  );

  // Fill any gaps with direct lookups (cache-friendly).
  await Promise.all(
    missingIds
      .filter((id) => !byId.has(id))
      .map(async (id) => {
        const company = await getCompanyById(id);
        if (company?.companyName) {
          byId.set(id, company.companyName);
        }
      }),
  );

  return rows.map((row) => {
    if (row.company?.trim() || !row.companyId) return row;
    return {
      ...row,
      company: byId.get(row.companyId) ?? null,
    };
  });
}

export async function getEventsSchemaWarnings(): Promise<string[]> {
  try {
    const hasEventCompany = await listHasColumn("events", "EventCompany");
    if (hasEventCompany) {
      return [];
    }
    return [EVENT_COMPANY_MISSING_WARNING];
  } catch {
    return [EVENT_COMPANY_MISSING_WARNING];
  }
}

export async function listAdminEvents(companyName?: string | null) {
  const items = await getListItemsByKey("events", { top: 5000 });
  const mapped = items
    .map(mapEvent)
    .filter((row): row is AdminEventRecord => row !== null);
  const withNames = await attachEventCompanyNames(mapped);

  return withNames
    .filter((row) => matchesCompany(row.company, companyName))
    .sort((a, b) => {
      const aTime = a.eventDate ? new Date(a.eventDate).getTime() : 0;
      const bTime = b.eventDate ? new Date(b.eventDate).getTime() : 0;
      return bTime - aTime;
    });
}

function buildEventCompanyPayload(input: Record<string, unknown>): {
  fields: SharePointFields;
  companyName: string | null;
  companyId: string | null;
} {
  const companyId =
    optionalText(input.companyId) ?? optionalText(input.eventCompanyId);
  const companyName =
    optionalText(input.company) ??
    optionalText(input.companyName) ??
    optionalText(input.eventCompany);

  const fields: SharePointFields = {};
  if (companyId) {
    // Graph lookup write — never send display name as EventCompany text.
    fields[eventFields.eventCompanyLookupId] = Number.isNaN(Number(companyId))
      ? companyId
      : Number(companyId);
  }

  return { fields, companyName, companyId };
}

/** SharePoint DoNotSync is Text ("Yes"/"No"), not Boolean. */
function toDoNotSyncText(value: boolean | undefined): "Yes" | "No" | undefined {
  if (value === undefined) return undefined;
  return value ? "Yes" : "No";
}

function omitNullishFields(fields: SharePointFields): SharePointFields {
  const next: SharePointFields = {};
  for (const [key, value] of Object.entries(fields)) {
    if (value === null || value === undefined) continue;
    next[key] = value;
  }
  return next;
}

function rethrowEventWriteError(error: unknown): never {
  const message =
    error && typeof error === "object" && "message" in error
      ? String((error as { message?: unknown }).message ?? "")
      : String(error ?? "");
  const body =
    error && typeof error === "object" && "body" in error
      ? String((error as { body?: unknown }).body ?? "")
      : "";
  const combined = `${message}\n${body}`;

  if (/unique constraints/i.test(combined)) {
    throw new ValidationError(
      'SharePoint "Event Company" has Enforce unique values enabled, so only one event can use each company. In SharePoint Events list settings, open the Event Company column, turn off Enforce unique values, then try again.',
    );
  }

  throw error;
}

export async function createAdminEvent(input: Record<string, unknown>) {
  const title = requireText(input.title, "Event title");
  const { fields: companyFields, companyName, companyId } =
    buildEventCompanyPayload(input);
  if (!companyFields[eventFields.eventCompanyLookupId]) {
    throw new ValidationError(
      "Company is required. Select a company from the list.",
    );
  }

  const doNotSync = optionalBool(input.doNotSync) ?? false;
  const payload = omitNullishFields({
    ...toSharePointFields("events", {
      title,
      customerVisible: optionalBool(input.customerVisible) ?? true,
      trainingAddress: optionalText(input.trainingAddress),
      location: optionalText(input.location),
      eventDate: asDateTimeInput(input.eventDate),
      endDate: asDateTimeInput(input.endDate),
      description: optionalText(input.description),
      // Text column — boolean true/false causes Graph generalException.
      doNotSync: toDoNotSyncText(doNotSync),
      syncStatus: doNotSync ? "Skipped" : "Pending",
      syncDirection: "SharePointToOutlook",
      lastSyncSource: "SharePoint",
    }),
    ...companyFields,
  });

  let item;
  try {
    item = await createListItemByKey("events", payload);
  } catch (error) {
    rethrowEventWriteError(error);
  }

  const { syncEventSharePointToOutlook } = await import(
    "@/lib/services/eventOutlookSyncService"
  );
  await syncEventSharePointToOutlook(item.id);

  const refreshed = await getListItemByKey("events", item.id);
  const mapped = mapEvent(refreshed ?? item);
  if (!mapped) throw new Error("Created event could not be mapped.");

  const [withName] = await attachEventCompanyNames([
    {
      ...mapped,
      company: mapped.company ?? companyName,
      companyId: mapped.companyId ?? companyId,
    },
  ]);
  return withName;
}

export async function updateAdminEvent(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("events", id);
  if (!existing) throw new NotFoundError("Event not found.");

  const doNotSync =
    input.doNotSync === undefined ? undefined : optionalBool(input.doNotSync);

  const payload = omitNullishFields(
    toSharePointFields("events", {
      title: optionalText(input.title) ?? undefined,
      customerVisible: optionalBool(input.customerVisible),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      location:
        input.location === undefined ? undefined : optionalText(input.location),
      eventDate:
        input.eventDate === undefined
          ? undefined
          : asDateTimeInput(input.eventDate),
      endDate:
        input.endDate === undefined ? undefined : asDateTimeInput(input.endDate),
      description:
        input.description === undefined
          ? undefined
          : optionalText(input.description),
      doNotSync: toDoNotSyncText(doNotSync),
    }),
  );

  const companyTouched =
    input.companyId !== undefined ||
    input.company !== undefined ||
    input.companyName !== undefined ||
    input.eventCompany !== undefined ||
    input.eventCompanyId !== undefined;

  if (companyTouched) {
    const { fields: companyFields } = buildEventCompanyPayload(input);
    if (!companyFields[eventFields.eventCompanyLookupId]) {
      throw new ValidationError(
        "Company is required. Select a company from the list.",
      );
    }
    Object.assign(payload, companyFields);
  }

  // Portal edits are SharePoint-sourced; prepare one-way Outlook sync metadata.
  if (doNotSync === true) {
    payload[eventFields.syncStatus] = "Skipped";
    payload[eventFields.syncDirection] = "SharePointToOutlook";
    payload[eventFields.lastSyncSource] = "SharePoint";
  } else if (
    doNotSync === false ||
    input.title !== undefined ||
    input.eventDate !== undefined ||
    input.endDate !== undefined ||
    input.location !== undefined ||
    input.trainingAddress !== undefined ||
    input.description !== undefined ||
    companyTouched
  ) {
    payload[eventFields.syncStatus] = "Pending";
    payload[eventFields.syncDirection] = "SharePointToOutlook";
    payload[eventFields.lastSyncSource] = "SharePoint";
  }

  let item;
  try {
    item = await updateListItemFieldsByKey("events", id, payload);
  } catch (error) {
    rethrowEventWriteError(error);
  }

  const { syncEventSharePointToOutlook } = await import(
    "@/lib/services/eventOutlookSyncService"
  );
  await syncEventSharePointToOutlook(id);

  const refreshed = await getListItemByKey("events", id);
  const mapped = mapEvent(refreshed ?? item);
  if (!mapped) throw new Error("Updated event could not be mapped.");
  const [withName] = await attachEventCompanyNames([mapped]);
  return withName;
}

/* ───────────────── Offers ───────────────── */

const offerFields = getSharePointFields("offersPromotions");

export interface AdminOfferRecord {
  id: string;
  title: string;
  category: string | null;
  customerVisible: boolean;
  startDate: string | null;
  endDate: string | null;
  description: string | null;
  status: string | null;
}

function mapOffer(item: SharePointListItem): AdminOfferRecord | null {
  const title = asString(item.fields[offerFields.title]);
  if (!title) return null;
  return {
    id: item.id,
    title,
    category: asNullableString(item.fields[offerFields.category]),
    customerVisible: asBoolean(item.fields[offerFields.customerVisible]),
    startDate: asNullableString(item.fields[offerFields.startDate]),
    endDate: asNullableString(item.fields[offerFields.endDate]),
    description: asNullableString(item.fields[offerFields.shortDescription]),
    status: asNullableString(item.fields[offerFields.status]),
  };
}

export async function listAdminOffers() {
  const items = await getListItemsByKey("offersPromotions", { top: 5000 });
  return items
    .map(mapOffer)
    .filter((row): row is AdminOfferRecord => row !== null)
    .sort((a, b) => a.title.localeCompare(b.title));
}

export async function createAdminOffer(input: Record<string, unknown>) {
  const title = requireText(input.title, "Offer title");
  const payload = toSharePointFields("offersPromotions", {
    title,
    category: optionalText(input.category),
    customerVisible: optionalBool(input.customerVisible) ?? true,
    startDate: asDateInput(input.startDate),
    endDate: asDateInput(input.endDate),
    shortDescription:
      optionalText(input.description) ??
      optionalText(input.shortDescription),
    status: optionalText(input.status) ?? "Active",
  });
  const item = await createListItemByKey("offersPromotions", payload);
  const mapped = mapOffer(item);
  if (!mapped) throw new Error("Created offer could not be mapped.");
  return mapped;
}

export async function updateAdminOffer(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("offersPromotions", id);
  if (!existing) throw new NotFoundError("Offer not found.");
  const payload = toSharePointFields("offersPromotions", {
    title: optionalText(input.title) ?? undefined,
    category:
      input.category === undefined ? undefined : optionalText(input.category),
    customerVisible: optionalBool(input.customerVisible),
    startDate:
      input.startDate === undefined
        ? undefined
        : asDateInput(input.startDate),
    endDate:
      input.endDate === undefined ? undefined : asDateInput(input.endDate),
    shortDescription:
      input.description === undefined && input.shortDescription === undefined
        ? undefined
        : (optionalText(input.description) ??
          optionalText(input.shortDescription)),
    status: optionalText(input.status) ?? undefined,
  });
  const item = await updateListItemFieldsByKey(
    "offersPromotions",
    id,
    payload,
  );
  const mapped = mapOffer(item);
  if (!mapped) throw new Error("Updated offer could not be mapped.");
  return mapped;
}

/* ───────────────── Permissions ───────────────── */

const permissionFields = getSharePointFields("permissions");

export interface AdminPermissionRecord {
  id: string;
  userEmail: string;
  roleType: RoleType;
  status: string;
  companyId: string | null;
  companyName: string | null;
  accessScope: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
}

function mapPermission(item: SharePointListItem): AdminPermissionRecord | null {
  const userEmail = asString(item.fields[permissionFields.userEmail]);
  const roleType = normalizePermissionRoleType(
    item.fields[permissionFields.roleType],
  );
  if (!userEmail || !roleType) return null;
  const companyId =
    asString(item.fields[permissionFields.companyLookupId]) ??
    (typeof item.fields[permissionFields.company] === "object"
      ? asString(
          (item.fields[permissionFields.company] as { LookupId?: unknown })
            .LookupId,
        )
      : null) ??
    null;
  return {
    id: item.id,
    userEmail: userEmail.toLowerCase(),
    roleType,
    status: asNullableString(item.fields[permissionFields.status]) ?? "Inactive",
    companyId,
    companyName: asLookupOrString(item.fields[permissionFields.company]),
    accessScope: asNullableString(item.fields[permissionFields.accessScope]),
    canView: asBoolean(item.fields[permissionFields.canView]),
    canDownload: asBoolean(item.fields[permissionFields.canDownload]),
    canEdit: asBoolean(item.fields[permissionFields.canEdit]),
  };
}

export async function listAdminPermissions() {
  const items = await getListItemsByKey("permissions", { top: 5000 });
  return items
    .map(mapPermission)
    .filter((row): row is AdminPermissionRecord => row !== null)
    .sort((a, b) => a.userEmail.localeCompare(b.userEmail));
}

export async function createAdminPermission(input: Record<string, unknown>) {
  const userEmail = requireText(input.userEmail, "User email").toLowerCase();
  const roleType = normalizePermissionRoleType(input.roleType);
  if (!roleType) {
    throw new ValidationError(
      "Role must be Training Manager (Admin) or Supervisor (Customer).",
    );
  }
  const companyId = optionalText(input.companyId);
  const companyName = optionalText(input.companyName);
  if (roleType === "Customer" && !companyId && !companyName) {
    throw new ValidationError("Customer permissions require a company.");
  }

  const payload: SharePointFields = toSharePointFields("permissions", {
    userEmail,
    roleType: toSharePointRoleType(roleType),
    status: optionalText(input.status) ?? "Active",
    accessScope: optionalText(input.accessScope) ?? "Company",
    canView: optionalBool(input.canView) ?? true,
    canDownload: optionalBool(input.canDownload) ?? false,
    canEdit: optionalBool(input.canEdit) ?? false,
  });

  if (companyId) {
    payload[permissionFields.companyLookupId] = Number.isNaN(Number(companyId))
      ? companyId
      : Number(companyId);
  } else if (companyName) {
    payload[permissionFields.company] = companyName;
  }

  const item = await createListItemByKey("permissions", payload);
  const mapped = mapPermission(item);
  if (!mapped) throw new Error("Created permission could not be mapped.");
  return mapped;
}

export async function updateAdminPermission(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("permissions", id);
  if (!existing) throw new NotFoundError("Permission not found.");

  let sharePointRole: string | undefined;
  if (input.roleType !== undefined) {
    const roleType = normalizePermissionRoleType(input.roleType);
    if (!roleType) {
      throw new ValidationError(
        "Role must be Training Manager (Admin) or Supervisor (Customer).",
      );
    }
    sharePointRole = toSharePointRoleType(roleType);
  }

  const payload: SharePointFields = toSharePointFields("permissions", {
    userEmail:
      input.userEmail === undefined
        ? undefined
        : requireText(input.userEmail, "User email").toLowerCase(),
    roleType: sharePointRole,
    status: optionalText(input.status) ?? undefined,
    accessScope:
      input.accessScope === undefined
        ? undefined
        : optionalText(input.accessScope),
    canView: optionalBool(input.canView),
    canDownload: optionalBool(input.canDownload),
    canEdit: optionalBool(input.canEdit),
  });

  if (input.companyId !== undefined) {
    const companyId = optionalText(input.companyId);
    if (companyId) {
      payload[permissionFields.companyLookupId] = Number.isNaN(
        Number(companyId),
      )
        ? companyId
        : Number(companyId);
    }
  } else if (input.companyName !== undefined) {
    payload[permissionFields.company] = optionalText(input.companyName);
  }

  const item = await updateListItemFieldsByKey("permissions", id, payload);
  const mapped = mapPermission(item);
  if (!mapped) throw new Error("Updated permission could not be mapped.");
  return mapped;
}
