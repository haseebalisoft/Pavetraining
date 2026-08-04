import "server-only";

import { allocateNextCompanyNumber } from "@/lib/companyNumber";
import { allocateNextWorkforceNumber } from "@/lib/workforceNumber";
import { NotFoundError } from "@/lib/services/errorHandler";
import { ValidationError } from "@/lib/services/validationService";
import { assertNotProtectedAdmin } from "@/lib/auth/protectedAdmins";
import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import {
  asBoolean,
  asLookupOrString,
  asMultiChoiceText,
  asNullableString,
  asString,
  buildSchemaFieldEqualsFilter,
  createListItemByKey,
  deleteListItemByKey,
  extractLookupId,
  getListItemById,
  getListItemByKey,
  getListItems,
  getListItemsByKey,
  listHasColumn,
  toSharePointFields,
  updateListItemFieldsByKey,
  type SharePointFields,
  type SharePointListItem,
} from "@/lib/services/sharePointListService";
import { getSharePointListId } from "@/lib/config/sharepoint";
import { revalidateSharePointList } from "@/lib/cache/sharePointCache";
import { revalidatePath } from "next/cache";
import {
  normalizePermissionFormRole,
  normalizePermissionRoleType,
  permissionFormRoleFromSharePoint,
  resolveCustomerRole,
  roleLabelFor,
  toSharePointRoleType,
  type PermissionFormRole,
} from "@/lib/services/permissionService";
import { mapCompanyFields, getCompanyById } from "@/lib/services/companyService";
import { parseThumbnailField } from "@/lib/services/listThumbnailService";
import {
  CLIENT_MATRIX_CATEGORY_COLUMNS,
  CLIENT_MATRIX_DISPLAY_HEADERS,
} from "@/lib/services/bulkUpload/clientTemplateHeaders";
import {
  earliestDateFromColumns,
  listTrainingMatrixExampleRows,
  stripExampleMatrixId,
  upsertTrainingMatrixExampleRow,
} from "@/lib/services/bulkUpload/trainingMatrixExampleService";

import { stripSharePointHtml } from "@/lib/text/stripSharePointHtml";
import type { Company, CustomerRoleType, RoleType } from "@/types/models";
import type {
  AdminDocumentRecord,
  DocumentMetadataStatus,
} from "@/types/adminDocuments";
import {
  bookingStatusFromFreeBusy,
  freeBusyFromBookingStatus,
  normalizeBookingStatus,
  type BookingStatus,
} from "@/lib/services/bookingStatusService";

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

/** SharePoint Workforce Status choices are `Active` and `inactive`. */
function normalizeWorkforceStatus(
  value: unknown,
  fallback: string | null = "Active",
): string | null {
  const text = optionalText(value);
  if (!text) return fallback;
  const key = text.toLowerCase();
  if (key === "active") return "Active";
  if (key === "inactive") return "inactive";
  return text;
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
        : normalizeCompanySize(optionalText(input.companySize)),
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
    // CompanyLogo is Thumbnail — never write free text / placeholders via Graph.
    status:
      input.status === undefined
        ? partial
          ? undefined
          : "Active"
        : normalizeCompanyStatus(optionalText(input.status)),
  });
}

/** SharePoint Company Size choices: Small | Medium | Large | Enterprise. */
function normalizeCompanySize(value: string | null | undefined): string | null {
  if (!value?.trim()) return null;
  const key = value.trim().toLowerCase();
  if (key === "small") return "Small";
  if (key === "medium") return "Medium";
  if (key === "large") return "Large";
  if (key === "enterprise") return "Enterprise";
  return value.trim();
}

/** SharePoint Company Status choices: Active | Inactive. */
function normalizeCompanyStatus(
  value: string | null | undefined,
): string | null {
  if (!value?.trim()) return "Active";
  const key = value.trim().toLowerCase();
  if (key === "active") return "Active";
  if (key === "inactive") return "Inactive";
  if (key === "on hold" || key === "onhold" || key === "hold") return "Inactive";
  return "Active";
}

export async function listAdminCompanies(): Promise<Company[]> {
  const items = await getListItemsByKey("company", { top: 5000 });
  return items
    .map((item) => mapCompanyFields(item.id, item.fields))
    .filter((row): row is Company => row !== null)
    .sort((a, b) => a.companyName.localeCompare(b.companyName));
}

export async function createAdminCompany(input: Record<string, unknown>) {
  // Bulk callers pass the live in-memory list so we don't re-fetch Companies
  // from Graph for every new company in the same import (was ~1s×N).
  const companies = Array.isArray(input.existingCompanies)
    ? (input.existingCompanies as Company[])
    : await listAdminCompanies();
  const provided = optionalText(input.companyNumber);
  const companyNumber = provided || allocateNextCompanyNumber(companies);

  if (
    companies.some(
      (row) =>
        (row.companyNumber ?? "").trim().toLowerCase() ===
        companyNumber.toLowerCase(),
    )
  ) {
    throw new ValidationError(
      `Company number "${companyNumber}" is already in use.`,
    );
  }

  const payload = companyWritePayload(
    { ...input, companyNumber },
    false,
  );
  const item = await createListItemByKey("company", payload);
  const mapped = mapCompanyFields(item.id, item.fields);
  if (!mapped) throw new Error("Created company could not be mapped.");

  const skipFolders =
    input.skipDocumentFolders === true ||
    input.skipDocumentFolders === "true" ||
    input.bulkMode === true ||
    input.bulkMode === "true";

  if (skipFolders) {
    return { ...mapped, folderWarning: undefined };
  }

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

export async function deleteAdminCompany(id: string) {
  const {
    deleteCompanyWithRelatedData,
  } = await import("@/lib/services/companyCascadeDeleteService");
  const result = await deleteCompanyWithRelatedData(id);
  if (!result.companyDeleted) {
    throw new Error(
      result.errors.join(" | ") || "Company could not be deleted.",
    );
  }
  return result;
}

export async function bulkDeleteAdminCompanies(ids: string[]) {
  const unique = Array.from(
    new Set(ids.map((id) => String(id).trim()).filter(Boolean)),
  );
  if (unique.length === 0) {
    throw new Error("No company ids provided.");
  }
  const {
    deleteCompaniesWithRelatedData,
  } = await import("@/lib/services/companyCascadeDeleteService");
  return deleteCompaniesWithRelatedData(unique);
}

/* ───────────────── Workforce ───────────────── */

const workforceFields = getSharePointFields("workforce");

export interface AdminWorkforceRecord {
  id: string;
  workforceNumber: string | null;
  candidateName: string;
  companyId: string | null;
  companyName: string;
  companyNumber: string | null;
  trainingManager: string | null;
  supervisor: string | null;
  candidateAddress: string | null;
  email: string | null;
  contactNumber: string | null;
  dateOfBirth: string | null;
  niNumber: string | null;
  nporsNumbers: string | null;
  cscsNumber: string | null;
  cscsExpiry: string | null;
  swqrNumber: string | null;
  swqrExpiry: string | null;
  eusrNumber: string | null;
  eusrExpiry: string | null;
  inHouseCertificationNumber: string | null;
  department: string | null;
  status: string | null;
  notes: string | null;
  /** App media URL for Candidate Photo thumbnail, when set. */
  photoUrl: string | null;
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

/** Permissions List rows used to resolve Workforce Training manager / Supervisor lookups. */
export type PermissionPersonRef = {
  id: string;
  name: string | null;
  userEmail: string;
  roleType: RoleType;
  companyId: string | null;
  status: string;
};

function permissionPersonKey(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

export async function loadPermissionPeople(): Promise<PermissionPersonRef[]> {
  const fields = getSharePointFields("permissions");
  const items = await getListItemsByKey("permissions", { top: 5000 });
  const rows: PermissionPersonRef[] = [];
  for (const item of items) {
    const userEmail = asString(item.fields[fields.userEmail]);
    const roleType = normalizePermissionRoleType(item.fields[fields.roleType]);
    if (!userEmail || !roleType) continue;
    const companyId =
      asString(item.fields[fields.companyLookupId]) ??
      extractLookupId(item.fields, fields.company) ??
      null;
    rows.push({
      id: item.id,
      name: asNullableString(item.fields[fields.name]),
      userEmail: userEmail.toLowerCase(),
      roleType,
      companyId,
      status: asNullableString(item.fields[fields.status]) ?? "Inactive",
    });
  }
  return rows;
}

function permissionNameByIdMap(
  people: PermissionPersonRef[],
): Map<string, string> {
  const map = new Map<string, string>();
  for (const person of people) {
    const label = person.name?.trim() || person.userEmail;
    if (label) map.set(person.id, label);
  }
  return map;
}

export function findPermissionPerson(
  people: PermissionPersonRef[],
  displayNameOrEmail: string | null | undefined,
): PermissionPersonRef | null {
  const raw = displayNameOrEmail?.trim();
  if (!raw) return null;
  const key = permissionPersonKey(raw);
  return (
    people.find((row) => permissionPersonKey(row.name ?? "") === key) ??
    people.find((row) => row.userEmail === key) ??
    null
  );
}

/**
 * Resolve or create a Permissions List row for Workforce TM/Supervisor lookups.
 * SharePoint scheme: these columns are Lookups to Permissions.Name.
 */
export async function ensurePermissionPerson(input: {
  displayName: string;
  roleType: RoleType;
  companyId?: string | null;
  people?: PermissionPersonRef[];
}): Promise<{ person: PermissionPersonRef; people: PermissionPersonRef[]; created: boolean }> {
  const displayName = input.displayName.trim();
  if (!displayName) {
    throw new ValidationError("Display name is required for permission person.");
  }

  let people = input.people ?? (await loadPermissionPeople());
  const existing = findPermissionPerson(people, displayName);
  if (existing) {
    return { person: existing, people, created: false };
  }

  const slug = displayName
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "")
    .slice(0, 40);
  const userEmail = `import.${slug || "person"}.${Date.now()}@pave.local`;
  const fields = getSharePointFields("permissions");
  const payload: SharePointFields = toSharePointFields("permissions", {
    userEmail,
    roleType: toSharePointRoleType(input.roleType),
    status: "Active",
    accessScope: "Full Company",
    canView: true,
    canDownload: false,
    canEdit: false,
    name: displayName,
  });
  if (input.companyId) {
    payload[fields.companyLookupId] = Number.isNaN(Number(input.companyId))
      ? input.companyId
      : Number(input.companyId);
  }

  const item = await createListItemByKey("permissions", payload);
  const person: PermissionPersonRef = {
    id: item.id,
    name: displayName,
    userEmail,
    roleType: input.roleType,
    companyId: input.companyId ?? null,
    status: "Active",
  };
  people = [...people, person];
  return { person, people, created: true };
}

async function applyWorkforcePersonLookups(
  payload: SharePointFields,
  input: Record<string, unknown>,
  options?: {
    companyId?: string | null;
    createIfMissing?: boolean;
    people?: PermissionPersonRef[];
  },
): Promise<{
  people: PermissionPersonRef[];
  trainingManagerName: string | null;
  supervisorName: string | null;
}> {
  // Never write free text into Lookup columns — SharePoint rejects / ignores it.
  delete payload[workforceFields.trainingManager];
  delete payload[workforceFields.supervisor];

  let people = options?.people ?? (await loadPermissionPeople());
  let trainingManagerName: string | null = null;
  let supervisorName: string | null = null;

  const resolveOne = async (
    value: string | null | undefined,
    roleType: RoleType,
    lookupIdField: string,
  ): Promise<string | null> => {
    const text = value?.trim();
    if (!text) {
      payload[lookupIdField] = null;
      return null;
    }
    const key = permissionPersonKey(text);
    const companyId = options?.companyId ?? null;
    let hit =
      people.find((row) => {
        if (row.roleType !== roleType) return false;
        if ((row.status || "").toLowerCase() !== "active") return false;
        if (companyId && row.companyId && row.companyId !== companyId) {
          return false;
        }
        return (
          permissionPersonKey(row.name ?? "") === key || row.userEmail === key
        );
      }) ?? null;
    if (!hit) {
      hit = findPermissionPerson(people, text);
      if (hit && hit.roleType !== roleType) {
        hit = null;
      }
    }
    if (!hit && options?.createIfMissing) {
      const ensured = await ensurePermissionPerson({
        displayName: text,
        roleType,
        companyId: options.companyId,
        people,
      });
      people = ensured.people;
      hit = ensured.person;
    }
    if (!hit) {
      throw new ValidationError(
        `"${text}" was not found in Permissions as an active ${
          roleType === "Admin" ? "Training Manager" : "Supervisor"
        } for this company. Add them under Permissions first.`,
      );
    }
    payload[lookupIdField] = Number(hit.id);
    return hit.name?.trim() || hit.userEmail;
  };

  if (input.trainingManager !== undefined) {
    trainingManagerName = await resolveOne(
      optionalText(input.trainingManager),
      "Admin",
      `${workforceFields.trainingManager}LookupId`,
    );
  }
  if (input.supervisor !== undefined) {
    supervisorName = await resolveOne(
      optionalText(input.supervisor),
      "Customer",
      `${workforceFields.supervisor}LookupId`,
    );
  }

  return { people, trainingManagerName, supervisorName };
}

type DepartmentRef = {
  id: string;
  name: string;
};

function departmentKey(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

export async function loadDepartments(): Promise<DepartmentRef[]> {
  const fields = getSharePointFields("departments");
  const items = await getListItemsByKey("departments", { top: 5000 });
  const rows: DepartmentRef[] = [];
  for (const item of items) {
    const name =
      asNullableString(item.fields[fields.name]) ??
      asNullableString(item.fields[fields.title]);
    if (!name) continue;
    rows.push({ id: item.id, name });
  }
  return rows;
}

function departmentNameByIdMap(
  departments: DepartmentRef[],
): Map<string, string> {
  return new Map(departments.map((row) => [row.id, row.name] as const));
}

export function findDepartment(
  departments: DepartmentRef[],
  name: string | null | undefined,
): DepartmentRef | null {
  const raw = name?.trim();
  if (!raw) return null;
  const key = departmentKey(raw);
  return departments.find((row) => departmentKey(row.name) === key) ?? null;
}

/**
 * Resolve or create a Departments list row for Workforce ` Department` lookup.
 */
export async function ensureDepartment(input: {
  name: string;
  departments?: DepartmentRef[];
}): Promise<{
  department: DepartmentRef;
  departments: DepartmentRef[];
  created: boolean;
}> {
  const name = input.name.trim();
  if (!name) {
    throw new ValidationError("Department name is required.");
  }

  let departments = input.departments ?? (await loadDepartments());
  const existing = findDepartment(departments, name);
  if (existing) {
    return { department: existing, departments, created: false };
  }

  const fields = getSharePointFields("departments");
  const payload: SharePointFields = toSharePointFields("departments", {
    title: name,
    name,
  });
  const item = await createListItemByKey("departments", payload);
  const department: DepartmentRef = {
    id: item.id,
    name:
      asNullableString(item.fields[fields.name]) ??
      asNullableString(item.fields[fields.title]) ??
      name,
  };
  departments = [...departments, department];
  return { department, departments, created: true };
}

async function applyWorkforceDepartmentLookup(
  payload: SharePointFields,
  input: Record<string, unknown>,
  options?: {
    createIfMissing?: boolean;
    departments?: DepartmentRef[];
    companyId?: string | null;
    companyName?: string | null;
    /** Prefetched company departments — avoids a full Departments+Companies Graph read per row. */
    departmentRecords?: Array<{
      id: string;
      name: string;
      companyId?: string | null;
      companyName?: string | null;
    }>;
  },
): Promise<{
  departments: DepartmentRef[];
  departmentName: string | null;
  departmentRecords: Array<{
    id: string;
    name: string;
    companyId?: string | null;
    companyName?: string | null;
  }>;
}> {
  // Never write free text into Lookup column Department0.
  delete payload[workforceFields.departmentText];
  delete payload[workforceFields.department];

  let departments = options?.departments ?? (await loadDepartments());
  let departmentRecords = options?.departmentRecords
    ? [...options.departmentRecords]
    : [];
  let departmentName: string | null = null;

  const departmentInput =
    input.departmentText !== undefined
      ? input.departmentText
      : input.department !== undefined
        ? input.department
        : undefined;

  if (departmentInput === undefined) {
    return { departments, departmentName: null, departmentRecords };
  }

  const text = optionalText(departmentInput);
  if (!text) {
    payload[`${workforceFields.departmentText}LookupId`] = null;
    return { departments, departmentName: null, departmentRecords };
  }

  // Prefer company-scoped Departments list (Enterprise).
  if (options?.companyId) {
    const { createAdminDepartment, listAdminDepartments } = await import(
      "@/lib/services/departmentService"
    );
    let companyDepts =
      options.departmentRecords !== undefined
        ? departmentRecords.filter(
            (row) =>
              !row.companyId ||
              row.companyId === options.companyId ||
              !options.companyId,
          )
        : await listAdminDepartments(options.companyId);

    const createIfMissing = Boolean(options.createIfMissing);
    let resolved =
      companyDepts.find((row) => row.id === text) ??
      companyDepts.find(
        (row) => departmentKey(row.name) === departmentKey(text),
      ) ??
      null;

    if (!resolved && createIfMissing) {
      resolved = await createAdminDepartment({
        name: text,
        companyId: options.companyId,
        companyName: options.companyName,
        // Caller already scanned the in-memory company dept list.
        skipDuplicateScan: options.departmentRecords !== undefined,
      });
      companyDepts = [...companyDepts, resolved];
    }

    if (!resolved) {
      throw new ValidationError(
        createIfMissing
          ? `Department "${text}" could not be created for this company. Add it under Admin → Departments and try again.`
          : `Department "${text}" was not found for this company. Add it under Admin → Departments first.`,
      );
    }
    payload[`${workforceFields.departmentText}LookupId`] = Number(resolved.id);
    departmentName = resolved.name;
    departments = companyDepts.map((row) => ({ id: row.id, name: row.name }));
    departmentRecords = companyDepts;
    return { departments, departmentName, departmentRecords };
  }

  let hit = findDepartment(departments, text);
  if (!hit && options?.createIfMissing) {
    const ensured = await ensureDepartment({ name: text, departments });
    departments = ensured.departments;
    hit = ensured.department;
  }
  if (!hit) {
    throw new ValidationError(
      `Department "${text}" was not found in Departments. Add it under Departments or re-import to auto-create.`,
    );
  }

  payload[`${workforceFields.departmentText}LookupId`] = Number(hit.id);
  departmentName = hit.name;
  return { departments, departmentName, departmentRecords };
}

function mapWorkforce(
  item: SharePointListItem,
  companyNameById?: Map<string, string>,
  companyNumberById?: Map<string, string | null>,
  permissionNameById?: Map<string, string>,
  departmentNameById?: Map<string, string>,
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

  const companyNumber =
    asNullableString(item.fields.Company_x0020_Name_x003a__x0020_) ??
    (companyLookupId && companyNumberById
      ? (companyNumberById.get(companyLookupId) ?? null)
      : null);

  const trainingManagerLookupId = extractLookupId(
    item.fields,
    workforceFields.trainingManager,
  );
  const supervisorLookupId = extractLookupId(
    item.fields,
    workforceFields.supervisor,
  );
  const departmentLookupId = extractLookupId(
    item.fields,
    workforceFields.departmentText,
  );

  return {
    id: item.id,
    workforceNumber: asNullableString(
      item.fields[workforceFields.workforceNumber],
    ),
    candidateName,
    companyId: companyLookupId,
    companyName,
    companyNumber,
    trainingManager:
      asLookupOrString(item.fields[workforceFields.trainingManager]) ??
      asNullableString(item.fields[workforceFields.trainingManager]) ??
      (trainingManagerLookupId && permissionNameById
        ? (permissionNameById.get(trainingManagerLookupId) ?? null)
        : null),
    supervisor:
      asLookupOrString(item.fields[workforceFields.supervisor]) ??
      asNullableString(item.fields[workforceFields.supervisor]) ??
      (supervisorLookupId && permissionNameById
        ? (permissionNameById.get(supervisorLookupId) ?? null)
        : null),
    candidateAddress: asNullableString(
      item.fields[workforceFields.candidateAddress],
    ),
    email: asNullableString(item.fields[workforceFields.email]),
    contactNumber: asNullableString(item.fields[workforceFields.contactNumber]),
    dateOfBirth: asNullableString(item.fields[workforceFields.dateOfBirth]),
    niNumber: asNullableString(item.fields[workforceFields.niNumber]),
    nporsNumbers: asNullableString(item.fields[workforceFields.nporsNumbers]),
    cscsNumber: asNullableString(item.fields[workforceFields.cscsNumber]),
    cscsExpiry: asNullableString(item.fields[workforceFields.cscsExpiry]),
    swqrNumber: asNullableString(item.fields[workforceFields.swqrNumber]),
    swqrExpiry: asNullableString(item.fields[workforceFields.swqrExpiry]),
    eusrNumber: asNullableString(item.fields[workforceFields.eusrNumber]),
    eusrExpiry: asNullableString(item.fields[workforceFields.eusrExpiry]),
    inHouseCertificationNumber: asNullableString(
      item.fields[workforceFields.inHouseCertificationNumber],
    ),
    department:
      mapWorkforceDepartment(item.fields[workforceFields.departmentText]) ??
      (departmentLookupId && departmentNameById
        ? (departmentNameById.get(departmentLookupId) ?? null)
        : null) ??
      mapWorkforceDepartment(item.fields[workforceFields.department]),
    status: asNullableString(item.fields[workforceFields.status]),
    notes: asNullableString(item.fields[workforceFields.notes]),
    photoUrl: parseThumbnailField(item.fields[workforceFields.photo])
      ? `/api/media/workforce/${item.id}/photo`
      : null,
  };
}

export async function listAdminWorkforce(
  companyName?: string | null,
  options?: {
    companies?: Awaited<ReturnType<typeof listAdminCompanies>>;
    people?: PermissionPersonRef[];
    departments?: DepartmentRef[];
  },
) {
  const [items, companies, people, departments] = await Promise.all([
    getListItemsByKey("workforce", { top: 5000 }),
    options?.companies
      ? Promise.resolve(options.companies)
      : listAdminCompanies(),
    options?.people
      ? Promise.resolve(options.people)
      : loadPermissionPeople(),
    options?.departments
      ? Promise.resolve(options.departments)
      : loadDepartments(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const companyNumberById = new Map(
    companies.map((row) => [row.id, row.companyNumber] as const),
  );
  const permissionNameById = permissionNameByIdMap(people);
  const departmentNameById = departmentNameByIdMap(departments);
  return items
    .map((item) =>
      mapWorkforce(
        item,
        companyNameById,
        companyNumberById,
        permissionNameById,
        departmentNameById,
      ),
    )
    .filter((row): row is AdminWorkforceRecord => {
      if (!row) return false;
      return matchesCompany(row.companyName, companyName);
    })
    .sort((a, b) => a.candidateName.localeCompare(b.candidateName));
}

export async function createAdminWorkforce(input: Record<string, unknown>) {
  const candidateName = requireText(input.candidateName, "Candidate name");
  const companyName = requireText(input.companyName, "Company");
  const providedWorkforceNumber = optionalText(input.workforceNumber);
  const providedCompanyId = optionalText(input.companyId);
  const isBulk =
    input.bulkMode === true ||
    input.bulkMode === "true";

  let company: Company | null = null;

  // Prefer explicit id from bulk / callers. Same-request SharePoint cache can
  // miss a company that ensureCompany just created via createAdminCompany.
  // In bulkMode, trust id + name and skip listAdminCompanies entirely.
  if (providedCompanyId && isBulk) {
    company = {
      id: providedCompanyId,
      title: companyName,
      companyName,
      companyNumber: optionalText(input.companyNumber),
      companySize: null,
      registeredAddress: null,
      companyRegNumber: null,
      vatNo: null,
      telNo: null,
      email: null,
      mainContact: null,
      accountsContactName: null,
      accountsAddress: null,
      accountsContactNumber: null,
      accountsEmail: null,
      notesPricesAgreed: null,
      companyLogo: null,
      status: "Active",
    } satisfies Company;
  } else if (providedCompanyId) {
    const companies = await listAdminCompanies();
    company =
      companies.find((row) => row.id === providedCompanyId) ??
      ({
        id: providedCompanyId,
        title: companyName,
        companyName,
        companyNumber: optionalText(input.companyNumber),
        companySize: null,
        registeredAddress: null,
        companyRegNumber: null,
        vatNo: null,
        telNo: null,
        email: null,
        mainContact: null,
        accountsContactName: null,
        accountsAddress: null,
        accountsContactNumber: null,
        accountsEmail: null,
        notesPricesAgreed: null,
        companyLogo: null,
        status: "Active",
      } satisfies Company);
  }

  if (!company) {
    const companies = await listAdminCompanies();
    company =
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
  }

  if (!company) {
    throw new ValidationError(`Company "${companyName}" was not found.`);
  }

  let workforceNumber = providedWorkforceNumber;
  if (workforceNumber) {
    if (isBulk) {
      // In-memory clash check only — remote Graph filter per row made 50-row
      // imports take ~100s.
      const known = new Set(
        (
          Array.isArray(input.knownWorkforceNumbers)
            ? input.knownWorkforceNumbers
            : []
        )
          .map((value) => String(value ?? "").trim().toLowerCase())
          .filter(Boolean),
      );
      if (known.has(workforceNumber.toLowerCase())) {
        throw new ValidationError(
          `Workforce number "${workforceNumber}" is already in use.`,
        );
      }
    } else {
      const clashes = await getListItemsByKey("workforce", {
        filter: buildSchemaFieldEqualsFilter(
          "workforce",
          "workforceNumber",
          workforceNumber,
        ),
        top: 5,
      });
      if (clashes.length > 0) {
        throw new ValidationError(
          `Workforce number "${workforceNumber}" is already in use.`,
        );
      }
    }
  } else {
    const existingWorkforce = await listAdminWorkforce();
    workforceNumber = allocateNextWorkforceNumber(existingWorkforce);
  }

  // CompanyName is a Lookup — write LookupId, not free text.
  // Training manager / Supervisor / Department are Lookups (SharePoint scheme).
  const payload = toSharePointFields("workforce", {
    title: candidateName,
    candidateName,
    workforceNumber,
    dateOfBirth: asDateInput(input.dateOfBirth),
    status: normalizeWorkforceStatus(input.status, "Active"),
    email: optionalText(input.email),
    candidateAddress: optionalText(input.candidateAddress),
    contactNumber: optionalText(input.contactNumber),
    niNumber: optionalText(input.niNumber),
    cscsNumber: optionalText(input.cscsNumber),
    swqrNumber: optionalText(input.swqrNumber),
    eusrNumber: optionalText(input.eusrNumber),
    nporsNumbers: optionalText(input.nporsNumbers),
    inHouseCertificationNumber: optionalText(
      input.inHouseCertificationNumber,
    ),
    cscsExpiry: asDateInput(input.cscsExpiry),
    swqrExpiry: asDateInput(input.swqrExpiry),
    eusrExpiry: asDateInput(input.eusrExpiry),
    notes: optionalText(input.notes),
  });
  payload.CompanyNameLookupId = Number(company.id);

  const createIfMissing =
    input.createMissingPermissionPeople !== false &&
    input.createMissingPermissionPeople !== "false";
  const seededPeople = Array.isArray(input.permissionPeople)
    ? (input.permissionPeople as PermissionPersonRef[])
    : undefined;
  const seededDepartments = Array.isArray(input.departmentRecords)
    ? (input.departmentRecords as Array<{
        id: string;
        name: string;
        companyId?: string | null;
        companyName?: string | null;
      }>)
    : undefined;
  const personLookups = await applyWorkforcePersonLookups(payload, input, {
    companyId: company.id,
    createIfMissing,
    people: seededPeople,
  });
  const departmentLookups = await applyWorkforceDepartmentLookup(
    payload,
    input,
    {
      createIfMissing,
      companyId: company.id,
      companyName: company.companyName,
      departmentRecords: seededDepartments,
    },
  );

  const item = await createListItemByKey("workforce", payload);
  const mapped =
    mapWorkforce(
      item,
      new Map([[company.id, company.companyName]]),
      new Map([[company.id, company.companyNumber]]),
      permissionNameByIdMap(personLookups.people),
      departmentNameByIdMap(departmentLookups.departments),
    ) ??
    ({
      id: item.id,
      workforceNumber,
      candidateName,
      companyId: company.id,
      companyName: company.companyName,
      companyNumber: company.companyNumber,
      trainingManager:
        personLookups.trainingManagerName ??
        optionalText(input.trainingManager),
      supervisor:
        personLookups.supervisorName ?? optionalText(input.supervisor),
      candidateAddress: optionalText(input.candidateAddress),
      email: optionalText(input.email),
      contactNumber: optionalText(input.contactNumber),
      dateOfBirth: asNullableString(payload[workforceFields.dateOfBirth]),
      niNumber: optionalText(input.niNumber),
      nporsNumbers: optionalText(input.nporsNumbers),
      cscsNumber: optionalText(input.cscsNumber),
      cscsExpiry: asNullableString(
        asDateInput(input.cscsExpiry) ?? null,
      ),
      swqrNumber: optionalText(input.swqrNumber),
      swqrExpiry: asNullableString(
        asDateInput(input.swqrExpiry) ?? null,
      ),
      eusrNumber: optionalText(input.eusrNumber),
      eusrExpiry: asNullableString(
        asDateInput(input.eusrExpiry) ?? null,
      ),
      inHouseCertificationNumber: optionalText(
        input.inHouseCertificationNumber,
      ),
      department:
        departmentLookups.departmentName ??
        optionalText(input.departmentText) ??
        optionalText(input.department),
      status: normalizeWorkforceStatus(input.status, "Active"),
      notes: optionalText(input.notes),
      photoUrl: null,
    } satisfies AdminWorkforceRecord);

  const skipFolders =
    input.skipDocumentFolders === true ||
    input.skipDocumentFolders === "true" ||
    input.bulkMode === true ||
    input.bulkMode === "true";
  const skipMatrix =
    input.skipMatrixSeed === true ||
    input.skipMatrixSeed === "true" ||
    input.bulkMode === true ||
    input.bulkMode === "true";

  let folderWarning: string | undefined;
  if (!skipFolders) {
    const { ensureCandidateDocumentFolders } = await import(
      "@/lib/services/customerDocumentsFolderService"
    );
    const folderResult = await ensureCandidateDocumentFolders({
      companyName: company.companyName,
      companyNumber: company.companyNumber ?? null,
      candidateName,
      workforceNumber,
    });
    folderWarning = folderResult.ok ? undefined : folderResult.warning;
  }

  // Seed a Training Matrix row so the candidate appears from day one.
  // Skipped in bulkMode — matrix import / manual sync can fill expiry columns.
  let matrixSeedWarning: string | undefined;
  if (!skipMatrix) {
    try {
      const { upsertTrainingMatrixExampleRow } = await import(
        "@/lib/services/bulkUpload/trainingMatrixExampleService"
      );
      await upsertTrainingMatrixExampleRow({
        candidateName,
        source: {
          Name: candidateName,
          DOB: asNullableString(asDateInput(input.dateOfBirth) ?? null),
          "CSCS Expiry": asNullableString(asDateInput(input.cscsExpiry) ?? null),
          "EUSR Expiry": asNullableString(asDateInput(input.eusrExpiry) ?? null),
          "NRSWA Expiry": asNullableString(asDateInput(input.swqrExpiry) ?? null),
        },
      });
    } catch (error) {
      console.error(
        "[workforce] matrix seed failed (candidate still created):",
        error,
      );
      matrixSeedWarning =
        "Candidate saved, but Training Matrix row could not be created. Open Training Matrix and refresh, or contact support.";
    }
  }

  return {
    ...mapped,
    folderWarning,
    matrixSeedWarning,
    // Bulk importers reuse this to avoid reloading Permissions after every row.
    permissionPeople: personLookups.people,
    departmentRecords: departmentLookups.departmentRecords,
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
    status:
      input.status === undefined
        ? undefined
        : normalizeWorkforceStatus(input.status, null),
    email: input.email === undefined ? undefined : optionalText(input.email),
    candidateAddress:
      input.candidateAddress === undefined
        ? undefined
        : optionalText(input.candidateAddress),
    contactNumber:
      input.contactNumber === undefined
        ? undefined
        : optionalText(input.contactNumber),
    niNumber:
      input.niNumber === undefined ? undefined : optionalText(input.niNumber),
    cscsNumber:
      input.cscsNumber === undefined
        ? undefined
        : optionalText(input.cscsNumber),
    swqrNumber:
      input.swqrNumber === undefined
        ? undefined
        : optionalText(input.swqrNumber),
    eusrNumber:
      input.eusrNumber === undefined
        ? undefined
        : optionalText(input.eusrNumber),
    nporsNumbers:
      input.nporsNumbers === undefined
        ? undefined
        : optionalText(input.nporsNumbers),
    inHouseCertificationNumber:
      input.inHouseCertificationNumber === undefined
        ? undefined
        : optionalText(input.inHouseCertificationNumber),
    cscsExpiry:
      input.cscsExpiry === undefined
        ? undefined
        : asDateInput(input.cscsExpiry),
    swqrExpiry:
      input.swqrExpiry === undefined
        ? undefined
        : asDateInput(input.swqrExpiry),
    eusrExpiry:
      input.eusrExpiry === undefined
        ? undefined
        : asDateInput(input.eusrExpiry),
    notes: input.notes === undefined ? undefined : optionalText(input.notes),
  });

  let companyIdForPeople: string | null =
    extractLookupId(existing.fields, workforceFields.companyName) ?? null;

  const providedCompanyId = optionalText(input.companyId);
  if (providedCompanyId) {
    payload.CompanyNameLookupId = Number(providedCompanyId);
    companyIdForPeople = providedCompanyId;
  } else if (input.companyName !== undefined) {
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
      companyIdForPeople = company.id;
    }
  }

  const createIfMissing =
    input.createMissingPermissionPeople !== false &&
    input.createMissingPermissionPeople !== "false";
  const seededPeople = Array.isArray(input.permissionPeople)
    ? (input.permissionPeople as PermissionPersonRef[])
    : undefined;
  const seededDepartments = Array.isArray(input.departmentRecords)
    ? (input.departmentRecords as Array<{
        id: string;
        name: string;
        companyId?: string | null;
        companyName?: string | null;
      }>)
    : undefined;
  const personLookups = await applyWorkforcePersonLookups(payload, input, {
    companyId: companyIdForPeople,
    createIfMissing,
    people: seededPeople,
  });
  const departmentLookups = await applyWorkforceDepartmentLookup(
    payload,
    input,
    {
      createIfMissing,
      companyId: companyIdForPeople,
      companyName: optionalText(input.companyName),
      departmentRecords: seededDepartments,
    },
  );

  const item = await updateListItemFieldsByKey("workforce", id, payload);
  const companies = await listAdminCompanies();
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const companyNumberById = new Map(
    companies.map((row) => [row.id, row.companyNumber] as const),
  );
  const mapped = mapWorkforce(
    item,
    companyNameById,
    companyNumberById,
    permissionNameByIdMap(personLookups.people),
    departmentNameByIdMap(departmentLookups.departments),
  );
  if (!mapped) throw new Error("Updated candidate could not be mapped.");
  return mapped;
}

/**
 * Remove a Workforce candidate. Training register history is kept
 * (NPORS/EUSR/etc. rows stay; their Candidate lookup is cleared so SharePoint
 * Restrict Delete will allow removing the Workforce item). Matrix seed rows
 * for the same name are removed best-effort.
 */
export async function deleteAdminWorkforce(id: string): Promise<void> {
  const trimmed = String(id ?? "").trim();
  if (!trimmed) {
    throw new ValidationError("Candidate id is required.");
  }

  const existing = await getListItemByKey("workforce", trimmed);
  if (!existing) throw new NotFoundError("Candidate not found.");

  const candidateName =
    asLookupOrString(existing.fields[workforceFields.candidateName]) ??
    asString(existing.fields[workforceFields.candidateName]);

  // Clear inbound Restrict-Delete lookups before removing Workforce.
  const { clearInboundLookupsToWorkforce } = await import(
    "@/lib/services/adminSafeDelete"
  );
  await clearInboundLookupsToWorkforce(trimmed);

  // Training Matrix Update is Title=Name (often not a lookup) — remove matching seeds first.
  if (candidateName?.trim()) {
    try {
      const exampleRows = await listTrainingMatrixExampleRows();
      const key = candidateName.trim().toLowerCase();
      const matches = exampleRows.filter(
        (row) => row.candidateName.trim().toLowerCase() === key,
      );
      await Promise.all(
        matches.map((row) =>
          deleteListItemByKey("trainingMatrixExample", row.id).catch(() => null),
        ),
      );
    } catch (error) {
      console.warn(
        "[workforce] pre-delete matrix seed cleanup failed:",
        error instanceof Error ? error.message : error,
      );
    }
  }

  try {
    await deleteListItemByKey("workforce", trimmed);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const related =
      /related to another item|cannot be deleted because/i.test(message);
    throw new ValidationError(
      related
        ? "SharePoint still blocks deleting this candidate because another list references them. Cleared Candidate lookups on registers, documents, and matrix rows — if it keeps failing, ask a Site Owner which list has a lookup to Workforce with “Restrict delete”."
        : message.includes("SharePoint") || message.includes("delete")
          ? message
          : `Could not delete this candidate. ${message}`,
    );
  }
}

/* ───────────────── Training Matrix ───────────────── */

const matrixFields = getSharePointFields("trainingMatrix");

export interface AdminMatrixRecord {
  id: string;
  candidateName: string;
  companyName: string | null;
  department: string | null;
  dateOfBirth: string | null;
  overallStatus: string | null;
  needsReview: boolean;
  matrixNotes: string | null;
  nextExpiryDate: string | null;
  /** Optional — from Training Matrix Update “CSCS Expiry” column. */
  cscsExpiry?: string | null;
  ssstsExpiry?: string | null;
  smstsExpiry?: string | null;
  nrswaExpiry?: string | null;
  eusrExpiry?: string | null;
  n001Expiry: string | null;
  n003Expiry: string | null;
  n004Expiry: string | null;
  n010Expiry: string | null;
  n020Expiry: string | null;
  n021Expiry: string | null;
  n027Expiry: string | null;
  n100Expiry: string | null;
  /**
   * Values keyed by exact Training matrix example.xlsx headers
   * (Name, DOB, CSCS Expiry, N001 - Ind FLT, …).
   */
  columnValues: Record<string, string | null>;
  /** Headers manually set in admin — register sync will not overwrite these. */
  manualOverrideHeaders?: string[];
  n031Expiry?: string | null;
  /** SharePoint column is still "Face ift" — UI label is Face Fit. */
  faceFitExpiry?: string | null;
  /** Workforce List item id when the matrix Name matches a candidate. */
  workforceId?: string | null;
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
    dateOfBirth: workforceHit?.dateOfBirth ?? null,
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
    columnValues: {},
  };
}

function asDateOnly(value: string | null | undefined): string | null {
  if (!value?.trim()) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10);
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value.trim();
  return parsed.toISOString().slice(0, 10);
}

function matrixRowKey(candidate: string, company: string | null): string {
  return `${(candidate ?? "").trim().toLowerCase()}|${(company ?? "")
    .trim()
    .toLowerCase()}`;
}

async function loadMatrixCategoryExpiryIndex(): Promise<
  Map<string, Map<string, string>>
> {
  const index = new Map<string, Map<string, string>>();
  try {
    const [items, categoryItems] = await Promise.all([
      getListItemsByKey("trainingMatrixCategoryRecords", { top: 5000 }),
      getListItemsByKey("nporsCategories", { top: 5000 }).catch(() => []),
    ]);
    const catFields = getSharePointFields("trainingMatrixCategoryRecords");
    const nporsFields = getSharePointFields("nporsCategories");
    const codeByCategoryId = new Map<string, string>();
    for (const item of categoryItems) {
      const title =
        asNullableString(item.fields[nporsFields.title]) ??
        asNullableString(item.fields[nporsFields.categoryCode]);
      const code = title
        ? title.match(/^(N\d+[A-Z]?|CSCS|SSSTS|SMSTS|NRSWA|EUSR|FACEFIT)/i)?.[1]
            ?.toUpperCase() ?? null
        : null;
      if (code) codeByCategoryId.set(item.id, code);
    }

    const workforce = await listAdminWorkforce();
    const workforceById = new Map(
      workforce.map((row) => [row.id, row] as const),
    );
    const companies = await listAdminCompanies();
    const companyById = new Map(
      companies.map((row) => [row.id, row.companyName] as const),
    );

    for (const item of items) {
      const candidateLookupId = extractLookupId(
        item.fields,
        catFields.candidateName,
      );
      const companyLookupId = extractLookupId(
        item.fields,
        catFields.companyName,
      );
      const categoryLookupId = extractLookupId(
        item.fields,
        catFields.categoryCode,
      );
      const candidate =
        asLookupOrString(item.fields[catFields.candidateName]) ??
        (candidateLookupId
          ? (workforceById.get(candidateLookupId)?.candidateName ?? null)
          : null);
      const company =
        asLookupOrString(item.fields[catFields.companyName]) ??
        (companyLookupId
          ? (companyById.get(companyLookupId) ?? null)
          : null);
      const code =
        (categoryLookupId
          ? (codeByCategoryId.get(categoryLookupId) ?? null)
          : null) ??
        extractCategoryCodeFromValue(
          asLookupOrString(item.fields[catFields.categoryCode]) ??
            asNullableString(item.fields[catFields.categoryCode]) ??
            asNullableString(item.fields[catFields.categoryName]),
        );
      const expiry = asDateOnly(
        asNullableString(item.fields[catFields.expiryDate]),
      );
      if (!candidate || !code || !expiry) continue;
      const key = matrixRowKey(candidate, company);
      const bucket = index.get(key) ?? new Map<string, string>();
      bucket.set(code.trim().toUpperCase(), expiry);
      index.set(key, bucket);
    }
  } catch {
    // Category list may be unset in some environments — matrix page still works.
  }
  return index;
}

function extractCategoryCodeFromValue(
  value: string | null | undefined,
): string | null {
  if (!value?.trim()) return null;
  const text = value.trim();
  const meta = ["CSCS", "SSSTS", "SMSTS", "NRSWA", "EUSR", "FACEFIT"];
  const upper = text.toUpperCase();
  for (const code of meta) {
    if (
      upper === code ||
      upper.startsWith(`${code} `) ||
      upper.startsWith(`${code}-`)
    ) {
      return code;
    }
  }
  return text.match(/^(N\d+[A-Z]?)/i)?.[1]?.toUpperCase() ?? null;
}

function buildMatrixColumnValues(
  row: AdminMatrixRecord,
  byCode: Map<string, string> | undefined,
  workforce?: AdminWorkforceRecord | null,
): Record<string, string | null> {
  const values: Record<string, string | null> = {
    Name: row.candidateName,
    DOB: asDateOnly(row.dateOfBirth ?? workforce?.dateOfBirth),
    "CSCS Expiry":
      byCode?.get("CSCS") ?? asDateOnly(workforce?.cscsExpiry) ?? null,
    "SSSTS Expiry": byCode?.get("SSSTS") ?? null,
    "SMSTS Expiry": byCode?.get("SMSTS") ?? null,
    "NRSWA Expiry": byCode?.get("NRSWA") ?? null,
    "EUSR Expiry":
      byCode?.get("EUSR") ?? asDateOnly(workforce?.eusrExpiry) ?? null,
    "Face ift": byCode?.get("FACEFIT") ?? null,
  };

  const matrixFieldByCode: Record<string, string | null> = {
    N001: asDateOnly(row.n001Expiry),
    N003: asDateOnly(row.n003Expiry),
    N004: asDateOnly(row.n004Expiry),
    N010: asDateOnly(row.n010Expiry),
    N020: asDateOnly(row.n020Expiry),
    N021: asDateOnly(row.n021Expiry),
    N027: asDateOnly(row.n027Expiry),
    N100: asDateOnly(row.n100Expiry),
  };

  for (const column of CLIENT_MATRIX_CATEGORY_COLUMNS) {
    values[column.header] =
      byCode?.get(column.code.toUpperCase()) ??
      matrixFieldByCode[column.code.toUpperCase()] ??
      null;
  }

  return values;
}

export async function listAdminMatrix(companyName?: string | null) {
  // Portal matrix UI + register sync both use SharePoint "Training Matrix Update".
  const [workforce, exampleRows] = await Promise.all([
    listAdminWorkforce(),
    listTrainingMatrixExampleRows(),
  ]);

  const companyKey = companyName?.trim().toLowerCase() || "";
  const workforceForCompany = companyKey
    ? workforce.filter(
        (row) => row.companyName.trim().toLowerCase() === companyKey,
      )
    : workforce;

  const workforceByName = new Map<string, AdminWorkforceRecord>();
  for (const row of workforceForCompany) {
    const key = row.candidateName.trim().toLowerCase();
    if (key && !workforceByName.has(key)) workforceByName.set(key, row);
  }

  // When scoping to a company, only keep matrix rows whose candidate is in
  // that company's workforce (Training Matrix Update has no company column).
  const companyNameKeys = companyKey
    ? new Set(workforceByName.keys())
    : null;

  return exampleRows
    .filter((example) => {
      if (!companyNameKeys) return true;
      const key = example.candidateName.trim().toLowerCase();
      return Boolean(key && companyNameKeys.has(key));
    })
    .map((example) => {
      const wf =
        workforceByName.get(example.candidateName.trim().toLowerCase()) ?? null;
      const columnValues = { ...example.columnValues };
      // Prefer Workforce card expiries when matrix cell is blank.
      if (!columnValues["CSCS Expiry"]?.trim() && wf?.cscsExpiry) {
        columnValues["CSCS Expiry"] = wf.cscsExpiry;
      }
      if (!columnValues["EUSR Expiry"]?.trim() && wf?.eusrExpiry) {
        columnValues["EUSR Expiry"] = wf.eusrExpiry;
      }
      if (!columnValues["NRSWA Expiry"]?.trim() && wf?.swqrExpiry) {
        columnValues["NRSWA Expiry"] = wf.swqrExpiry;
      }
      if (!columnValues.DOB?.trim() && wf?.dateOfBirth) {
        columnValues.DOB = wf.dateOfBirth;
      }
      const nextExpiryDate =
        example.nextExpiryDate ?? earliestDateFromColumns(columnValues);

      const record: AdminMatrixRecord = {
        id: `example:${example.id}`,
        candidateName: example.candidateName,
        companyName: wf?.companyName ?? (companyName?.trim() || null),
        department: wf?.department ?? null,
        dateOfBirth: example.dateOfBirth ?? wf?.dateOfBirth ?? null,
        overallStatus: null,
        needsReview: !nextExpiryDate,
        matrixNotes: null,
        nextExpiryDate,
        cscsExpiry: columnValues["CSCS Expiry"] ?? null,
        ssstsExpiry: columnValues["SSSTS Expiry"] ?? null,
        smstsExpiry: columnValues["SMSTS Expiry"] ?? null,
        nrswaExpiry: columnValues["NRSWA Expiry"] ?? null,
        eusrExpiry: columnValues["EUSR Expiry"] ?? null,
        n001Expiry: columnValues["N001 - Ind FLT"] ?? null,
        n003Expiry: columnValues["N003 - Reach Lift Truck"] ?? null,
        n004Expiry: columnValues["N004 - Lorry Mounted Lift Truck"] ?? null,
        n010Expiry: columnValues["N010 - Telescopic Handler"] ?? null,
        n020Expiry: columnValues["N020 - Tiltrotator System"] ?? null,
        n021Expiry: columnValues["N021 - Suction Excavator"] ?? null,
        n027Expiry:
          columnValues["N027 - Excavation Marshal - Banksperson"] ?? null,
        n100Expiry: columnValues["N100 - Exc Crane"] ?? null,
        n031Expiry: columnValues["N031 - Asbestos Awareness"] ?? null,
        faceFitExpiry: columnValues["Face ift"] ?? null,
        columnValues,
        manualOverrideHeaders: example.manualOverrides ?? [],
        workforceId: wf?.id ?? null,
      };
      return record;
    })
    .concat(
      // Workforce candidates with no matrix row yet still appear (seed may have failed).
      workforceForCompany
        .filter((wf) => {
          const key = wf.candidateName.trim().toLowerCase();
          return (
            Boolean(key) &&
            !exampleRows.some(
              (example) =>
                example.candidateName.trim().toLowerCase() === key,
            )
          );
        })
        .map((wf) => {
          const columnValues: Record<string, string | null> = {
            Name: wf.candidateName,
            DOB: wf.dateOfBirth,
            "CSCS Expiry": wf.cscsExpiry,
            "EUSR Expiry": wf.eusrExpiry,
            "NRSWA Expiry": wf.swqrExpiry,
          };
          const nextExpiryDate = earliestDateFromColumns(columnValues);
          const record: AdminMatrixRecord = {
            id: `workforce-only:${wf.id}`,
            candidateName: wf.candidateName,
            companyName: wf.companyName,
            department: wf.department,
            dateOfBirth: wf.dateOfBirth,
            overallStatus: null,
            needsReview: true,
            matrixNotes: null,
            nextExpiryDate,
            cscsExpiry: wf.cscsExpiry,
            ssstsExpiry: null,
            smstsExpiry: null,
            nrswaExpiry: wf.swqrExpiry,
            eusrExpiry: wf.eusrExpiry,
            n001Expiry: null,
            n003Expiry: null,
            n004Expiry: null,
            n010Expiry: null,
            n020Expiry: null,
            n021Expiry: null,
            n027Expiry: null,
            n100Expiry: null,
            n031Expiry: null,
            faceFitExpiry: null,
            columnValues,
            manualOverrideHeaders: [],
            workforceId: wf.id,
          };
          return record;
        }),
    );
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
      dateOfBirth: candidate.dateOfBirth,
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
      columnValues: {},
    } satisfies AdminMatrixRecord);
  return {
    ...mapped,
    dateOfBirth: mapped.dateOfBirth ?? candidate.dateOfBirth,
    columnValues: buildMatrixColumnValues(mapped, undefined, candidate),
  };
}

export async function deleteAdminMatrix(id: string) {
  if (id.startsWith("example:")) {
    const exampleId = stripExampleMatrixId(id);
    if (!exampleId) throw new NotFoundError("Matrix record not found.");
    await deleteListItemByKey("trainingMatrixExample", exampleId);
    return;
  }
  if (id.startsWith("workforce-only:")) {
    throw new ValidationError(
      "This candidate has no matrix row to delete yet.",
    );
  }

  const existing = await getListItemByKey("trainingMatrix", id);
  if (!existing) throw new NotFoundError("Matrix record not found.");
  await deleteListItemByKey("trainingMatrix", id);
}

export async function deleteAdminRegister(
  key: AdminRegisterKey,
  id: string,
): Promise<void> {
  const existing = await getListItemByKey(key, id);
  if (!existing) throw new NotFoundError("Training record not found.");
  await deleteListItemByKey(key, id);
}

export async function deleteAdminNvq(id: string): Promise<void> {
  const existing = await getListItemByKey("nvqRegister", id);
  if (!existing) throw new NotFoundError("NVQ record not found.");
  await deleteListItemByKey("nvqRegister", id);
}

export async function deleteAdminOffer(id: string): Promise<void> {
  const existing = await getListItemByKey("offersPromotions", id);
  if (!existing) throw new NotFoundError("Offer not found.");
  await deleteListItemByKey("offersPromotions", id);
}

export async function deleteAdminDocument(id: string): Promise<void> {
  const existing = await getListItemByKey("customerDocuments", id);
  if (!existing) throw new NotFoundError("Document not found.");
  await deleteListItemByKey("customerDocuments", id);
}

export async function updateAdminMatrix(
  id: string,
  input: Record<string, unknown>,
) {
  // First edit for a Workforce candidate that has no matrix row yet — create it.
  if (id.startsWith("workforce-only:")) {
    const workforceId = id.slice("workforce-only:".length).trim();
    if (!workforceId) throw new NotFoundError("Matrix record not found.");
    const workforce = await listAdminWorkforce();
    const wf = workforce.find((row) => row.id === workforceId);
    if (!wf) throw new NotFoundError("Candidate not found.");

    const seedSource: Record<string, string | null> = {
      Name: wf.candidateName,
      DOB: wf.dateOfBirth,
      "CSCS Expiry": wf.cscsExpiry,
      "EUSR Expiry": wf.eusrExpiry,
      "NRSWA Expiry": wf.swqrExpiry,
    };
    const created = await upsertTrainingMatrixExampleRow({
      candidateName: wf.candidateName,
      source: seedSource,
    });
    return updateAdminMatrix(`example:${created.id}`, input);
  }

  if (id.startsWith("example:")) {
    const exampleId = stripExampleMatrixId(id);
    if (!exampleId) throw new NotFoundError("Matrix record not found.");

    const existingRows = await listTrainingMatrixExampleRows();
    const existing = existingRows.find((row) => row.id === exampleId);
    if (!existing) throw new NotFoundError("Matrix record not found.");

    const source: Record<string, string | null> = {
      ...existing.columnValues,
      Name: existing.candidateName,
      DOB: existing.dateOfBirth,
    };

    if (input.candidateName !== undefined || input.Name !== undefined) {
      source.Name =
        optionalText(input.candidateName) ??
        optionalText(input.Name) ??
        existing.candidateName;
    }
    if (input.dateOfBirth !== undefined || input.DOB !== undefined) {
      source.DOB =
        asDateInput(input.dateOfBirth) ??
        asDateInput(input.DOB) ??
        null;
    }

    const namedDateFields: Array<[string, string]> = [
      ["cscsExpiry", "CSCS Expiry"],
      ["ssstsExpiry", "SSSTS Expiry"],
      ["smstsExpiry", "SMSTS Expiry"],
      ["nrswaExpiry", "NRSWA Expiry"],
      ["eusrExpiry", "EUSR Expiry"],
      ["faceFitExpiry", "Face ift"],
      ["n001Expiry", "N001 - Ind FLT"],
      ["n003Expiry", "N003 - Reach Lift Truck"],
      ["n004Expiry", "N004 - Lorry Mounted Lift Truck"],
      ["n010Expiry", "N010 - Telescopic Handler"],
      ["n020Expiry", "N020 - Tiltrotator System"],
      ["n021Expiry", "N021 - Suction Excavator"],
      ["n027Expiry", "N027 - Excavation Marshal - Banksperson"],
      ["n100Expiry", "N100 - Exc Crane"],
      ["n031Expiry", "N031 - Asbestos Awareness"],
    ];
    const changedHeaders: string[] = [];
    for (const [field, header] of namedDateFields) {
      if (input[field] !== undefined) {
        source[header] = asDateInput(input[field]) ?? null;
        changedHeaders.push(header);
      }
    }

    if (input.columnValues && typeof input.columnValues === "object") {
      for (const [key, value] of Object.entries(
        input.columnValues as Record<string, unknown>,
      )) {
        if (value === undefined) continue;
        source[key] =
          value == null || value === ""
            ? null
            : asDateInput(value) ?? String(value);
        changedHeaders.push(key);
      }
    }

    // Allow direct header keys from the edit form (e.g. "N006 - Side Loader").
    for (const [key, value] of Object.entries(input)) {
      if (
        key === "candidateName" ||
        key === "companyName" ||
        key === "department" ||
        key === "columnValues" ||
        key === "id" ||
        key === "manualOverrideHeaders"
      ) {
        continue;
      }
      if (CLIENT_MATRIX_DISPLAY_HEADERS.includes(key as never)) {
        source[key] =
          value == null || value === ""
            ? null
            : asDateInput(value) ?? String(value);
        changedHeaders.push(key);
      }
    }

    const name = source.Name?.trim() || existing.candidateName;
    if (!name) {
      throw new ValidationError("Candidate name is required.");
    }

    const { mergeManualOverrides } = await import(
      "@/lib/training/matrixManualOverrides"
    );
    const manualOverrides = mergeManualOverrides(
      existing.manualOverrides ?? [],
      changedHeaders.filter(
        (h) => h !== "Name" && h !== "DOB" && CLIENT_MATRIX_DISPLAY_HEADERS.includes(h as never),
      ),
    );

    await upsertTrainingMatrixExampleRow({
      candidateName: name,
      existingItemId: exampleId,
      source: { ...source, Name: name },
      manualOverrides,
    });

    const rows = await listAdminMatrix();
    const updated = rows.find((row) => row.id === `example:${exampleId}`);
    if (!updated) {
      throw new NotFoundError("Matrix record not found after update.");
    }
    return updated;
  }

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
  const wf =
    (mapped.companyName
      ? workforce.find(
          (row) =>
            matrixRowKey(row.candidateName, row.companyName) ===
            matrixRowKey(mapped.candidateName, mapped.companyName),
        )
      : null) ??
    workforce.find(
      (row) =>
        row.candidateName.trim().toLowerCase() ===
        mapped.candidateName.trim().toLowerCase(),
    ) ??
    null;
  const expiryByPerson = await loadMatrixCategoryExpiryIndex();
  const expiryByCode = expiryByPerson.get(
    matrixRowKey(mapped.candidateName, mapped.companyName),
  );
  return {
    ...mapped,
    dateOfBirth: mapped.dateOfBirth ?? wf?.dateOfBirth ?? null,
    columnValues: buildMatrixColumnValues(mapped, expiryByCode, wf),
  };
}

export type AdminRegisterKey =
  | "nporsRegister"
  | "eusrRegister"
  | "nrswaRegister"
  | "inHouseCertificates";

export interface AdminTrainingRecord {
  id: string;
  candidateName: string;
  companyName: string;
  /** From Workforce when candidate lookup resolves. */
  workforceNumber?: string | null;
  trainingDate: string | null;
  trainingAddress: string | null;
  trainingOutcome: string | null;
  outcomeDate?: string | null;
  assessorTrainer?: string | null;
  outcomeNotes?: string | null;
  notes?: string | null;
  customerVisible: boolean;
  expiry: string | null;
  // register-specific
  nporsNumber?: string | null;
  noviceOrEwt?: string | null;
  nporsCategory?: string | null;
  eusrNumber?: string | null;
  eusrCategory?: string | null;
  cardType?: string | null;
  swqrNumber?: string | null;
  course?: string | null;
  streetworksCategory?: string | null;
  /** Streetworks multi-day end date (stored in OutcomeNotes until SP has a column). */
  trainingDateEnd?: string | null;
  certificateCategory?: string | null;
  courseCategory?: string | null;
  inHouseCertificationNumber?: string | null;
}

type RegisterLookupMaps = {
  companyNameById: Map<string, string>;
  companyNumberById: Map<string, string>;
  workforceNameById: Map<string, string>;
  workforceCompanyById: Map<string, string>;
  workforceNumberById: Map<string, string>;
};

async function loadRegisterLookupMaps(): Promise<RegisterLookupMaps> {
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const companyNumberById = new Map(
    companies
      .filter((row) => row.companyNumber?.trim())
      .map((row) => [row.id, row.companyNumber as string] as const),
  );
  const workforceNameById = new Map(
    workforce.map((row) => [row.id, row.candidateName] as const),
  );
  const workforceCompanyById = new Map(
    workforce.map((row) => [row.id, row.companyName] as const),
  );
  const workforceNumberById = new Map(
    workforce
      .filter((row) => row.workforceNumber?.trim())
      .map((row) => [row.id, row.workforceNumber as string] as const),
  );
  return {
    companyNameById,
    companyNumberById,
    workforceNameById,
    workforceCompanyById,
    workforceNumberById,
  };
}

function resolveRegisterPeople(
  fields: SharePointFields,
  candidateField: string,
  companyField: string,
  lookups: RegisterLookupMaps,
): {
  candidateName: string;
  companyName: string;
  workforceNumber: string | null;
} | null {
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
  const workforceNumber = candidateLookupId
    ? (lookups.workforceNumberById.get(candidateLookupId) ?? null)
    : null;
  return { candidateName, companyName, workforceNumber };
}

/** Split admin form multi-choice text into a SharePoint MultiChoice array. */
function asMultiChoiceWrite(input: unknown): string[] | null | undefined {
  if (input === undefined) return undefined;
  const text = optionalText(input);
  if (!text) return null;
  const parts = text
    .split(/[;,|]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  return parts.length ? parts : null;
}

const TRAINING_END_MARKER = /\[TrainingEnd:(\d{4}-\d{2}-\d{2})\]\s*/i;

function extractTrainingDateEnd(notes: string | null): {
  trainingDateEnd: string | null;
  outcomeNotes: string | null;
} {
  if (!notes?.trim()) {
    return { trainingDateEnd: null, outcomeNotes: notes };
  }
  const match = notes.match(TRAINING_END_MARKER);
  if (!match) {
    return { trainingDateEnd: null, outcomeNotes: notes };
  }
  const cleaned = notes.replace(TRAINING_END_MARKER, "").trim();
  return {
    trainingDateEnd: match[1] ?? null,
    outcomeNotes: cleaned || null,
  };
}

function embedTrainingDateEnd(
  notes: string | null | undefined,
  trainingDateEnd: unknown,
): string | null | undefined {
  if (notes === undefined && trainingDateEnd === undefined) {
    return undefined;
  }
  const endText =
    trainingDateEnd === undefined
      ? undefined
      : optionalText(trainingDateEnd);
  const base =
    notes === undefined
      ? undefined
      : notes === null
        ? null
        : String(notes).replace(TRAINING_END_MARKER, "").trim() || null;

  if (endText === undefined) {
    return base;
  }
  if (!endText) {
    return base ?? null;
  }
  const marker = `[TrainingEnd:${endText}]`;
  if (!base) {
    return marker;
  }
  return `${marker} ${base}`;
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
      outcomeDate: asNullableString(item.fields[f.outcomeDate]),
      assessorTrainer: asNullableString(item.fields[f.assessorTrainer]),
      outcomeNotes: asNullableString(item.fields[f.outcomeNotes]),
      notes: asNullableString(item.fields[f.notes]),
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
      outcomeDate: asNullableString(item.fields[f.outcomeDate]),
      assessorTrainer: asNullableString(item.fields[f.assessorTrainer]),
      outcomeNotes: asNullableString(item.fields[f.outcomeNotes]),
      notes: asNullableString(item.fields[f.notes]),
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      expiry: asNullableString(item.fields[f.expiry]),
      eusrNumber: asNullableString(item.fields[f.eusrNumber]),
      eusrCategory: asMultiChoiceText(item.fields[f.eusrCategory]),
      cardType: asNullableString(item.fields[f.cardType]),
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
    const rawNotes = asNullableString(item.fields[f.outcomeNotes]);
    const { trainingDateEnd, outcomeNotes } = extractTrainingDateEnd(rawNotes);
    return {
      id: item.id,
      ...people,
      trainingDate: asNullableString(item.fields[f.trainingDate]),
      trainingAddress: asNullableString(item.fields[f.trainingAddress]),
      trainingOutcome: asNullableString(item.fields[f.trainingOutcome]),
      outcomeDate: asNullableString(item.fields[f.outcomeDate]),
      assessorTrainer: asNullableString(item.fields[f.assessorTrainer]),
      outcomeNotes,
      customerVisible: asBoolean(item.fields[f.customerVisible]),
      expiry: asNullableString(item.fields[f.expiryDate]),
      swqrNumber: asNullableString(item.fields[f.swqrNumber]),
      course: asNullableString(item.fields[f.course]),
      streetworksCategory: asMultiChoiceText(
        item.fields[f.streetworksCategory],
      ),
      trainingDateEnd,
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
    outcomeDate: asNullableString(item.fields[f.outcomeDate]),
    assessorTrainer: asNullableString(item.fields[f.assessorTrainer]),
    outcomeNotes: asNullableString(item.fields[f.outcomeNotes]),
    notes: asNullableString(item.fields[f.notes]),
    customerVisible: asBoolean(item.fields[f.customerVisible]),
    expiry: asNullableString(item.fields[f.expiryDate]),
    course: asNullableString(item.fields[f.courseCategory]),
    certificateCategory: asNullableString(item.fields[f.certificateCategory]),
    courseCategory: asNullableString(item.fields[f.courseCategory]),
    inHouseCertificationNumber:
      asLookupOrString(
        item.fields.Candidate_x0020_Name_x003a__x002,
      ) ?? null,
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
          : (() => {
              const text = optionalText(input.nporsCategory);
              if (!text) return null;
              // SharePoint NPORS Category is multi-choice (checkboxes).
              return text
                .split(/[;,|]+/)
                .map((part) => part.trim())
                .filter(Boolean)
                .map((part) => {
                  const upper = part.toUpperCase();
                  const match = upper.match(/\bN\d{3}\b/);
                  return match?.[0] ?? part;
                });
            })(),
      trainingOutcome: normalizedOutcome,
      outcomeDate:
        input.outcomeDate === undefined
          ? undefined
          : asDateInput(input.outcomeDate),
      assessorTrainer:
        input.assessorTrainer === undefined
          ? undefined
          : optionalText(input.assessorTrainer),
      outcomeNotes:
        input.outcomeNotes === undefined
          ? undefined
          : optionalText(input.outcomeNotes),
      notes: input.notes === undefined ? undefined : optionalText(input.notes),
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
      // eusrNumber is a projected Workforce lookup — never write it.
      eusrCategory:
        input.eusrCategory === undefined
          ? undefined
          : asMultiChoiceWrite(input.eusrCategory),
      cardType:
        input.cardType === undefined
          ? undefined
          : optionalText(input.cardType),
      trainingDate:
        input.trainingDate === undefined
          ? undefined
          : asDateInput(input.trainingDate),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      trainingOutcome: normalizedOutcome,
      outcomeDate:
        input.outcomeDate === undefined
          ? undefined
          : asDateInput(input.outcomeDate),
      assessorTrainer:
        input.assessorTrainer === undefined
          ? undefined
          : optionalText(input.assessorTrainer),
      outcomeNotes:
        input.outcomeNotes === undefined
          ? undefined
          : optionalText(input.outcomeNotes),
      notes: input.notes === undefined ? undefined : optionalText(input.notes),
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
      // swqrNumber is a projected Workforce lookup — never write it.
      course:
        input.course === undefined ? undefined : optionalText(input.course),
      streetworksCategory:
        input.streetworksCategory === undefined
          ? undefined
          : asMultiChoiceWrite(input.streetworksCategory),
      trainingDate:
        input.trainingDate === undefined
          ? undefined
          : asDateInput(input.trainingDate),
      trainingAddress:
        input.trainingAddress === undefined
          ? undefined
          : optionalText(input.trainingAddress),
      trainingOutcome: normalizedOutcome,
      outcomeDate:
        input.outcomeDate === undefined
          ? undefined
          : asDateInput(input.outcomeDate),
      assessorTrainer:
        input.assessorTrainer === undefined
          ? undefined
          : optionalText(input.assessorTrainer),
      outcomeNotes:
        input.outcomeNotes === undefined && input.trainingDateEnd === undefined
          ? undefined
          : embedTrainingDateEnd(
              input.outcomeNotes === undefined
                ? undefined
                : optionalText(input.outcomeNotes),
              input.trainingDateEnd,
            ),
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
    // Live In-House list stores the course on CertificateCategory (CourseCategory
    // is often missing). Accept Course / CourseCategory / Certificate Category.
    certificateCategory:
      input.certificateCategory === undefined &&
      input.course === undefined &&
      input.courseCategory === undefined
        ? undefined
        : optionalText(
            input.certificateCategory ?? input.course ?? input.courseCategory,
          ),
    courseDate:
      input.trainingDate === undefined
        ? undefined
        : asDateInput(input.trainingDate),
    trainingAddress:
      input.trainingAddress === undefined
        ? undefined
        : optionalText(input.trainingAddress),
    trainingOutcome: normalizedOutcome,
    outcomeDate:
      input.outcomeDate === undefined
        ? undefined
        : asDateInput(input.outcomeDate),
    assessorTrainer:
      input.assessorTrainer === undefined
        ? undefined
        : optionalText(input.assessorTrainer),
    outcomeNotes:
      input.outcomeNotes === undefined
        ? undefined
        : optionalText(input.outcomeNotes),
    notes: input.notes === undefined ? undefined : optionalText(input.notes),
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
  const workforceId = optionalText(input.workforceId);
  const candidateName =
    mode === "create"
      ? requireText(input.candidateName, "Candidate name")
      : optionalText(input.candidateName);
  const companyName =
    mode === "create"
      ? requireText(input.companyName, "Company")
      : optionalText(input.companyName);

  if (!workforceId && !candidateName && !companyName) return;

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
    (workforceId
      ? (workforce.find((row) => row.id === workforceId) ?? null)
      : null) ??
    (candidateName
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
      : null);

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

/**
 * Choice / MultiChoice columns often reject app writes when SP options are
 * out of date. Create the row without them, then retry those fields alone.
 */
function stripDeferredChoiceFields(
  key: AdminRegisterKey,
  payload: SharePointFields,
): SharePointFields {
  const deferred: SharePointFields = {};
  // Internal SharePoint names (from schema), not app field keys.
  const fieldsByKey: Partial<Record<AdminRegisterKey, string[]>> = {
    nporsRegister: ["NPORSCategory"],
    eusrRegister: ["EusrCategory", "CardType"],
    nrswaRegister: ["Course", "StreetworksCategory"],
    inHouseCertificates: ["CertificateCategory"],
  };
  for (const field of fieldsByKey[key] ?? []) {
    if (payload[field] !== undefined) {
      deferred[field] = payload[field];
      delete payload[field];
    }
  }
  return deferred;
}

async function applyDeferredChoiceFields(
  key: AdminRegisterKey,
  itemId: string,
  deferred: SharePointFields,
): Promise<string[]> {
  const failures: string[] = [];
  for (const [field, value] of Object.entries(deferred)) {
    if (value === undefined) continue;
    try {
      await updateListItemFieldsByKey(key, itemId, { [field]: value });
    } catch (error) {
      console.warn(
        `[${key}] ${field} write failed for #${itemId}; row saved without that choice (update SharePoint choices).`,
        error,
      );
      failures.push(
        `${field} could not be saved — update SharePoint choice options (Site Owner pack), then edit this record.`,
      );
    }
  }
  return failures;
}

export async function createAdminRegister(
  key: AdminRegisterKey,
  input: Record<string, unknown>,
): Promise<{ record: AdminTrainingRecord; choiceWarnings: string[] }> {
  const payload = registerWritePayload(key, input, "create");
  const deferred = stripDeferredChoiceFields(key, payload);

  await applyRegisterLookupIds(payload, input, "create");
  let item;
  try {
    item = await createListItemByKey(key, payload);
  } catch (error) {
    const detail =
      error instanceof Error ? error.message : "SharePoint create failed.";
    throw new ValidationError(
      `Could not save ${key === "nrswaRegister" ? "Streetworks" : key === "eusrRegister" ? "EUSR" : "training"} record: ${detail}`,
    );
  }

  const choiceWarnings = await applyDeferredChoiceFields(
    key,
    item.id,
    deferred,
  );
  if (Object.keys(deferred).length > 0) {
    item = (await getListItemByKey(key, item.id)) ?? item;
    // Keep submitted choice values for matrix sync even if SP rejected them.
    for (const [field, value] of Object.entries(deferred)) {
      item.fields = { ...item.fields, [field]: value };
    }
  }

  const lookups = await loadRegisterLookupMaps();
  const mapped = mapRegister(key, item, lookups);
  if (!mapped) throw new Error("Created training record could not be mapped.");

  const submittedCategory = optionalText(input.nporsCategory);
  if (submittedCategory && !mapped.nporsCategory) {
    mapped.nporsCategory = submittedCategory;
  }
  const submittedCourse = optionalText(input.course);
  if (submittedCourse && !mapped.course) {
    mapped.course = submittedCourse;
  }
  if (submittedCourse && !mapped.courseCategory) {
    mapped.courseCategory = submittedCourse;
  }
  const submittedCert =
    optionalText(input.certificateCategory) ?? submittedCourse;
  if (submittedCert && !mapped.certificateCategory) {
    mapped.certificateCategory = submittedCert;
  }
  const submittedEusr = optionalText(input.eusrCategory);
  if (submittedEusr && !mapped.eusrCategory) {
    mapped.eusrCategory = submittedEusr;
  }
  return { record: mapped, choiceWarnings };
}

export async function updateAdminRegister(
  key: AdminRegisterKey,
  id: string,
  input: Record<string, unknown>,
): Promise<{ record: AdminTrainingRecord; choiceWarnings: string[] }> {
  const existing = await getListItemByKey(key, id);
  if (!existing) throw new NotFoundError("Training record not found.");
  const payload = registerWritePayload(key, input, "update");
  const deferred = stripDeferredChoiceFields(key, payload);
  await applyRegisterLookupIds(payload, input, "update");
  let item;
  try {
    item = await updateListItemFieldsByKey(key, id, payload);
  } catch (error) {
    const detail =
      error instanceof Error ? error.message : "SharePoint update failed.";
    throw new ValidationError(
      `Could not update training record: ${detail}`,
    );
  }

  const choiceWarnings = await applyDeferredChoiceFields(key, id, deferred);
  if (Object.keys(deferred).length > 0) {
    item = (await getListItemByKey(key, id)) ?? item;
    for (const [field, value] of Object.entries(deferred)) {
      item.fields = { ...item.fields, [field]: value };
    }
  }

  const lookups = await loadRegisterLookupMaps();
  const mapped = mapRegister(key, item, lookups);
  if (!mapped) throw new Error("Updated training record could not be mapped.");
  const submittedCategory = optionalText(input.nporsCategory);
  if (submittedCategory) {
    mapped.nporsCategory = submittedCategory;
  }
  const submittedCourse = optionalText(input.course);
  if (submittedCourse) {
    mapped.course = submittedCourse;
    if (!mapped.courseCategory) mapped.courseCategory = submittedCourse;
  }
  const submittedEusr = optionalText(input.eusrCategory);
  if (submittedEusr) {
    mapped.eusrCategory = submittedEusr;
  }
  const submittedCert =
    optionalText(input.certificateCategory) ?? submittedCourse;
  if (submittedCert) {
    mapped.certificateCategory = submittedCert;
  }
  return { record: mapped, choiceWarnings };
}

/* ───────────────── NVQ ───────────────── */

const nvqFields = getSharePointFields("nvqRegister");

export interface AdminNvqRecord {
  id: string;
  candidateName: string;
  companyName: string | null;
  companyNumber: string | null;
  workforceNumber: string | null;
  niNumber: string | null;
  ulnNumber: string | null;
  nvqTitle: string | null;
  boltOn: string | null;
  poNumber: string | null;
  cardSchemeCategory: string | null;
  cardExtensionDateNeeded: string | null;
  siteAddress: string | null;
  siteContact: string | null;
  englishUnderstandingConfirmed: boolean;
  tcAcknowledged: boolean;
  gdprConsent: boolean;
  dateRegistered: string | null;
  inductionDate: string | null;
  stageOfNvq: string | null;
  notes: string | null;
  completedDate: string | null;
  certificationDate: string | null;
  customerUpdateNotes: string | null;
  trainingOutcome: string | null;
  outcomeDate: string | null;
  assessorTrainer: string | null;
  outcomeNotes: string | null;
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

  const companyNumber = companyLookupId
    ? (lookups.companyNumberById.get(companyLookupId) ?? null)
    : null;
  const workforceNumber = candidateLookupId
    ? (lookups.workforceNumberById.get(candidateLookupId) ?? null)
    : null;

  const completedDate = asNullableString(item.fields[nvqFields.completedDate]);
  return {
    id: item.id,
    candidateName,
    companyName,
    companyNumber,
    workforceNumber,
    niNumber: asNullableString(item.fields[nvqFields.niNumber]),
    ulnNumber: asNullableString(item.fields[nvqFields.ulnNumber]),
    nvqTitle: asNullableString(item.fields[nvqFields.nvqTitle]),
    boltOn: asNullableString(item.fields[nvqFields.boltonNvq]),
    poNumber: asNullableString(item.fields[nvqFields.poNumber]),
    cardSchemeCategory: asNullableString(
      item.fields[nvqFields.cardSchemeCategory],
    ),
    cardExtensionDateNeeded: asNullableString(
      item.fields[nvqFields.cardExtensionDateNeeded],
    ),
    siteAddress: asNullableString(item.fields[nvqFields.siteAddress]),
    siteContact: asNullableString(item.fields[nvqFields.siteContact]),
    englishUnderstandingConfirmed: asBoolean(
      item.fields[nvqFields.englishUnderstandingConfirmed],
    ),
    tcAcknowledged: asBoolean(item.fields[nvqFields.tcAcknowledged]),
    gdprConsent: asBoolean(item.fields[nvqFields.gdprConsent]),
    dateRegistered: asNullableString(item.fields[nvqFields.dateRegistered]),
    inductionDate: asNullableString(item.fields[nvqFields.dateInductionBooked]),
    stageOfNvq: asNullableString(item.fields[nvqFields.stageOfNvq]),
    notes: asNullableString(item.fields[nvqFields.notes]),
    completedDate,
    certificationDate: asNullableString(
      item.fields[nvqFields.certificationDate],
    ),
    customerUpdateNotes: asNullableString(
      item.fields[nvqFields.customerUpdateNotes],
    ),
    trainingOutcome: asNullableString(item.fields[nvqFields.trainingOutcome]),
    outcomeDate: asNullableString(item.fields[nvqFields.outcomeDate]),
    assessorTrainer: asNullableString(item.fields[nvqFields.assessorTrainer]),
    outcomeNotes: asNullableString(item.fields[nvqFields.outcomeNotes]),
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
  const workforceId = optionalText(input.workforceId);
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
    (workforceId
      ? (workforce.find((row) => row.id === workforceId) ?? null)
      : null) ??
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
    niNumber: optionalText(input.niNumber),
    ulnNumber: optionalText(input.ulnNumber),
    nvqTitle: optionalText(input.nvqTitle),
    boltonNvq: optionalText(input.boltOn),
    poNumber: optionalText(input.poNumber),
    cardSchemeCategory: optionalText(input.cardSchemeCategory),
    cardExtensionDateNeeded: asDateInput(input.cardExtensionDateNeeded),
    siteAddress: optionalText(input.siteAddress),
    siteContact: optionalText(input.siteContact),
    englishUnderstandingConfirmed:
      optionalBool(input.englishUnderstandingConfirmed) ?? false,
    tcAcknowledged: optionalBool(input.tcAcknowledged) ?? false,
    gdprConsent: optionalBool(input.gdprConsent) ?? false,
    dateRegistered: asDateInput(input.dateRegistered),
    dateInductionBooked: asDateInput(input.inductionDate),
    stageOfNvq: optionalText(input.stageOfNvq),
    notes: optionalText(input.notes),
    completedDate: asDateInput(input.completedDate),
    certificationDate: asDateInput(input.certificationDate),
    customerUpdateNotes: optionalText(input.customerUpdateNotes),
    trainingOutcome: optionalText(input.trainingOutcome),
    outcomeDate: asDateInput(input.outcomeDate),
    assessorTrainer: optionalText(input.assessorTrainer),
    outcomeNotes: optionalText(input.outcomeNotes),
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
    niNumber:
      input.niNumber === undefined ? undefined : optionalText(input.niNumber),
    ulnNumber:
      input.ulnNumber === undefined ? undefined : optionalText(input.ulnNumber),
    nvqTitle:
      input.nvqTitle === undefined ? undefined : optionalText(input.nvqTitle),
    boltonNvq:
      input.boltOn === undefined ? undefined : optionalText(input.boltOn),
    poNumber:
      input.poNumber === undefined ? undefined : optionalText(input.poNumber),
    cardSchemeCategory:
      input.cardSchemeCategory === undefined
        ? undefined
        : optionalText(input.cardSchemeCategory),
    cardExtensionDateNeeded:
      input.cardExtensionDateNeeded === undefined
        ? undefined
        : asDateInput(input.cardExtensionDateNeeded),
    siteAddress:
      input.siteAddress === undefined
        ? undefined
        : optionalText(input.siteAddress),
    siteContact:
      input.siteContact === undefined
        ? undefined
        : optionalText(input.siteContact),
    englishUnderstandingConfirmed: optionalBool(
      input.englishUnderstandingConfirmed,
    ),
    tcAcknowledged: optionalBool(input.tcAcknowledged),
    gdprConsent: optionalBool(input.gdprConsent),
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
    notes: input.notes === undefined ? undefined : optionalText(input.notes),
    completedDate:
      input.completedDate === undefined
        ? undefined
        : asDateInput(input.completedDate),
    certificationDate:
      input.certificationDate === undefined
        ? undefined
        : asDateInput(input.certificationDate),
    customerUpdateNotes:
      input.customerUpdateNotes === undefined
        ? undefined
        : optionalText(input.customerUpdateNotes),
    trainingOutcome:
      input.trainingOutcome === undefined
        ? undefined
        : optionalText(input.trainingOutcome),
    outcomeDate:
      input.outcomeDate === undefined
        ? undefined
        : asDateInput(input.outcomeDate),
    assessorTrainer:
      input.assessorTrainer === undefined
        ? undefined
        : optionalText(input.assessorTrainer),
    outcomeNotes:
      input.outcomeNotes === undefined
        ? undefined
        : optionalText(input.outcomeNotes),
    customerVisible: optionalBool(input.customerVisible),
  });

  if (
    input.workforceId !== undefined ||
    input.candidateName !== undefined ||
    input.companyName !== undefined
  ) {
    const lookups = await loadRegisterLookupMaps();
    const [companies, workforce] = await Promise.all([
      listAdminCompanies(),
      listAdminWorkforce(),
    ]);
    const existingMapped = mapNvq(existing, lookups);
    const workforceId = optionalText(input.workforceId);
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

    if (workforceId || candidateName) {
      const companyKey = companyName.trim().toLowerCase();
      const candidate =
        (workforceId
          ? (workforce.find((row) => row.id === workforceId) ?? null)
          : null) ??
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
  internalNotes: string | null;
  /**
   * Tentative = offered dates (Outlook tentative).
   * Confirmed = booking confirmed (Outlook busy + TM email).
   */
  bookingStatus: BookingStatus;
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
    internalNotes: stripSharePointHtml(
      asNullableString(item.fields[eventFields.internalNotes]),
    ),
    bookingStatus: bookingStatusFromFreeBusy(item.fields[eventFields.freeBusy]),
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
    const [hasEventCompany, hasInternalNotes] = await Promise.all([
      listHasColumn("events", "EventCompany"),
      listHasColumn("events", "InternalNotes"),
    ]);
    const warnings: string[] = [];
    if (!hasEventCompany) warnings.push(EVENT_COMPANY_MISSING_WARNING);
    if (!hasInternalNotes) {
      warnings.push(
        "Events list needs the InternalNotes multiple-lines text column.",
      );
    }
    return warnings;
  } catch {
    return [
      EVENT_COMPANY_MISSING_WARNING,
      "Events list needs the InternalNotes multiple-lines text column.",
    ];
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
  const bookingStatus =
    normalizeBookingStatus(input.bookingStatus) ?? "Tentative";
  const payload = omitNullishFields({
    ...toSharePointFields("events", {
      title,
      customerVisible: optionalBool(input.customerVisible) ?? true,
      trainingAddress: optionalText(input.trainingAddress),
      location: optionalText(input.location),
      eventDate: asDateTimeInput(input.eventDate),
      endDate: asDateTimeInput(input.endDate),
      description: optionalText(input.description),
      internalNotes: optionalText(input.internalNotes),
      freeBusy: freeBusyFromBookingStatus(bookingStatus),
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

  let bookingNotification: {
    attempted: boolean;
    skipped: boolean;
    skipReason?: string;
    recipients: string[];
    results: Array<{ status: string; recipientEmail: string; errorMessage?: string | null }>;
  } | null = null;

  try {
    const { notifyBookingConfirmed } = await import(
      "@/lib/services/eventBookingNotificationService"
    );
    bookingNotification = await notifyBookingConfirmed({
      eventId: withName.id,
      title: withName.title,
      companyId: withName.companyId,
      companyName: withName.company,
      eventDate: withName.eventDate,
      endDate: withName.endDate,
      location: withName.location,
      previousStatus: null,
      nextStatus: withName.bookingStatus,
    });
  } catch (error) {
    console.error("[events] booking confirmation notify failed:", error);
    bookingNotification = {
      attempted: true,
      skipped: true,
      skipReason:
        error instanceof Error
          ? error.message
          : "Booking confirmation notify failed.",
      recipients: [],
      results: [],
    };
  }

  return { record: withName, bookingNotification };
}

export async function updateAdminEvent(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("events", id);
  if (!existing) throw new NotFoundError("Event not found.");
  const previousMapped = mapEvent(existing);
  const previousStatus = previousMapped?.bookingStatus ?? null;

  const doNotSync =
    input.doNotSync === undefined ? undefined : optionalBool(input.doNotSync);

  const bookingStatus =
    input.bookingStatus === undefined
      ? undefined
      : (normalizeBookingStatus(input.bookingStatus) ?? "Tentative");

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
      internalNotes:
        input.internalNotes === undefined
          ? undefined
          : optionalText(input.internalNotes),
      freeBusy:
        bookingStatus === undefined
          ? undefined
          : freeBusyFromBookingStatus(bookingStatus),
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
    bookingStatus !== undefined ||
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

  let bookingNotification: {
    attempted: boolean;
    skipped: boolean;
    skipReason?: string;
    recipients: string[];
    results: Array<{ status: string; recipientEmail: string; errorMessage?: string | null }>;
  } | null = null;

  try {
    const { notifyBookingConfirmed } = await import(
      "@/lib/services/eventBookingNotificationService"
    );
    bookingNotification = await notifyBookingConfirmed({
      eventId: withName.id,
      title: withName.title,
      companyId: withName.companyId,
      companyName: withName.company,
      eventDate: withName.eventDate,
      endDate: withName.endDate,
      location: withName.location,
      previousStatus,
      nextStatus: withName.bookingStatus,
    });
  } catch (error) {
    console.error("[events] booking confirmation notify failed:", error);
    bookingNotification = {
      attempted: true,
      skipped: true,
      skipReason:
        error instanceof Error
          ? error.message
          : "Booking confirmation notify failed.",
      recipients: [],
      results: [],
    };
  }

  return { record: withName, bookingNotification };
}

export async function deleteAdminEvent(id: string): Promise<void> {
  const existing = await getListItemByKey("events", id);
  if (!existing) throw new NotFoundError("Event not found.");

  const mapped = mapEvent(existing);
  const outlookEventId = mapped?.outlookEventId?.trim() || null;

  if (outlookEventId) {
    try {
      const { deleteOutlookEvent } = await import(
        "@/lib/services/calendar/outlookCalendarService"
      );
      await deleteOutlookEvent(outlookEventId);
    } catch (error) {
      // SharePoint remains source of truth — still remove the portal event.
      console.error(
        `[events] Outlook delete failed for ${id} (${outlookEventId}):`,
        error,
      );
    }
  }

  await deleteListItemByKey("events", id);
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
  /** Permissions List Name — used by Workforce Training manager / Supervisor lookups. */
  name: string | null;
  /** Routing bucket: Admin portal vs Customer portal. */
  roleType: RoleType;
  /** Admin form value — includes first-class Candidate. */
  permissionRole: PermissionFormRole;
  sharePointRoleType: string;
  customerRole: CustomerRoleType | null;
  roleLabel: string;
  status: string;
  companyId: string | null;
  companyName: string | null;
  accessScope: string | null;
  /** Department names this person covers (TM / Supervisor portal filter). */
  departmentScopes: string[];
  /** Pipe/semicolon joined for admin multiselect form. */
  departmentsAllowed: string | null;
  canView: boolean;
  canDownload: boolean;
  canEdit: boolean;
}

/** Multi-lookup LookupId values (DepartmentsAllowed, etc.). */
function extractMultiLookupIds(
  fields: SharePointFields,
  fieldInternalName: string,
): string[] {
  const sibling = fields[`${fieldInternalName}LookupId`];
  if (Array.isArray(sibling)) {
    return sibling
      .map((entry) => String(entry ?? "").trim())
      .filter(Boolean);
  }
  if (typeof sibling === "number" || typeof sibling === "string") {
    const text = String(sibling).trim();
    return text ? [text] : [];
  }

  const direct = fields[fieldInternalName];
  if (!Array.isArray(direct)) return [];
  return direct
    .map((entry) => {
      if (entry && typeof entry === "object" && "LookupId" in entry) {
        const id = (entry as { LookupId?: unknown }).LookupId;
        return id === null || id === undefined ? "" : String(id).trim();
      }
      return "";
    })
    .filter(Boolean);
}

function mapPermission(item: SharePointListItem): AdminPermissionRecord | null {
  const userEmail = asString(item.fields[permissionFields.userEmail]);
  const sharePointRoleType =
    asString(item.fields[permissionFields.roleType])?.trim() || "";
  const roleType = normalizePermissionRoleType(sharePointRoleType);
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
  const accessScopeRaw =
    asNullableString(item.fields[permissionFields.accessScope]) ?? null;
  const accessScope =
    normalizeAccessScopeChoice(accessScopeRaw) ??
    (accessScopeRaw?.trim() || null);
  const customerRole = resolveCustomerRole(
    sharePointRoleType,
    accessScope || "Full Company",
  );
  const permissionRole = permissionFormRoleFromSharePoint(
    sharePointRoleType,
    accessScope || "Full Company",
  );
  const departmentScopes = Array.from(
    new Set([
      ...(asMultiChoiceText(item.fields[permissionFields.departments]) ?? "")
        .split(/[;,#|]+/)
        .map((part) => part.trim())
        .filter(Boolean),
      ...(asMultiChoiceText(item.fields[permissionFields.departmentsAllowed]) ??
        "")
        .split(/[;,#|]+/)
        .map((part) => part.trim())
        .filter(Boolean),
    ]),
  );
  const departmentAllowedIds = extractMultiLookupIds(
    item.fields,
    permissionFields.departmentsAllowed,
  );
  return {
    id: item.id,
    userEmail: userEmail.toLowerCase(),
    name: asNullableString(item.fields[permissionFields.name]),
    roleType,
    permissionRole,
    sharePointRoleType,
    customerRole,
    roleLabel: roleLabelFor(sharePointRoleType, customerRole),
    status: asNullableString(item.fields[permissionFields.status]) ?? "Inactive",
    companyId,
    companyName: asLookupOrString(item.fields[permissionFields.company]),
    // Blank/legacy SharePoint values break the required Access scope select.
    accessScope: accessScope || "Full Company",
    departmentScopes,
    // Prefer LookupIds for the admin multiselect form; fall back to names.
    departmentsAllowed: departmentAllowedIds.length
      ? departmentAllowedIds.join(", ")
      : departmentScopes.length
        ? departmentScopes.join(", ")
        : null,
    canView: asBoolean(item.fields[permissionFields.canView]),
    canDownload: asBoolean(item.fields[permissionFields.canDownload]),
    canEdit: asBoolean(item.fields[permissionFields.canEdit]),
  };
}

export async function listAdminPermissions(
  companies?: Awaited<ReturnType<typeof listAdminCompanies>>,
) {
  // Uncached: admin delete/save must never re-serve a removed Permissions row.
  const listId = getSharePointListId("permissions");
  const [items, companyRows] = await Promise.all([
    getListItems(listId, { top: 5000 }),
    companies ? Promise.resolve(companies) : listAdminCompanies(),
  ]);
  const companyNameById = new Map(
    companyRows.map((row) => [row.id, row.companyName] as const),
  );
  return items
    .map(mapPermission)
    .filter((row): row is AdminPermissionRecord => row !== null)
    .map((row) => ({
      ...row,
      // Graph often returns Company LookupId without LookupValue.
      companyName:
        row.companyName?.trim() ||
        (row.companyId ? (companyNameById.get(row.companyId) ?? null) : null),
    }))
    .sort((a, b) => a.userEmail.localeCompare(b.userEmail));
}

function normalizeAccessScopeChoice(value: unknown): string | null {
  const raw = optionalText(value);
  if (!raw) return null;
  const s = raw.toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (s.includes("candidate")) return "Candidate Only";
  if (s.includes("department")) return "Department Only";
  if (
    s.includes("full") ||
    s === "company" ||
    s === "all" ||
    s.includes("all compan") ||
    s === "company wide" ||
    s === "companywide"
  ) {
    return "Full Company";
  }
  return raw;
}

export async function createAdminPermission(input: Record<string, unknown>) {
  const userEmail = requireText(input.userEmail, "User email").toLowerCase();
  // Accept legacy roleType values plus first-class Candidate.
  const permissionRole =
    normalizePermissionFormRole(input.permissionRole) ??
    normalizePermissionFormRole(input.roleType);
  if (!permissionRole) {
    throw new ValidationError(
      "Role must be Training Manager, Supervisor, or Candidate.",
    );
  }
  const companyId = optionalText(input.companyId);
  const companyName = optionalText(input.companyName);
  if (
    (permissionRole === "Customer" || permissionRole === "Candidate") &&
    !companyId &&
    !companyName
  ) {
    throw new ValidationError("Customer / Candidate permissions require a company.");
  }

  const accessScope =
    permissionRole === "Candidate"
      ? "Candidate Only"
      : (normalizeAccessScopeChoice(input.accessScope) ?? "Full Company");
  const displayName =
    optionalText(input.name) ??
    userEmail.split("@")[0]?.replace(/[._]/g, " ") ??
    null;

  const payload: SharePointFields = toSharePointFields("permissions", {
    userEmail,
    roleType: toSharePointRoleType(permissionRole),
    status: optionalText(input.status) ?? "Active",
    accessScope,
    canView: optionalBool(input.canView) ?? true,
    canDownload: optionalBool(input.canDownload) ?? false,
    canEdit: optionalBool(input.canEdit) ?? false,
    name: displayName ?? undefined,
  });

  if (companyId) {
    payload[permissionFields.companyLookupId] = Number.isNaN(Number(companyId))
      ? companyId
      : Number(companyId);
  } else if (companyName) {
    payload[permissionFields.company] = companyName;
  }

  await applyPermissionDepartmentCoverage(payload, input, companyId);

  let item;
  try {
    item = await createListItemByKey("permissions", payload);
  } catch (error) {
    rethrowPermissionWriteError(error);
  }
  const mapped = mapPermission(item);
  if (!mapped) throw new Error("Created permission could not be mapped.");

  // One invite email when someone is added to Permissions (never fail create).
  try {
    const { sendPortalInviteNotification } = await import(
      "@/lib/services/notificationService"
    );
    await sendPortalInviteNotification({
      to: mapped.userEmail,
      displayName: mapped.name,
      companyName: mapped.companyName,
      roleLabel: mapped.roleLabel,
      itemId: mapped.id,
    });
  } catch {
    // Invite delivery is best-effort; permission row is already saved.
  }

  return mapped;
}

export async function updateAdminPermission(
  id: string,
  input: Record<string, unknown>,
) {
  const existing = await getListItemByKey("permissions", id);
  if (!existing) throw new NotFoundError("Permission not found.");
  const previous = mapPermission(existing);

  const nextStatus = optionalText(input.status);
  if (
    nextStatus &&
    nextStatus.toLowerCase() === "inactive" &&
    previous?.userEmail
  ) {
    try {
      assertNotProtectedAdmin(previous.userEmail, "deactivate");
    } catch (error) {
      throw new ValidationError(
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  const nextEmail =
    input.userEmail === undefined
      ? undefined
      : requireText(input.userEmail, "User email").toLowerCase();
  if (
    nextEmail &&
    previous?.userEmail &&
    nextEmail !== previous.userEmail.toLowerCase()
  ) {
    try {
      assertNotProtectedAdmin(previous.userEmail, "delete");
    } catch (error) {
      throw new ValidationError(
        "This Permissions account is protected — its email cannot be changed from the app.",
      );
    }
  }

  let sharePointRole: string | undefined;
  let permissionRole: PermissionFormRole | null = null;
  if (input.permissionRole !== undefined || input.roleType !== undefined) {
    permissionRole =
      normalizePermissionFormRole(input.permissionRole) ??
      normalizePermissionFormRole(input.roleType);
    if (!permissionRole) {
      throw new ValidationError(
        "Role must be Training Manager, Supervisor, or Candidate.",
      );
    }
    sharePointRole = toSharePointRoleType(permissionRole);
  }

  const forcedCandidateScope =
    sharePointRole === "Candidate" ||
    normalizePermissionFormRole(input.permissionRole) === "Candidate" ||
    normalizePermissionFormRole(input.roleType) === "Candidate";

  let accessScope: string | undefined = forcedCandidateScope
    ? "Candidate Only"
    : input.accessScope === undefined
      ? undefined
      : (normalizeAccessScopeChoice(input.accessScope) ?? undefined);

  const companyIdForDepts =
    optionalText(input.companyId) ??
    extractLookupId(existing.fields, permissionFields.company);

  // Parse department coverage early so we can validate Department Only.
  const deptRaw =
    input.departmentsAllowed !== undefined
      ? input.departmentsAllowed
      : input.departmentScopes;
  const deptTokens = Array.isArray(deptRaw)
    ? deptRaw.map((part) => String(part).trim()).filter(Boolean)
    : String(deptRaw ?? "")
        .split(/[;,#|]+/)
        .map((part) => part.trim())
        .filter(Boolean);

  if (
    accessScope &&
    accessScope.toLowerCase().includes("department") &&
    deptTokens.length === 0
  ) {
    // Saving Department Only with no coverage is invalid and often trips Graph.
    // Fall back to Full Company so admins can still edit company/role/email.
    accessScope = "Full Company";
  }

  const corePayload: SharePointFields = toSharePointFields("permissions", {
    userEmail:
      input.userEmail === undefined
        ? undefined
        : requireText(input.userEmail, "User email").toLowerCase(),
    status: optionalText(input.status) ?? undefined,
    canView: optionalBool(input.canView),
    canDownload: optionalBool(input.canDownload),
    canEdit: optionalBool(input.canEdit),
    name:
      input.name === undefined
        ? undefined
        : (optionalText(input.name) ?? undefined),
  });

  if (input.companyId !== undefined) {
    const companyId = optionalText(input.companyId);
    if (companyId) {
      corePayload[permissionFields.companyLookupId] = Number.isNaN(
        Number(companyId),
      )
        ? companyId
        : Number(companyId);
    }
  } else if (input.companyName !== undefined) {
    const companyName = optionalText(input.companyName);
    if (companyName) {
      corePayload[permissionFields.company] = companyName;
    }
  }

  // Choice / lookup fields — write separately so a bad AccessScope choice
  // does not block email/company/status updates.
  const deferred: SharePointFields = {};
  if (sharePointRole) {
    deferred[permissionFields.roleType] = sharePointRole;
  }
  if (accessScope) {
    deferred[permissionFields.accessScope] = accessScope;
  }

  try {
    await applyPermissionDepartmentCoverage(
      deferred,
      {
        ...input,
        accessScope: accessScope ?? input.accessScope,
        // Only apply dept write when tokens are present.
        departmentsAllowed:
          deptTokens.length > 0 ? deptTokens.join(", ") : undefined,
        departmentScopes: undefined,
      },
      companyIdForDepts,
    );
  } catch (error) {
    if (error instanceof ValidationError) throw error;
    rethrowPermissionWriteError(error);
  }

  let item;
  try {
    item = await updateListItemFieldsByKey("permissions", id, corePayload);
  } catch (error) {
    rethrowPermissionWriteError(error);
  }

  const choiceWarnings: string[] = [];
  for (const [field, value] of Object.entries(deferred)) {
    if (value === undefined) continue;
    try {
      await updateListItemFieldsByKey("permissions", id, { [field]: value });
    } catch (error) {
      console.warn(
        `[permissions] deferred field ${field} failed for #${id}`,
        error,
      );
      choiceWarnings.push(
        field === permissionFields.accessScope
          ? "Access scope could not be saved — check SharePoint AccessScope choices."
          : field === permissionFields.roleType
            ? "Role could not be saved — check SharePoint RoleType choices."
            : `${field} could not be saved.`,
      );
    }
  }

  const refreshed = await getListItemByKey("permissions", id);
  let mapped = mapPermission(refreshed ?? item);
  if (!mapped) {
    // Graph sometimes returns sparse fields after patch — build from input + prior.
    mapped = previous
      ? {
          ...previous,
          userEmail:
            optionalText(input.userEmail)?.toLowerCase() ?? previous.userEmail,
          name:
            input.name === undefined
              ? previous.name
              : (optionalText(input.name) ?? previous.name),
          status: optionalText(input.status) ?? previous.status,
          accessScope: accessScope ?? previous.accessScope,
          companyId: optionalText(input.companyId) ?? previous.companyId,
          companyName:
            optionalText(input.companyName) ?? previous.companyName,
          permissionRole: permissionRole ?? previous.permissionRole,
          canView: optionalBool(input.canView) ?? previous.canView,
          canDownload: optionalBool(input.canDownload) ?? previous.canDownload,
          canEdit: optionalBool(input.canEdit) ?? previous.canEdit,
        }
      : null;
  }
  if (!mapped) {
    throw new ValidationError(
      "Permission was saved but could not be reloaded. Refresh the page.",
    );
  }

  return { record: mapped, choiceWarnings };
}

export async function deleteAdminPermission(id: string): Promise<void> {
  const trimmed = String(id ?? "").trim();
  if (!trimmed) {
    throw new ValidationError("Permission id is required.");
  }

  // Uncached existence check — cached nulls used to early-return "success"
  // while SharePoint still had the row (then refresh brought it back).
  const listId = getSharePointListId("permissions");
  const existing = await getListItemById(listId, trimmed);
  if (!existing) {
    revalidateSharePointList("permissions");
    return;
  }

  const mapped = mapPermission(existing);
  try {
    assertNotProtectedAdmin(mapped?.userEmail, "delete");
  } catch (error) {
    throw new ValidationError(
      error instanceof Error ? error.message : String(error),
    );
  }

  // Workforce Training manager / Supervisor are Lookups → Permissions with
  // Restrict Delete. Clear those refs first or SharePoint blocks the delete.
  const { clearInboundLookupsToPermission } = await import(
    "@/lib/services/adminSafeDelete"
  );
  await clearInboundLookupsToPermission(trimmed);

  try {
    await deleteListItemByKey("permissions", trimmed);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    const related =
      /related to another item|cannot be deleted because/i.test(message);
    throw new ValidationError(
      related
        ? "SharePoint still blocks deleting this permission because another list row references it. Cleared Workforce Training manager / Supervisor links — if it keeps failing, ask a Site Owner which list has a lookup to Permissions with “Restrict delete”."
        : message.includes("SharePoint") || message.includes("delete")
          ? message
          : `Could not delete this permission. ${message}`,
    );
  }

  // Prove the Graph delete/recycle stuck before telling the UI it worked.
  const stillThere = await getListItemById(listId, trimmed);
  if (stillThere) {
    throw new ValidationError(
      "SharePoint still has this Permissions row after delete. It may be locked by Restrict Delete or a retention rule — check linked Workforce Training manager / Supervisor fields, then retry.",
    );
  }
  revalidateSharePointList("permissions");
  revalidatePath("/admin/permissions");
}

/**
 * Write DepartmentsAllowed (multi-lookup) from form names / ids.
 * Also set AccessScope to Department Only when coverage is provided for TM/Supervisor.
 */
async function applyPermissionDepartmentCoverage(
  payload: SharePointFields,
  input: Record<string, unknown>,
  companyId: string | null | undefined,
): Promise<void> {
  if (
    input.departmentsAllowed === undefined &&
    input.departmentScopes === undefined
  ) {
    return;
  }

  const raw =
    input.departmentsAllowed !== undefined
      ? input.departmentsAllowed
      : input.departmentScopes;
  const names = Array.isArray(raw)
    ? raw.map((part) => String(part).trim()).filter(Boolean)
    : String(raw ?? "")
        .split(/[;,#|]+/)
        .map((part) => part.trim())
        .filter(Boolean);

  // Empty coverage is normal for Full Company — do not load Departments or write
  // LookupId:[] (Graph often returns a 500 for an empty multi-lookup on create).
  if (names.length === 0) {
    return;
  }

  let departments;
  try {
    const { listAdminDepartments } = await import(
      "@/lib/services/departmentService"
    );
    departments = await listAdminDepartments(companyId ?? null);
  } catch (error) {
    console.error("[permissions] listAdminDepartments failed:", error);
    throw new ValidationError(
      "Could not load departments for this company. Check Admin → Departments, then try again.",
    );
  }

  const ids: number[] = [];
  for (const name of names) {
    const hit =
      departments.find((row) => row.id === name) ??
      departments.find(
        (row) =>
          row.name.trim().toLowerCase() === name.trim().toLowerCase(),
      );
    if (!hit) {
      throw new ValidationError(
        `Department "${name}" was not found for this company. Add it under Departments first.`,
      );
    }
    ids.push(Number(hit.id));
  }

  payload[`${permissionFields.departmentsAllowed}LookupId`] = ids;

  // When coverage is set, keep portal scoped to those departments.
  const currentScope = String(
    payload[permissionFields.accessScope] ?? input.accessScope ?? "",
  ).toLowerCase();
  if (!currentScope.includes("candidate")) {
    payload[permissionFields.accessScope] = "Department Only";
  }
}

function rethrowPermissionWriteError(error: unknown): never {
  if (error instanceof ValidationError) throw error;
  const message = error instanceof Error ? error.message : String(error);
  const lower = message.toLowerCase();
  if (
    lower.includes("accessscope") ||
    lower.includes("roletype") ||
    lower.includes("choice") ||
    lower.includes("departmentsallowed")
  ) {
    throw new ValidationError(
      "Could not save this permission — check Role, Access scope, and Departments match SharePoint choices, then try again.",
    );
  }
  if (
    lower.includes("generalexception") ||
    lower.includes("invalid request") ||
    lower.includes("invalidrequest")
  ) {
    throw new ValidationError(
      "SharePoint rejected this permission. Check the company is selected and try again. If it keeps failing, ask a Site Owner to verify the Permissions list columns.",
    );
  }
  throw new ValidationError(
    message.trim()
      ? `Could not save permission. ${message}`
      : "Could not save permission. Please try again.",
  );
}
