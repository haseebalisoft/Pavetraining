import "server-only";

import { getSharePointFields } from "@/lib/schema/sharepointSchema";
import { getAllCompanies } from "@/lib/services/companyService";
import {
  NotFoundError,
  ValidationError,
} from "@/lib/services/errorHandler";
import {
  asLookupOrString,
  asNullableString,
  asString,
  createListItemByKey,
  extractLookupId,
  getListItemByKey,
  getListItemsByKey,
  toSharePointFields,
  updateListItemFieldsByKey,
  type SharePointFields,
} from "@/lib/services/sharePointListService";
import {
  DEFAULT_DEPARTMENT_STATUS,
  MAX_DEPARTMENTS_PER_COMPANY,
  normalizeDepartmentStatus,
  type AdminDepartmentRecord,
  type DepartmentStatus,
} from "@/lib/services/departmentTypes";

export {
  DEFAULT_DEPARTMENT_STATUS,
  DEPARTMENT_STATUSES,
  MAX_DEPARTMENTS_PER_COMPANY,
  isDepartmentActive,
  normalizeDepartmentStatus,
  type AdminDepartmentRecord,
  type DepartmentStatus,
} from "@/lib/services/departmentTypes";

const fields = getSharePointFields("departments");

export function departmentKey(value: string): string {
  return value.trim().toLowerCase().replace(/\s+/g, " ");
}

function mapDepartment(
  item: { id: string; fields: SharePointFields },
  companyNameById?: Map<string, string>,
): AdminDepartmentRecord | null {
  const name =
    asNullableString(item.fields[fields.name]) ??
    asNullableString(item.fields[fields.title]);
  if (!name?.trim()) return null;

  const companyId =
    extractLookupId(item.fields, fields.company) ??
    asString(item.fields[`${fields.company}LookupId`]);
  const companyName =
    asLookupOrString(item.fields[fields.company]) ??
    (companyId && companyNameById
      ? (companyNameById.get(companyId) ?? null)
      : null);

  return {
    id: item.id,
    name: name.trim(),
    companyId: companyId ?? null,
    companyName,
    status: normalizeDepartmentStatus(asNullableString(item.fields[fields.status])),
    notes: asNullableString(item.fields[fields.notes]),
  };
}

/**
 * Company departments, ACTIVE ONLY by default.
 *
 * Defaulting to active means any caller that feeds a picker (Workforce,
 * Permissions, bulk import matching) fails safe — a deactivated department can
 * never be offered for a new assignment just because a call site forgot a flag.
 * Pass `includeInactive: true` for admin screens that must show and reactivate
 * them, and for cascade delete which has to see every row.
 */
export async function listAdminDepartments(
  companyId?: string | null,
  options: { includeInactive?: boolean } = {},
): Promise<AdminDepartmentRecord[]> {
  const includeInactive = options.includeInactive ?? false;
  const [items, companies] = await Promise.all([
    getListItemsByKey("departments", { top: 5000 }),
    getAllCompanies(),
  ]);
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );

  return items
    .map((item) => mapDepartment(item, companyNameById))
    .filter((row): row is AdminDepartmentRecord => {
      if (!row) return false;
      if (!includeInactive && row.status !== "Active") return false;
      if (!companyId?.trim()) return true;
      return row.companyId === companyId.trim();
    })
    .sort((a, b) => {
      const companyCmp = (a.companyName ?? "").localeCompare(b.companyName ?? "");
      if (companyCmp !== 0) return companyCmp;
      return a.name.localeCompare(b.name);
    });
}

export async function createAdminDepartment(input: {
  name: string;
  companyId: string;
  /** Optional when company was created earlier in the same request (cache lag). */
  companyName?: string | null;
  /** When true, skip a second Departments list scan (caller already checked). */
  skipDuplicateScan?: boolean;
  status?: DepartmentStatus;
  notes?: string | null;
}): Promise<AdminDepartmentRecord> {
  const name = input.name.trim();
  const companyId = input.companyId.trim();
  if (!name) throw new ValidationError("Department name is required.");
  if (!companyId) throw new ValidationError("Company is required.");

  const companies = input.companyName?.trim()
    ? null
    : await getAllCompanies();
  const company =
    (companies?.find((row) => row.id === companyId) ?? null) ||
    (input.companyName?.trim()
      ? {
          id: companyId,
          companyName: input.companyName.trim(),
        }
      : null);
  if (!company) {
    throw new ValidationError("Company was not found.");
  }

  if (!input.skipDuplicateScan) {
    // Duplicate check must see INACTIVE rows too, or a deactivated "Civils"
    // would let a second "Civils" be created. The cap counts only Active, so a
    // deactivated department does not permanently consume one of the slots.
    const existing = await listAdminDepartments(companyId, {
      includeInactive: true,
    });
    const activeCount = existing.filter((row) => row.status === "Active").length;
    if (activeCount >= MAX_DEPARTMENTS_PER_COMPANY) {
      throw new ValidationError(
        `This company already has ${MAX_DEPARTMENTS_PER_COMPANY} active departments (maximum). Deactivate one first.`,
      );
    }
    const clash = existing.find(
      (row) => departmentKey(row.name) === departmentKey(name),
    );
    if (clash) {
      throw new ValidationError(
        clash.status === "Active"
          ? `Department "${name}" already exists for this company.`
          : `Department "${name}" already exists for this company but is Inactive. Set it back to Active instead of creating a duplicate.`,
      );
    }
  }

  const payload: SharePointFields = toSharePointFields("departments", {
    title: name,
    name,
    status: input.status ?? DEFAULT_DEPARTMENT_STATUS,
    notes: input.notes ?? null,
  });
  payload[`${fields.company}LookupId`] = Number(companyId);

  const item = await createListItemByKey("departments", payload);
  const mapped = mapDepartment(
    item,
    new Map([[company.id, company.companyName]]),
  );
  if (!mapped) throw new Error("Created department could not be mapped.");
  return mapped;
}

export async function updateAdminDepartment(
  id: string,
  input: {
    name?: string;
    companyId?: string;
    status?: DepartmentStatus;
    notes?: string | null;
  },
): Promise<{ record: AdminDepartmentRecord; previousStatus: DepartmentStatus }> {
  const existing = await getListItemByKey("departments", id);
  if (!existing) throw new NotFoundError("Department not found.");
  const previousStatus = normalizeDepartmentStatus(
    asNullableString(existing.fields[fields.status]),
  );

  const payload: SharePointFields = {};
  if (input.name !== undefined) {
    const name = input.name.trim();
    if (!name) throw new ValidationError("Department name is required.");
    payload[fields.title] = name;
    payload[fields.name] = name;
  }
  if (input.companyId !== undefined) {
    const companyId = input.companyId.trim();
    if (!companyId) throw new ValidationError("Company is required.");
    payload[`${fields.company}LookupId`] = Number(companyId);
  }
  if (input.status !== undefined) {
    payload[fields.status] = input.status;
  }
  if (input.notes !== undefined) {
    payload[fields.notes] = input.notes ?? "";
  }

  const item = await updateListItemFieldsByKey("departments", id, payload);
  const companies = await getAllCompanies();
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const mapped = mapDepartment(item, companyNameById);
  if (!mapped) throw new Error("Updated department could not be mapped.");
  return { record: mapped, previousStatus };
}

export async function deleteAdminDepartment(
  id: string,
): Promise<AdminDepartmentRecord> {
  const existing = await getListItemByKey("departments", id);
  if (!existing) throw new NotFoundError("Department not found.");
  const companies = await getAllCompanies();
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const deleted = mapDepartment(existing, companyNameById);

  // Workforce Department0 + Permissions DepartmentsAllowed Restrict Delete.
  const {
    clearInboundLookupsToDepartment,
    safeDeleteListItem,
  } = await import("@/lib/services/adminSafeDelete");
  await clearInboundLookupsToDepartment(id);
  await safeDeleteListItem("departments", id, { label: "department" });

  return (
    deleted ?? {
      id,
      name: "",
      companyId: null,
      companyName: null,
      status: DEFAULT_DEPARTMENT_STATUS,
      notes: null,
    }
  );
}

/**
 * Resolve department by name within a company (for Workforce assign and bulk
 * import). Does not create orphans without a company.
 *
 * Matches ACTIVE departments only. When `createIfMissing` is set and the only
 * same-named department is Inactive, that row is REACTIVATED rather than
 * duplicated — two departments with one name would break the workforce picker
 * and the Permissions coverage scopes.
 */
export async function resolveCompanyDepartment(input: {
  name: string;
  companyId: string | null | undefined;
  createIfMissing?: boolean;
  companyName?: string | null;
}): Promise<AdminDepartmentRecord | null> {
  const name = input.name.trim();
  if (!name) return null;
  const companyId = input.companyId?.trim() || null;
  const rows = await listAdminDepartments(companyId, { includeInactive: true });
  const matches = rows.filter(
    (row) => departmentKey(row.name) === departmentKey(name),
  );

  const active = matches.find((row) => row.status === "Active");
  if (active) return active;

  if (!input.createIfMissing || !companyId) return null;

  const inactive = matches[0];
  if (inactive) {
    const { record } = await updateAdminDepartment(inactive.id, {
      status: "Active",
    });
    return record;
  }

  return createAdminDepartment({
    name,
    companyId,
    companyName: input.companyName,
  });
}
