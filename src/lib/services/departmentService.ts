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
  MAX_DEPARTMENTS_PER_COMPANY,
  type AdminDepartmentRecord,
} from "@/lib/services/departmentTypes";

export {
  MAX_DEPARTMENTS_PER_COMPANY,
  type AdminDepartmentRecord,
} from "@/lib/services/departmentTypes";

const fields = getSharePointFields("departments");

function departmentKey(value: string): string {
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
  };
}

export async function listAdminDepartments(
  companyId?: string | null,
): Promise<AdminDepartmentRecord[]> {
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
}): Promise<AdminDepartmentRecord> {
  const name = input.name.trim();
  const companyId = input.companyId.trim();
  if (!name) throw new ValidationError("Department name is required.");
  if (!companyId) throw new ValidationError("Company is required.");

  const companies = await getAllCompanies();
  const company = companies.find((row) => row.id === companyId);
  if (!company) {
    throw new ValidationError("Company was not found.");
  }

  const existing = await listAdminDepartments(companyId);
  if (existing.length >= MAX_DEPARTMENTS_PER_COMPANY) {
    throw new ValidationError(
      `This company already has ${MAX_DEPARTMENTS_PER_COMPANY} departments (maximum).`,
    );
  }
  if (existing.some((row) => departmentKey(row.name) === departmentKey(name))) {
    throw new ValidationError(`Department "${name}" already exists for this company.`);
  }

  const payload: SharePointFields = toSharePointFields("departments", {
    title: name,
    name,
  });
  payload[`${fields.company}LookupId`] = Number(companyId);

  const item = await createListItemByKey("departments", payload);
  const mapped = mapDepartment(item, new Map([[company.id, company.companyName]]));
  if (!mapped) throw new Error("Created department could not be mapped.");
  return mapped;
}

export async function updateAdminDepartment(
  id: string,
  input: { name?: string; companyId?: string },
): Promise<AdminDepartmentRecord> {
  const existing = await getListItemByKey("departments", id);
  if (!existing) throw new NotFoundError("Department not found.");

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

  const item = await updateListItemFieldsByKey("departments", id, payload);
  const companies = await getAllCompanies();
  const companyNameById = new Map(
    companies.map((row) => [row.id, row.companyName] as const),
  );
  const mapped = mapDepartment(item, companyNameById);
  if (!mapped) throw new Error("Updated department could not be mapped.");
  return mapped;
}

export async function deleteAdminDepartment(id: string): Promise<void> {
  const existing = await getListItemByKey("departments", id);
  if (!existing) throw new NotFoundError("Department not found.");

  // Workforce Department0 + Permissions DepartmentsAllowed Restrict Delete.
  const {
    clearInboundLookupsToDepartment,
    safeDeleteListItem,
  } = await import("@/lib/services/adminSafeDelete");
  await clearInboundLookupsToDepartment(id);
  await safeDeleteListItem("departments", id, { label: "department" });
}

/**
 * Resolve department by name within a company (for Workforce assign).
 * Does not create orphans without a company.
 */
export async function resolveCompanyDepartment(input: {
  name: string;
  companyId: string | null | undefined;
  createIfMissing?: boolean;
}): Promise<AdminDepartmentRecord | null> {
  const name = input.name.trim();
  if (!name) return null;
  const companyId = input.companyId?.trim() || null;
  const rows = await listAdminDepartments(companyId);
  const hit =
    rows.find((row) => departmentKey(row.name) === departmentKey(name)) ?? null;
  if (hit) return hit;
  if (!input.createIfMissing || !companyId) return null;
  return createAdminDepartment({ name, companyId });
}
