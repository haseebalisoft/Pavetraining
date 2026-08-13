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

/**
 * Detect SharePoint's `tP_E_INVALIDLOOKUPPARENT` failure and rewrite it into an
 * actionable message. This error means the lookup column exists on the target
 * list, but the row it's pointing at either doesn't exist, was recycled, or the
 * lookup is bound to a different SharePoint list than the one whose IDs we're
 * sending. Common client tenants triggers:
 *   - Company row was hard-deleted (not recycled) after the page loaded.
 *   - The Departments list's `Company` lookup was configured against a
 *     different Company list than the SHAREPOINT_COMPANY_LIST_ID env var
 *     currently points at.
 *
 * Returns a friendly message when the underlying Graph error contains the
 * marker, otherwise `null` so the caller can rethrow the original error.
 */
function describeLookupParentError(
  error: unknown,
  targetListLabel: string,
  lookupColumn: string,
  ctx: { companyId?: string; companyName?: string | null } = {},
): string | null {
  const body =
    error && typeof error === "object" && "body" in error
      ? String((error as { body?: unknown }).body ?? "")
      : "";
  const message = error instanceof Error ? error.message : String(error ?? "");
  const combined = `${body} ${message}`.toLowerCase();
  if (!combined.includes("invalidlookupparent")) return null;
  const parts = [
    `SharePoint rejected the ${targetListLabel} write because its ${lookupColumn} lookup could not resolve the parent Company`,
  ];
  if (ctx.companyName) {
    parts.push(`"${ctx.companyName}"`);
  }
  if (ctx.companyId) {
    parts.push(`(id ${ctx.companyId})`);
  }
  parts.push(
    `— either that Company row no longer exists in the Company List, or the ${targetListLabel} list's ${lookupColumn} lookup is configured against a different Company list than SHAREPOINT_COMPANY_LIST_ID points at. Verify the Departments list column settings in SharePoint.`,
  );
  return parts.join(" ");
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
  // Explicit integer check: an accidental company name or GUID in the ID slot
  // would silently become NaN and SharePoint would reject with the confusing
  // `tP_E_INVALIDLOOKUPPARENT` error further below.
  const numericCompanyId = Number(companyId);
  if (!Number.isInteger(numericCompanyId) || numericCompanyId <= 0) {
    throw new ValidationError(
      `Company id "${companyId}" is not a valid SharePoint row id. Pick the company from the dropdown so the numeric id is captured.`,
    );
  }
  payload[`${fields.company}LookupId`] = numericCompanyId;

  let item;
  try {
    item = await createListItemByKey("departments", payload);
  } catch (error) {
    const message = describeLookupParentError(error, "departments", "Company", {
      companyId: String(numericCompanyId),
      companyName: company.companyName,
    });
    if (message) {
      throw new Error(message);
    }
    throw error;
  }
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
  let numericCompanyIdForPatch: number | null = null;
  if (input.companyId !== undefined) {
    const companyId = input.companyId.trim();
    if (!companyId) throw new ValidationError("Company is required.");
    const numericCompanyId = Number(companyId);
    if (!Number.isInteger(numericCompanyId) || numericCompanyId <= 0) {
      throw new ValidationError(
        `Company id "${companyId}" is not a valid SharePoint row id. Pick the company from the dropdown so the numeric id is captured.`,
      );
    }
    numericCompanyIdForPatch = numericCompanyId;
    payload[`${fields.company}LookupId`] = numericCompanyId;
  }
  if (input.status !== undefined) {
    payload[fields.status] = input.status;
  }
  if (input.notes !== undefined) {
    payload[fields.notes] = input.notes ?? "";
  }

  let item;
  try {
    item = await updateListItemFieldsByKey("departments", id, payload);
  } catch (error) {
    if (numericCompanyIdForPatch !== null) {
      const message = describeLookupParentError(
        error,
        "departments",
        "Company",
        { companyId: String(numericCompanyIdForPatch) },
      );
      if (message) throw new Error(message);
    }
    throw error;
  }
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
