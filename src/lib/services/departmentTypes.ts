/** Client-safe department types / constants (no server-only). */

export const MAX_DEPARTMENTS_PER_COMPANY = 10;

/** SharePoint "Status" choice on the Departments list. */
export type DepartmentStatus = "Active" | "Inactive";

export const DEPARTMENT_STATUSES: DepartmentStatus[] = ["Active", "Inactive"];

export const DEFAULT_DEPARTMENT_STATUS: DepartmentStatus = "Active";

export type AdminDepartmentRecord = {
  id: string;
  name: string;
  companyId: string | null;
  companyName: string | null;
  /**
   * Never null on a mapped record: rows created before the Status column existed
   * read back as Active, so an unset choice cannot silently hide a department
   * from the Workforce / Permissions dropdowns.
   */
  status: DepartmentStatus;
  notes: string | null;
};

/**
 * Normalizes any stored Status value. Blank / unknown -> Active, so a legacy row
 * or a hand-edited cell can only ever fail open (visible), never disappear.
 */
export function normalizeDepartmentStatus(
  value: string | null | undefined,
): DepartmentStatus {
  return String(value ?? "").trim().toLowerCase() === "inactive"
    ? "Inactive"
    : "Active";
}

export function isDepartmentActive(department: {
  status?: string | null;
}): boolean {
  return normalizeDepartmentStatus(department.status) === "Active";
}
