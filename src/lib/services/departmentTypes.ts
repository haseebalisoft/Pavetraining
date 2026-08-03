/** Client-safe department types / constants (no server-only). */

export const MAX_DEPARTMENTS_PER_COMPANY = 10;

export type AdminDepartmentRecord = {
  id: string;
  name: string;
  companyId: string | null;
  companyName: string | null;
};
