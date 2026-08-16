/**
 * Single source of truth for admin top-nav items and role gates.
 *
 * Route-level guards (getAdminContext, withAdminApi) still enforce access —
 * this list only decides what the nav renders. Keep both aligned.
 *
 * Gate meaning:
 * - "any-admin"       → any user with `canAccessAdmin` (SP Admin OR Training Manager)
 * - "sharepoint-admin" → only SP RoleType === "Admin" (pure Admin, no customerRole)
 */
export type AdminNavGate = "any-admin" | "sharepoint-admin";

export type AdminNavItem = {
  href: string;
  label: string;
  exact?: boolean;
  gate: AdminNavGate;
};

export const ADMIN_PRIMARY_LINKS: AdminNavItem[] = [
  { href: "/admin", label: "Home", exact: true, gate: "any-admin" },
  { href: "/admin/companies", label: "Companies", gate: "any-admin" },
  { href: "/admin/departments", label: "Departments", gate: "any-admin" },
  { href: "/admin/workforce", label: "Workforce", gate: "any-admin" },
  { href: "/admin/training-matrix", label: "Matrix", gate: "any-admin" },
  { href: "/admin/documents", label: "Documents", gate: "any-admin" },
  { href: "/admin/events", label: "Calendar", gate: "any-admin" },
];

export const ADMIN_REGISTER_LINKS: AdminNavItem[] = [
  { href: "/admin/training-records", label: "All registers", gate: "any-admin" },
  { href: "/admin/training-records/npors", label: "NPORS", gate: "any-admin" },
  { href: "/admin/training-records/eusr", label: "EUSR", gate: "any-admin" },
  {
    href: "/admin/training-records/streetworks",
    label: "Streetworks",
    gate: "any-admin",
  },
  {
    href: "/admin/training-records/in-house",
    label: "In-House",
    gate: "any-admin",
  },
  { href: "/admin/nvq", label: "NVQ", gate: "any-admin" },
];

export const ADMIN_MORE_LINKS: AdminNavItem[] = [
  // Permissions edits who can access the portal — pure-Admin only.
  { href: "/admin/permissions", label: "Permissions", gate: "sharepoint-admin" },
  { href: "/admin/offers", label: "Offers", gate: "any-admin" },
  { href: "/admin/notifications", label: "Notifications", gate: "any-admin" },
  // Bulk Upload can mass-mutate workforce/matrix — pure-Admin only.
  { href: "/admin/bulk-upload", label: "Bulk Upload", gate: "sharepoint-admin" },
  { href: "/admin/logs", label: "Audit Log", gate: "any-admin" },
  { href: "/admin/settings", label: "Settings", gate: "any-admin" },
];

export function canSeeAdminNavItem(
  item: AdminNavItem,
  ctx: { isSharePointAdmin: boolean },
): boolean {
  if (item.gate === "any-admin") return true;
  if (item.gate === "sharepoint-admin") return ctx.isSharePointAdmin;
  return false;
}

/** Returns the flat set of hrefs a given role would see in the top-nav. */
export function allowedAdminNavHrefs(ctx: {
  isSharePointAdmin: boolean;
}): string[] {
  return [
    ...ADMIN_PRIMARY_LINKS,
    ...ADMIN_REGISTER_LINKS,
    ...ADMIN_MORE_LINKS,
  ]
    .filter((item) => canSeeAdminNavItem(item, ctx))
    .map((item) => item.href);
}
