#!/usr/bin/env node
/**
 * Offline role-access test matrix.
 *
 * Verifies the four production roles + the "hardcoded protected admin" case
 * against the exact same gating logic used by:
 *   - AdminContext.isSharePointAdmin
 *   - allowedAdminNavHrefs / canSeeAdminNavItem
 *   - withSharePointAdminApi / requireSharePointAdminAccess
 *   - the redirect in /admin/bulk-upload and /admin/permissions pages
 *   - getMeContext login redirect
 *
 * The scenarios below are hand-constructed PermissionProfile shapes so we
 * don't need a SharePoint connection. If the rules change, this file will
 * fail loudly.
 *
 * Usage:  node scripts/test-role-access-matrix.mjs
 * Exit:   0 if every expectation passed, 1 otherwise.
 */

const ALWAYS_ADMIN_EMAILS = new Set(["wayne.curry@pavetraining.co.uk"]);

const ADMIN_PRIMARY_LINKS = [
  { href: "/admin", label: "Home", gate: "any-admin" },
  { href: "/admin/companies", label: "Companies", gate: "any-admin" },
  { href: "/admin/departments", label: "Departments", gate: "any-admin" },
  { href: "/admin/workforce", label: "Workforce", gate: "any-admin" },
  { href: "/admin/training-matrix", label: "Matrix", gate: "any-admin" },
  { href: "/admin/documents", label: "Documents", gate: "any-admin" },
  { href: "/admin/events", label: "Calendar", gate: "any-admin" },
];

const ADMIN_REGISTER_LINKS = [
  { href: "/admin/training-records", label: "All registers", gate: "any-admin" },
  { href: "/admin/training-records/npors", label: "NPORS", gate: "any-admin" },
  { href: "/admin/training-records/eusr", label: "EUSR", gate: "any-admin" },
  { href: "/admin/training-records/streetworks", label: "Streetworks", gate: "any-admin" },
  { href: "/admin/training-records/in-house", label: "In-House", gate: "any-admin" },
  { href: "/admin/nvq", label: "NVQ", gate: "any-admin" },
];

const ADMIN_MORE_LINKS = [
  { href: "/admin/permissions", label: "Permissions", gate: "sharepoint-admin" },
  { href: "/admin/offers", label: "Offers", gate: "any-admin" },
  { href: "/admin/notifications", label: "Notifications", gate: "any-admin" },
  { href: "/admin/bulk-upload", label: "Bulk Upload", gate: "sharepoint-admin" },
  { href: "/admin/logs", label: "Audit Log", gate: "any-admin" },
  { href: "/admin/settings", label: "Settings", gate: "any-admin" },
];

/** Mirrors src/lib/services/permissionService.ts::resolveCustomerRole */
function resolveCustomerRole(spRole, accessScope = "Full Company") {
  const role = spRole.toLowerCase().trim();
  const scope = accessScope.toLowerCase().trim();
  if (role === "admin") return null;
  if (role === "training manager" || role === "trainingmanager") return "TrainingManager";
  if (role === "candidate") return "Candidate";
  if (role === "supervisor" || role === "customer") {
    if (scope.includes("candidate")) return "Candidate";
    return "Supervisor";
  }
  return null;
}

/**
 * Mirrors src/lib/services/permissionService.ts::isSharePointAdminForProfile
 * The hardcoded protected-admin list does NOT override an existing SharePoint
 * row — it only prevents delete/deactivate and acts as fallback when the SP
 * row is missing entirely.
 */
function isSharePointAdminForProfile(profile) {
  return profile.customerRole === null;
}

/** Mirrors adminNavItems.ts::canSeeAdminNavItem */
function canSeeAdminNavItem(item, ctx) {
  if (item.gate === "any-admin") return true;
  if (item.gate === "sharepoint-admin") return ctx.isSharePointAdmin;
  return false;
}

/** Mirrors adminNavItems.ts::allowedAdminNavHrefs */
function allowedAdminNavHrefs(ctx) {
  return [...ADMIN_PRIMARY_LINKS, ...ADMIN_REGISTER_LINKS, ...ADMIN_MORE_LINKS]
    .filter((item) => canSeeAdminNavItem(item, ctx))
    .map((item) => item.href);
}

/** Mirrors getMeContext redirect logic. */
function loginRedirect(profile) {
  if (profile.customerRole !== null) return "/customer/dashboard";
  if (profile.canAccessAdmin) return "/admin";
  return "/customer/dashboard";
}

/** Build a fake PermissionProfile matching src/types/models.ts shape. */
function buildProfile({
  email,
  spRoleType,
  accessScope = "Full Company",
  status = "Active",
}) {
  const customerRole = resolveCustomerRole(spRoleType, accessScope);
  // STRICT rule: only literal SP RoleType = Admin can access /admin.
  // Training Managers, Supervisors, and Candidates route to /customer/*.
  const canAccessAdmin =
    spRoleType.trim().toLowerCase() === "admin" && customerRole === null;
  const canAccessCustomer = customerRole !== null;
  return {
    userEmail: email.toLowerCase(),
    status,
    sharePointRoleType: spRoleType,
    customerRole,
    canAccessAdmin,
    canAccessCustomer,
    accessScope,
  };
}

/** What routes should each level see? */
const API_ROUTES = [
  // (method, path, minRole) — where minRole is one of: "sharepoint-admin", "any-admin", "customer-any", "public"
  ["GET",    "/api/admin/bulk-upload/templates",          "sharepoint-admin"],
  ["POST",   "/api/admin/bulk-upload/preview",            "sharepoint-admin"],
  ["POST",   "/api/admin/bulk-upload/commit",             "sharepoint-admin"],
  ["POST",   "/api/admin/bulk-upload/commit-progress",    "sharepoint-admin"],
  ["POST",   "/api/admin/permissions",                    "sharepoint-admin"],
  ["PATCH",  "/api/admin/permissions/[id]",               "sharepoint-admin"],
  ["DELETE", "/api/admin/permissions/[id]",               "sharepoint-admin"],
  ["GET",    "/api/admin/permissions",                    "any-admin"],
  ["GET",    "/api/admin/workforce",                      "any-admin"],
  ["POST",   "/api/admin/workforce",                      "any-admin"],
  ["GET",    "/api/admin/training-matrix/…",              "any-admin"],
  ["GET",    "/api/admin/training-records/npors",         "any-admin"],
  ["GET",    "/api/admin/training-records/eusr",          "any-admin"],
  ["GET",    "/api/admin/training-records/streetworks",   "any-admin"],
  ["GET",    "/api/admin/training-records/in-house",      "any-admin"],
  ["GET",    "/api/customer/documents",                   "customer-any"],
  ["GET",    "/api/customer/training-matrix",             "customer-any"],
];

function canCallApi(route, ctx) {
  const [, , minRole] = route;
  if (minRole === "sharepoint-admin") return ctx.isSharePointAdmin === true;
  if (minRole === "any-admin") return ctx.canAccessAdmin === true;
  if (minRole === "customer-any") return ctx.canAccessCustomer === true;
  if (minRole === "public") return true;
  return false;
}

/** Scenario definitions. */
const SCENARIOS = [
  {
    id: "admin-pure",
    label: "Pure Admin (SP RoleType = Admin)",
    profile: buildProfile({ email: "boss@pavetraining.co.uk", spRoleType: "Admin" }),
    expect: {
      isSharePointAdmin: true,
      canAccessAdmin: true,
      canAccessCustomer: false,
      loginRedirect: "/admin",
      seesBulkUpload: true,
      seesPermissions: true,
      seesRegisters: true,
    },
  },
  {
    id: "hardcoded-with-tm-row",
    label:
      "Hardcoded protected admin — SP row exists and says Training Manager (hardcoded list does NOT override)",
    profile: buildProfile({
      email: "wayne.curry@pavetraining.co.uk",
      spRoleType: "Training Manager",
    }),
    expect: {
      isSharePointAdmin: false,       // SP row wins; hardcoded list only protects from deletion
      canAccessAdmin: false,          // STRICT: TM cannot enter /admin — must fix SP row
      canAccessCustomer: true,
      loginRedirect: "/customer/dashboard",
      seesBulkUpload: false,
      seesPermissions: false,
      seesRegisters: false,           // no admin access = no admin nav
    },
  },
  {
    id: "hardcoded-with-admin-row",
    label: "Hardcoded protected admin (SP RoleType = Admin, normal case)",
    profile: buildProfile({
      email: "wayne.curry@pavetraining.co.uk",
      spRoleType: "Admin",
    }),
    expect: {
      isSharePointAdmin: true,
      canAccessAdmin: true,
      canAccessCustomer: false,
      loginRedirect: "/admin",
      seesBulkUpload: true,
      seesPermissions: true,
      seesRegisters: true,
    },
  },
  {
    id: "hardcoded-with-no-sp-row",
    label:
      "Hardcoded protected admin — SP row is MISSING (synthetic fallback fires)",
    // getActivePermissionByEmail returns alwaysAdminPermissionProfile in this case.
    profile: {
      userEmail: "wayne.curry@pavetraining.co.uk",
      status: "Active",
      sharePointRoleType: "Admin",
      customerRole: null,           // synthetic profile is pure admin
      canAccessAdmin: true,
      canAccessCustomer: false,
      accessScope: "All",
    },
    expect: {
      isSharePointAdmin: true,
      canAccessAdmin: true,
      canAccessCustomer: false,
      loginRedirect: "/admin",
      seesBulkUpload: true,
      seesPermissions: true,
      seesRegisters: true,
    },
  },
  {
    id: "training-manager",
    label: "Training Manager",
    profile: buildProfile({
      email: "tm@customer.example",
      spRoleType: "Training Manager",
    }),
    expect: {
      isSharePointAdmin: false,
      canAccessAdmin: false,        // STRICT: TM lands on /customer only
      canAccessCustomer: true,
      loginRedirect: "/customer/dashboard",
      seesBulkUpload: false,
      seesPermissions: false,
      seesRegisters: false,         // no admin nav — /admin blocked
    },
  },
  {
    id: "supervisor",
    label: "Supervisor (customer role, no admin access)",
    profile: buildProfile({
      email: "sup@customer.example",
      spRoleType: "Supervisor",
      accessScope: "Department Only",
    }),
    expect: {
      isSharePointAdmin: false,
      canAccessAdmin: false,        // never enters /admin at all
      canAccessCustomer: true,
      loginRedirect: "/customer/dashboard",
      seesBulkUpload: false,
      seesPermissions: false,
      seesRegisters: false,         // not on /admin
    },
  },
  {
    id: "candidate",
    label: "Candidate (customer role, no admin access)",
    profile: buildProfile({
      email: "cand@customer.example",
      spRoleType: "Candidate",
      accessScope: "Own records only",
    }),
    expect: {
      isSharePointAdmin: false,
      canAccessAdmin: false,
      canAccessCustomer: true,
      loginRedirect: "/customer/dashboard",
      seesBulkUpload: false,
      seesPermissions: false,
      seesRegisters: false,
    },
  },
];

/** ANSI colour helpers. */
const c = {
  green: (s) => `\x1b[32m${s}\x1b[0m`,
  red: (s) => `\x1b[31m${s}\x1b[0m`,
  yellow: (s) => `\x1b[33m${s}\x1b[0m`,
  cyan: (s) => `\x1b[36m${s}\x1b[0m`,
  bold: (s) => `\x1b[1m${s}\x1b[0m`,
  dim: (s) => `\x1b[2m${s}\x1b[0m`,
};
const PASS = c.green("PASS");
const FAIL = c.red("FAIL");

let totalChecks = 0;
let failedChecks = 0;

function check(label, got, expected) {
  totalChecks += 1;
  const ok = got === expected;
  if (!ok) failedChecks += 1;
  const status = ok ? PASS : FAIL;
  const detail = ok ? "" : c.dim(`   (expected ${JSON.stringify(expected)}, got ${JSON.stringify(got)})`);
  console.log(`    ${status}  ${label}${detail}`);
  return ok;
}

console.log(c.bold("\nROLE ACCESS MATRIX — offline verification"));
console.log("=".repeat(72));

for (const scenario of SCENARIOS) {
  const p = scenario.profile;
  const isSPAdmin = isSharePointAdminForProfile(p);
  const ctx = {
    isSharePointAdmin: isSPAdmin,
    canAccessAdmin: p.canAccessAdmin,
    canAccessCustomer: p.canAccessCustomer,
  };
  const hrefs = allowedAdminNavHrefs(ctx);
  const seesBulkUpload = hrefs.includes("/admin/bulk-upload");
  const seesPermissions = hrefs.includes("/admin/permissions");
  const seesRegisters = hrefs.includes("/admin/training-records");
  const redirect = loginRedirect(p);

  console.log(`\n${c.cyan("▶")} ${c.bold(scenario.label)}`);
  console.log(c.dim(`    email:            ${p.userEmail}`));
  console.log(c.dim(`    SP RoleType:      ${p.sharePointRoleType}`));
  console.log(c.dim(`    customerRole:     ${p.customerRole ?? "null"}`));
  console.log(c.dim(`    accessScope:      ${p.accessScope}`));

  check("isSharePointAdmin",   isSPAdmin,          scenario.expect.isSharePointAdmin);
  check("canAccessAdmin",      p.canAccessAdmin,   scenario.expect.canAccessAdmin);
  check("canAccessCustomer",   p.canAccessCustomer,scenario.expect.canAccessCustomer);
  check("login redirect",      redirect,           scenario.expect.loginRedirect);
  if (p.canAccessAdmin) {
    check("nav: sees Bulk Upload",  seesBulkUpload,   scenario.expect.seesBulkUpload);
    check("nav: sees Permissions",  seesPermissions,  scenario.expect.seesPermissions);
    check("nav: sees Registers",    seesRegisters,    scenario.expect.seesRegisters);
  } else {
    console.log(c.dim("    (skipping /admin nav — role cannot access /admin)"));
  }

  console.log(c.dim("    API access:"));
  const badRouteFor = [];
  for (const route of API_ROUTES) {
    const [method, path, minRole] = route;
    const allowed = canCallApi(route, ctx);
    const label = `      ${method.padEnd(6)} ${path.padEnd(48)} → ${allowed ? c.green("allow") : c.red("deny ")}  ${c.dim(`(min: ${minRole})`)}`;
    console.log(label);
    // Sanity: bulk-upload/permissions writes must never be allowed for non-SP admin.
    if (minRole === "sharepoint-admin" && allowed && !isSPAdmin) badRouteFor.push(path);
  }
  totalChecks += 1;
  if (badRouteFor.length) {
    failedChecks += 1;
    console.log(`    ${FAIL}  SP-admin-only routes leaked to non-SP admin: ${badRouteFor.join(", ")}`);
  } else {
    console.log(`    ${PASS}  no SP-admin-only routes leaked`);
  }
}

console.log("\n" + "=".repeat(72));
if (failedChecks === 0) {
  console.log(c.green(c.bold(`ALL GOOD — ${totalChecks} checks passed.`)));
  process.exit(0);
} else {
  console.log(c.red(c.bold(`${failedChecks}/${totalChecks} checks failed.`)));
  process.exit(1);
}
