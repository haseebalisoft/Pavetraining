/**
 * Diagnose why customer Candidates shows 0 for DBS.
 *   node --env-file=.env.local --import ./scripts/_register-e2e-hook.mjs scripts/diagnose-dbs-candidates.mjs
 */
import { getSharePointFields } from "../src/lib/schema/sharepointSchema.ts";
import { getSharePointListId } from "../src/lib/config/sharepoint.ts";
import {
  asLookupOrString,
  asString,
  extractLookupId,
  getListItems,
  getListItemsByKey,
} from "../src/lib/services/sharePointListService.ts";
import { getWorkforceByCompanyName } from "../src/lib/services/workforceService.ts";
import {
  filterCandidatesByAccess,
} from "../src/lib/services/customerAccessService.ts";
import {
  getActivePermissionByEmail,
  normalizeAccessScopeValue,
  resolveCustomerRole,
} from "../src/lib/services/permissionService.ts";

const wf = getSharePointFields("workforce");
const perm = getSharePointFields("permissions");

function companyLabel(fields) {
  const raw = fields[wf.companyName];
  return (
    asLookupOrString(raw) ||
    asString(raw) ||
    extractLookupId(fields, wf.companyName) ||
    "(blank)"
  );
}

const companies = await getListItemsByKey("company", { top: 500 });
const dbsCos = [];
for (const item of companies) {
  const name =
    asString(item.fields.Title) ||
    asString(item.fields.CompanyName) ||
    asString(item.fields.Company_x0020_Name) ||
    "";
  const number =
    asString(item.fields.CompanyNumber) ||
    asString(item.fields.Company_x0020_Number) ||
    "";
  if (/dbs/i.test(name) || /c00001/i.test(number) || number === "C00001") {
    dbsCos.push({
      id: item.id,
      name,
      number,
      status: asString(item.fields.Status),
    });
  }
}
console.log("=== Companies matching DBS / C00001 ===");
console.log(dbsCos);

const workforce = await getListItemsByKey("workforce", { top: 5000 });
const byCompany = new Map();
const dbsRows = [];
for (const item of workforce) {
  const company = companyLabel(item.fields);
  const companyId = extractLookupId(item.fields, wf.companyName);
  byCompany.set(company, (byCompany.get(company) || 0) + 1);
  const name =
    asString(item.fields[wf.candidateName]) || asString(item.fields.Title) || "";
  if (
    /dbs/i.test(String(company)) ||
    dbsCos.some((c) => String(c.id) === String(companyId))
  ) {
    dbsRows.push({
      id: item.id,
      name,
      company,
      companyId,
      tm:
        asLookupOrString(item.fields[wf.trainingManager]) ||
        asString(item.fields[wf.trainingManager]),
      status: asString(item.fields[wf.status]),
    });
  }
}
console.log("=== Workforce company tallies ===");
console.log(
  [...byCompany.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40),
);
console.log("=== Workforce linked to DBS ===", dbsRows.length);
console.log(dbsRows.slice(0, 30));

const permissions = await getListItems(getSharePointListId("permissions"), {
  top: 500,
});
const dbsPerms = [];
for (const item of permissions) {
  const email = asString(item.fields[perm.userEmail]);
  const role = asString(item.fields[perm.roleType]);
  const scope = asString(item.fields[perm.accessScope]);
  const status = asString(item.fields[perm.status]);
  const name = asString(item.fields[perm.name]);
  const companyVal = item.fields[perm.company];
  const companyId =
    asString(item.fields[perm.companyLookupId]) ||
    (companyVal && typeof companyVal === "object"
      ? String(companyVal.LookupId || "")
      : "");
  const companyName =
    typeof companyVal === "string"
      ? companyVal
      : companyVal?.LookupValue || "";
  if (
    /dbs/i.test(String(companyName)) ||
    dbsCos.some((c) => String(c.id) === String(companyId)) ||
    /dbs/i.test(String(email || ""))
  ) {
    const customerRole = resolveCustomerRole(role || "", scope || "Full Company");
    const normalized = normalizeAccessScopeValue(
      scope || "Full Company",
      customerRole,
      false,
    );
    dbsPerms.push({
      id: item.id,
      email,
      role,
      scope,
      status,
      name,
      companyId,
      companyName,
      customerRole,
      normalized,
    });
  }
}
console.log("=== Permissions for DBS ===");
console.log(dbsPerms);

// Simulate getWorkforceByCompanyName for each DBS company name
for (const co of dbsCos) {
  const rows = await getWorkforceByCompanyName(co.name);
  console.log(
    `=== getWorkforceByCompanyName(${JSON.stringify(co.name)}) => ${rows.length} ===`,
  );
  if (rows.length) {
    console.log(
      rows.slice(0, 5).map((r) => ({
        id: r.id,
        name: r.candidateName,
        company: r.companyName,
        tm: r.trainingManager,
      })),
    );
  }

  for (const p of dbsPerms.filter((x) => String(x.companyId) === String(co.id))) {
    const ctx = {
      loggedInEmail: (p.email || "").toLowerCase(),
      role: "Customer",
      customerRole: p.customerRole,
      roleLabel: p.role === "Training Manager" ? "Manager" : p.role,
      companyId: p.companyId,
      companyName: co.name,
      canView: true,
      canDownload: true,
      canEdit: false,
      accessScope: p.scope || "Full Company",
      normalizedAccessScope: p.normalized,
      departmentScopes: [],
      candidateScopeName: p.name,
      permissionStatus: p.status,
    };
    const allowed = filterCandidatesByAccess(rows, ctx);
    console.log(
      `  filter for ${p.email} scope=${p.normalized} => ${allowed.length}/${rows.length}`,
    );
  }
}
