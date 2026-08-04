/**
 * Offline unit checks for admin safe-delete cascade contracts.
 * Run: node scripts/test-admin-safe-delete.mjs
 *
 * Does not call SharePoint — validates cascade order and inbound-lookup maps
 * that keep Restrict Delete from blocking admin deletes.
 */

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function read(rel) {
  return readFileSync(join(root, rel), "utf8");
}

const safeDelete = read("src/lib/services/adminSafeDelete.ts");
const cascade = read("src/lib/services/companyCascadeDeleteService.ts");
const department = read("src/lib/services/departmentService.ts");
const crud = read("src/lib/services/adminCrudService.ts");

// 1) Company cascade order: workforce before departments before permissions
const orderMatch = safeDelete.match(
  /export const COMPANY_SAFE_DELETE_ORDER = \[([\s\S]*?)\] as const/,
);
assert.ok(orderMatch, "COMPANY_SAFE_DELETE_ORDER export present");
const order = [...orderMatch[1].matchAll(/"([^"]+)"/g)].map((m) => m[1]);
assert.deepEqual(
  order,
  [
    "customerDocuments",
    "trainingMatrix",
    "trainingMatrixCategoryRecords",
    "nporsRegister",
    "eusrRegister",
    "nrswaRegister",
    "inHouseCertificates",
    "nvqRegister",
    "events",
    "workforce",
    "departments",
    "permissions",
  ],
  "company cascade order must clear Candidate/Dept/Permission blockers",
);

const workforceIdx = order.indexOf("workforce");
const deptIdx = order.indexOf("departments");
const permIdx = order.indexOf("permissions");
assert.ok(workforceIdx < deptIdx, "workforce deleted before departments");
assert.ok(deptIdx < permIdx, "departments deleted before permissions");

// 2) Cascade service uses safe domain deletes
assert.match(
  cascade,
  /deleteMode:\s*"workforce"/,
  "cascade uses workforce-aware delete",
);
assert.match(
  cascade,
  /deleteMode:\s*"department"/,
  "cascade uses department-aware delete",
);
assert.match(
  cascade,
  /deleteMode:\s*"permission"/,
  "cascade uses permission-aware delete",
);

const cascadeBody = cascade.slice(
  cascade.indexOf("COMPANY_CASCADE_TARGETS"),
  cascade.indexOf("export interface CompanyCascadeResult"),
);
const workforcePos = cascadeBody.indexOf('listKey: "workforce"');
const deptPos = cascadeBody.indexOf('listKey: "departments"');
const permPos = cascadeBody.indexOf('listKey: "permissions"');
assert.ok(workforcePos > 0 && deptPos > workforcePos, "cascade: workforce before departments");
assert.ok(permPos > deptPos, "cascade: departments before permissions");

// 3) Department delete clears inbound lookups
assert.match(
  department,
  /clearInboundLookupsToDepartment/,
  "department delete clears Workforce + Permissions refs",
);

// 4) Workforce / Permission deletes use shared helper
assert.match(
  crud,
  /clearInboundLookupsToWorkforce/,
  "workforce delete clears candidate lookups",
);
assert.match(
  crud,
  /clearInboundLookupsToPermission/,
  "permission delete clears TM/Supervisor",
);

// 5) Standalone deletes exist for every admin list UI
for (const name of [
  "deleteAdminRegister",
  "deleteAdminNvq",
  "deleteAdminOffer",
  "deleteAdminDocument",
  "deleteAdminMatrix",
  "deleteAdminWorkforce",
  "deleteAdminPermission",
  "deleteAdminCompany",
  "deleteAdminEvent",
]) {
  assert.match(crud, new RegExp(`export async function ${name}`), `${name} exported`);
}

// 6) Inbound maps cover known Restrict Delete targets
assert.match(safeDelete, /CandidateNameLookupId/);
assert.match(safeDelete, /CandidateLookupId/);
assert.match(safeDelete, /Department0LookupId|departmentText\}LookupId/);
assert.match(safeDelete, /DepartmentsAllowedLookupId|departmentsAllowed\}LookupId/);
assert.match(safeDelete, /TrainingmanagerLookupId|trainingManager\}LookupId/);
assert.match(safeDelete, /SupervisorLookupId|supervisor\}LookupId/);

// 7) API DELETE handlers present
const deleteRoutes = [
  "src/app/api/admin/training-records/npors/[id]/route.ts",
  "src/app/api/admin/training-records/eusr/[id]/route.ts",
  "src/app/api/admin/training-records/streetworks/[id]/route.ts",
  "src/app/api/admin/training-records/in-house/[id]/route.ts",
  "src/app/api/admin/nvq/[id]/route.ts",
  "src/app/api/admin/offers/[id]/route.ts",
  "src/app/api/admin/documents/[id]/route.ts",
  "src/app/api/admin/workforce/[id]/route.ts",
  "src/app/api/admin/companies/[id]/route.ts",
  "src/app/api/admin/departments/[id]/route.ts",
  "src/app/api/admin/permissions/[id]/route.ts",
  "src/app/api/admin/events/[id]/route.ts",
  "src/app/api/admin/training-matrix/[id]/route.ts",
];
for (const rel of deleteRoutes) {
  const src = read(rel);
  assert.match(src, /export async function DELETE/, `${rel} has DELETE`);
}

// 8) UI wires deleteUrl for register/offers/nvq/matrix
assert.match(
  read("src/components/admin/pages/AdminRegisterClient.tsx"),
  /deleteUrl=\{/,
);
assert.match(
  read("src/components/admin/pages/AdminOffersClient.tsx"),
  /deleteUrl=\{/,
);
assert.match(
  read("src/components/admin/pages/AdminNvqClient.tsx"),
  /deleteUrl=\{/,
);
assert.match(
  read("src/components/admin/pages/AdminMatrixClient.tsx"),
  /deleteUrl=\{/,
);
assert.match(
  read("src/components/admin/pages/AdminDocumentsClient.tsx"),
  /method:\s*"DELETE"/,
);

console.log("OK — admin safe-delete contracts passed (" + deleteRoutes.length + " DELETE routes).");
