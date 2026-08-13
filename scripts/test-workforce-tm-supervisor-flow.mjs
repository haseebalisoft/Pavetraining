#!/usr/bin/env node
/**
 * Offline behavioural test for the Workforce Training Manager / Supervisor
 * refactor:
 *
 *   1. Strict SP RoleType filter on the Workforce dropdown — TM options are
 *      RoleType = "Training Manager" only (no real Admin, no Candidate).
 *   2. Sidecar text write path in `applyWorkforcePersonLookups`:
 *        - matched          → Lookup id + sidecar text both set
 *        - unmatched + no email + allowUnresolvedText → Permissions row with placeholder email
 *        - unmatched + email + createIfMissing        → real row + Lookup + sidecar
 *        - unmatched + no email + createIfMissing (admin form) → throws
 *        - cleared          → Lookup id null + sidecar text cleared
 *   3. Read fallback ordering — Lookup label → sidecar text → id map.
 *   4. Permissions page tab filter — tab id maps to SP RoleType strictly.
 */

const green = (s) => `\x1b[32m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const bold = (s) => `\x1b[1m${s}\x1b[0m`;

let checks = 0;
let failed = 0;
function assert(label, got, expected) {
  checks += 1;
  const ok = JSON.stringify(got) === JSON.stringify(expected);
  if (!ok) failed += 1;
  console.log(
    `${ok ? green("PASS") : red("FAIL")}  ${label}` +
      (ok
        ? ""
        : `\n         expected ${JSON.stringify(expected)}\n         got      ${JSON.stringify(got)}`),
  );
}

// ---- 1. Strict SP RoleType filter ---------------------------------------
// Mirror of AdminCrudPage sharePointRoleTypeFilter branch (see the render
// block that matches on `spRole === targetKey`).

function normalizeSpRole(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace("trainingmanager", "training manager");
}

function sharePointRoleTypeFilter(people, target, companyId) {
  const targetKey = target.toLowerCase().replace(/\s+/g, " ");
  return people
    .filter((p) => {
      if ((p.status || "").toLowerCase() !== "active") return false;
      const spRole = normalizeSpRole(p.sharePointRoleType);
      if (spRole !== targetKey && !(target === "Training Manager" && spRole === "trainingmanager")) {
        return false;
      }
      if (companyId && p.companyId && p.companyId !== companyId) return false;
      return true;
    });
}

const people = [
  { id: "1", name: "Real Admin", sharePointRoleType: "Admin", status: "Active", companyId: "10" },
  { id: "2", name: "Amelia TM", sharePointRoleType: "Training Manager", status: "Active", companyId: "10" },
  { id: "3", name: "Ben TM (other co)", sharePointRoleType: "Training Manager", status: "Active", companyId: "20" },
  { id: "4", name: "Callum Sup", sharePointRoleType: "Supervisor", status: "Active", companyId: "10" },
  { id: "5", name: "Cand One", sharePointRoleType: "Candidate", status: "Active", companyId: "10" },
  { id: "6", name: "Inactive TM", sharePointRoleType: "Training Manager", status: "Inactive", companyId: "10" },
  { id: "7", name: "Legacy TM", sharePointRoleType: "TrainingManager", status: "Active", companyId: "10" },
];

console.log(bold("\nWorkforce TM / Supervisor flow"));
console.log("=".repeat(64));

console.log(bold("\n1. strict sharePointRoleTypeFilter"));
{
  const tms = sharePointRoleTypeFilter(people, "Training Manager", "10");
  assert("TM dropdown for company 10 includes only active TMs (incl. legacy spelling)", tms.map((p) => p.id).sort(), ["2", "7"]);

  const sups = sharePointRoleTypeFilter(people, "Supervisor", "10");
  assert("Supervisor dropdown for company 10", sups.map((p) => p.id), ["4"]);

  const admins = sharePointRoleTypeFilter(people, "Admin", "10");
  assert("Admin dropdown excludes TMs and Supervisors", admins.map((p) => p.id), ["1"]);

  const otherCompanyTms = sharePointRoleTypeFilter(people, "Training Manager", "20");
  assert("TM dropdown for company 20 shows only that company's TMs", otherCompanyTms.map((p) => p.id), ["3"]);
}

// ---- 2. Sidecar text write path ----------------------------------------
// Mirror of applyWorkforcePersonLookups (with the four new branches).

function findPerson(people, key, roleType, companyId) {
  const target = String(key ?? "").trim().toLowerCase();
  if (!target) return null;
  return (
    people.find((p) => {
      if (p.roleType !== roleType) return false;
      if ((p.status || "").toLowerCase() !== "active") return false;
      if (companyId && p.companyId && p.companyId !== companyId) return false;
      return (
        String(p.userEmail ?? "").toLowerCase() === target ||
        String(p.name ?? "").toLowerCase() === target
      );
    }) ?? null
  );
}

function applyLookup({
  value,
  email,
  roleType,
  companyId,
  people,
  createIfMissing,
  allowUnresolvedText,
}) {
  const payload = {};
  const text = value?.trim();
  const lookupIdField = "LookupId";
  const textField = "Text";
  if (!text) {
    payload[lookupIdField] = null;
    payload[textField] = "";
    return { payload, people };
  }
  const trimmedEmail = email?.trim().toLowerCase() || null;
  let hit =
    (trimmedEmail ? findPerson(people, trimmedEmail, roleType, companyId) : null) ??
    findPerson(people, text, roleType, companyId);
  if (!hit && createIfMissing && (trimmedEmail || allowUnresolvedText)) {
    const slug = text
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ".")
      .replace(/^\.+|\.+$/g, "")
      .slice(0, 48) || "person";
    const role = roleType === "Admin" ? "tm" : "sp";
    const row = {
      id: String(people.length + 1),
      name: text,
      userEmail: trimmedEmail || `pending.${role}.${slug}@pave.local`,
      roleType,
      status: "Active",
      companyId,
    };
    people = [...people, row];
    hit = row;
  }
  if (!hit) {
    if (allowUnresolvedText) {
      payload[lookupIdField] = null;
      payload[textField] = text;
      return { payload, people };
    }
    throw new Error(`"${text}" not found — throw`);
  }
  payload[lookupIdField] = Number(hit.id);
  payload[textField] = hit.name?.trim() || hit.userEmail;
  return { payload, people };
}

console.log(bold("\n2. applyWorkforcePersonLookups sidecar writes"));

const seedPeople = [
  { id: "1", name: "Amelia Hart", userEmail: "amelia@a.co", roleType: "Admin", status: "Active", companyId: "10" },
  { id: "2", name: "Callum Price", userEmail: "callum@a.co", roleType: "Customer", status: "Active", companyId: "10" },
];

{
  const r = applyLookup({
    value: "Amelia Hart",
    email: null,
    roleType: "Admin",
    companyId: "10",
    people: seedPeople,
    createIfMissing: false,
    allowUnresolvedText: false,
  });
  assert("matched by name → Lookup id set", r.payload.LookupId, 1);
  assert("matched by name → sidecar text set", r.payload.Text, "Amelia Hart");
}

{
  const r = applyLookup({
    value: "Not In List",
    email: null,
    roleType: "Admin",
    companyId: "10",
    people: seedPeople,
    createIfMissing: true,
    allowUnresolvedText: true,
  });
  assert("unmatched + no email + bulk auto-create → Lookup id set", Number.isFinite(r.payload.LookupId), true);
  assert("unmatched + no email + bulk auto-create → placeholder email", r.people.at(-1).userEmail, "pending.tm.not.in.list@pave.local");
  assert("unmatched + no email + bulk auto-create → sidecar text set", r.payload.Text, "Not In List");
}

{
  const r = applyLookup({
    value: "Priya New",
    email: "priya@b.co",
    roleType: "Admin",
    companyId: "10",
    people: seedPeople,
    createIfMissing: true,
    allowUnresolvedText: true,
  });
  assert("unmatched + email → new row created", r.people.length, seedPeople.length + 1);
  assert("unmatched + email → new email persisted (no @pave.local)", r.people.at(-1).userEmail, "priya@b.co");
  assert("unmatched + email → Lookup id set to new row", r.payload.LookupId, Number(r.people.at(-1).id));
  assert("unmatched + email → sidecar text set", r.payload.Text, "Priya New");
}

{
  let threw = false;
  let msg = "";
  try {
    applyLookup({
      value: "Ghost",
      email: null,
      roleType: "Admin",
      companyId: "10",
      people: seedPeople,
      createIfMissing: true,
      allowUnresolvedText: false,
    });
  } catch (err) {
    threw = true;
    msg = err instanceof Error ? err.message : String(err);
  }
  assert("admin form path (createIfMissing, no email, no unresolvedText) → throws", threw, true);
  assert("throw message mentions 'not found'", /not found/.test(msg), true);
}

{
  const r = applyLookup({
    value: null,
    email: null,
    roleType: "Admin",
    companyId: "10",
    people: seedPeople,
    createIfMissing: true,
    allowUnresolvedText: false,
  });
  assert("cleared → Lookup id null", r.payload.LookupId, null);
  assert("cleared → sidecar text cleared", r.payload.Text, "");
}

// ---- 3. Read fallback ordering ----------------------------------------
// Mirror of mapWorkforce trainingManager block.

function readTM({ lookupLabel, sidecarText, lookupId, permissionNameById }) {
  return (
    lookupLabel ??
    sidecarText ??
    (lookupId && permissionNameById ? (permissionNameById.get(String(lookupId)) ?? null) : null)
  );
}

console.log(bold("\n3. read fallback ordering"));
{
  const perms = new Map([["42", "Priya From Id Map"]]);
  assert("Lookup label wins", readTM({ lookupLabel: "Live Name", sidecarText: "Old", lookupId: 42, permissionNameById: perms }), "Live Name");
  assert("no Lookup label → sidecar wins", readTM({ lookupLabel: null, sidecarText: "Sidecar", lookupId: 42, permissionNameById: perms }), "Sidecar");
  assert("no Lookup label + no sidecar → id map", readTM({ lookupLabel: null, sidecarText: null, lookupId: 42, permissionNameById: perms }), "Priya From Id Map");
  assert("nothing → null", readTM({ lookupLabel: null, sidecarText: null, lookupId: null, permissionNameById: perms }), null);
}

// ---- 4. Permissions tab filter ----------------------------------------
// Mirror of AdminPermissionsClient rowFilter (normalizes case + spelling).

function permissionTabFilter(rows, tabId) {
  if (tabId === "all") return rows;
  const target = normalizeSpRole(tabId);
  return rows.filter((row) => normalizeSpRole(row.sharePointRoleType) === target);
}

console.log(bold("\n4. Permissions tab filter"));
{
  const rows = [
    { id: "a", sharePointRoleType: "Admin" },
    { id: "b", sharePointRoleType: "Training Manager" },
    { id: "c", sharePointRoleType: "TrainingManager" },
    { id: "d", sharePointRoleType: "Supervisor" },
    { id: "e", sharePointRoleType: "Candidate" },
    { id: "f", sharePointRoleType: "" },
  ];
  assert("all tab → all rows", permissionTabFilter(rows, "all").length, 6);
  assert("Admin tab → only Admin", permissionTabFilter(rows, "Admin").map((r) => r.id), ["a"]);
  assert("Training Manager tab folds legacy TrainingManager", permissionTabFilter(rows, "Training Manager").map((r) => r.id), ["b", "c"]);
  assert("Supervisor tab", permissionTabFilter(rows, "Supervisor").map((r) => r.id), ["d"]);
  assert("Candidate tab", permissionTabFilter(rows, "Candidate").map((r) => r.id), ["e"]);
}

console.log(
  "\n" +
    (failed === 0
      ? green(bold(`ALL GOOD — ${checks} checks passed.`))
      : red(bold(`${failed}/${checks} checks failed.`))),
);
process.exit(failed === 0 ? 0 : 1);
