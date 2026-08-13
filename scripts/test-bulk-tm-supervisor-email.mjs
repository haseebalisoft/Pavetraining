#!/usr/bin/env node
/**
 * Offline behavioural test for the new Workforce bulk-upload behaviour where
 * Training manager email / Supervisor email are provided on the spreadsheet.
 *
 * Verifies:
 *   1. When TM name + email match an existing Permissions row → row reused (no duplicate).
 *   2. When TM email matches but name differs → still reused (email is preferred).
 *   3. When TM is new + email provided → new Permissions row created with real email.
 *   4. When TM is new + no email → placeholder Permissions row created with pave.local email.
 *   5. Supervisor path mirrors TM behaviour.
 *
 * Reimplements the matching + auto-create semantics from the real code so
 * we can validate them without a live SharePoint dependency.
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

/** Mirror of adminCrudService.ts::permissionPersonKey */
function key(v) {
  return String(v ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}

/** Mirror of adminCrudService.ts::findPermissionPerson (matches by name OR email). */
function findPermissionPerson(people, nameOrEmail) {
  const k = key(nameOrEmail);
  if (!k) return null;
  return (
    people.find((row) => key(row.name ?? "") === k) ??
    people.find((row) => key(row.userEmail ?? "") === k) ??
    null
  );
}

/**
 * Mirror of ensurePermissionPerson:
 *   - reused when the email OR display name matches an existing row
 *   - creates a real row when a real email is provided
 *   - name-only creates a pending.{tm|sp}.{name}@pave.local Permissions row
 */
function placeholderPermissionEmail(displayName, roleType) {
  const slug =
    String(displayName)
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ".")
      .replace(/^\.+|\.+$/g, "")
      .slice(0, 48) || "person";
  const role = roleType === "Admin" ? "tm" : "sp";
  return `pending.${role}.${slug}@pave.local`;
}

function ensurePermissionPerson({ displayName, userEmail, roleType, people }) {
  const providedEmail = userEmail?.trim().toLowerCase() || null;
  const resolvedEmail =
    providedEmail ?? placeholderPermissionEmail(displayName, roleType);
  const existing =
    findPermissionPerson(people, resolvedEmail) ??
    findPermissionPerson(people, displayName);
  if (existing) return { people, created: false, row: existing };
  const row = {
    id: String(people.length + 1),
    name: displayName,
    userEmail: resolvedEmail,
    roleType,
  };
  return { people: [...people, row], created: true, row };
}

const initial = [
  { id: "1", name: "Amelia Hart", userEmail: "amelia.hart@example.com", roleType: "Admin" },
  { id: "2", name: "Callum Price", userEmail: "callum.price@example.com", roleType: "Customer" },
];

console.log(bold("\nBulk upload TM/Supervisor email behaviour"));
console.log("=".repeat(64));

// Case 1: name + email both match existing row → reused
{
  const r = ensurePermissionPerson({
    displayName: "Amelia Hart",
    userEmail: "amelia.hart@example.com",
    roleType: "Admin",
    people: initial,
  });
  assert("1. matched by both — row reused, not created", r.created, false);
  assert("1. total people unchanged", r.people.length, initial.length);
  assert("1. matched the existing amelia id", r.row.id, "1");
}

// Case 2: email match; name in spreadsheet is different — email wins, reused.
{
  const r = ensurePermissionPerson({
    displayName: "A. Hart",
    userEmail: "amelia.hart@example.com",
    roleType: "Admin",
    people: initial,
  });
  assert("2. email match beats name mismatch — row reused", r.created, false);
  assert("2. reused amelia row", r.row.id, "1");
}

// Case 3: new person + real email → real Permissions row created
{
  const r = ensurePermissionPerson({
    displayName: "Priya Sharma",
    userEmail: "priya.sharma@example.com",
    roleType: "Admin",
    people: initial,
  });
  assert("3. new person + real email — created=true", r.created, true);
  assert("3. new row uses the real email (no pave.local)", r.row.userEmail, "priya.sharma@example.com");
  assert("3. name preserved on new row", r.row.name, "Priya Sharma");
  assert("3. people list grew", r.people.length, initial.length + 1);
}

// Case 4: new person + NO email → placeholder Permissions row (appears in list)
{
  const r = ensurePermissionPerson({
    displayName: "New Person",
    userEmail: null,
    roleType: "Admin",
    people: initial,
  });
  assert("4. new person + no email — created=true", r.created, true);
  assert(
    "4. placeholder email is pending.tm.new.person@pave.local",
    r.row.userEmail,
    "pending.tm.new.person@pave.local",
  );
  assert("4. name preserved on placeholder row", r.row.name, "New Person");
}

// Case 5: supervisor path — mirror of tests above with roleType=Customer
{
  const r = ensurePermissionPerson({
    displayName: "Callum Price",
    userEmail: "callum.price@example.com",
    roleType: "Customer",
    people: initial,
  });
  assert("5a. supervisor matched by name+email — row reused", r.created, false);

  const r2 = ensurePermissionPerson({
    displayName: "New Supervisor",
    userEmail: "new.sup@example.com",
    roleType: "Customer",
    people: initial,
  });
  assert("5b. new supervisor with real email — created", r2.created, true);
  assert("5b. real email preserved", r2.row.userEmail, "new.sup@example.com");
  assert("5b. roleType is Customer", r2.row.roleType, "Customer");
}

// Case 6: idempotency — running the same "ensure" twice does not duplicate.
{
  const first = ensurePermissionPerson({
    displayName: "Priya Sharma",
    userEmail: "priya.sharma@example.com",
    roleType: "Admin",
    people: initial,
  });
  const second = ensurePermissionPerson({
    displayName: "Priya Sharma",
    userEmail: "priya.sharma@example.com",
    roleType: "Admin",
    people: first.people,
  });
  assert("6. second ensure is idempotent — no new row", second.created, false);
  assert("6. people length after 2 ensures", second.people.length, initial.length + 1);
}

console.log(
  "\n" +
    (failed === 0
      ? green(bold(`ALL GOOD — ${checks} checks passed.`))
      : red(bold(`${failed}/${checks} checks failed.`))),
);
process.exit(failed === 0 ? 0 : 1);
