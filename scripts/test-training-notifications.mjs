/**
 * NPORS → Training Matrix / customer portal, and Training Manager result
 * emails: category parsing, "is this a new result", who may be emailed, and
 * what the admin is told.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-training-notifications.mjs
 *
 * (Node >= 22 for TypeScript type-stripping; the stub resolves "@/".)
 */
import test from "node:test";
import assert from "node:assert/strict";

const { parseNporsCategories } = await import("@/lib/training/nporsCategoryOptions");
const {
  checkTrainingManagerContact,
  resolveAction,
  summarizeTrainingManagerEmail,
} = await import("@/lib/services/trainingManagerNotificationRules");
const { NOTIFICATION_CATEGORIES, defaultNotificationTypes } = await import(
  "@/types/notifications"
);

// ---- NPORS categories → Training Matrix columns --------------------------

test("every NPORS code is read, not just the 8 legacy ones (Hasnain's record)", () => {
  // Live register #59 stores MultiChoice ["N217","N216","N214","N212"]; the old
  // parser knew only N001…N100, returned [], and the Pass never reached the Matrix.
  assert.deepEqual(parseNporsCategories(["N217", "N216", "N214", "N212"]), [
    "N217",
    "N216",
    "N214",
    "N212",
  ]);
  assert.deepEqual(parseNporsCategories("N217, N216; N214#N212"), [
    "N217",
    "N216",
    "N214",
    "N212",
  ]);
});

test("codes are exact: N301A is not also N301, labels are accepted", () => {
  assert.deepEqual(parseNporsCategories("N301A - Abrasive Wheels Awareness"), ["N301A"]);
  assert.deepEqual(parseNporsCategories(["N001 - Ind FLT", "n100 - exc crane"]), ["N001", "N100"]);
  assert.deepEqual(parseNporsCategories("N0011, XN001"), []);
  assert.deepEqual(parseNporsCategories(null), []);
});

// ---- result vs update -----------------------------------------------------

test("a Pass saved on create is a result email", () => {
  assert.equal(resolveAction("added", { trainingOutcome: "Pass" }), "confirmed");
  assert.equal(resolveAction("added", { trainingOutcome: null }), "added");
});

test("an edit sends a result email only when Pass/Fail changes", () => {
  assert.equal(resolveAction("updated", { trainingOutcome: "Pass" }, null), "confirmed");
  assert.equal(resolveAction("updated", { trainingOutcome: "Fail" }, "Pass"), "confirmed");
  assert.equal(resolveAction("updated", { trainingOutcome: "Pass" }, "Pass"), "updated");
  assert.equal(resolveAction("updated", { trainingOutcome: "Pass" }, undefined), "confirmed");
  assert.equal(resolveAction("updated", { trainingOutcome: null }, "Pass"), "updated");
});

// ---- who may be emailed -----------------------------------------------------

function profile(overrides = {}) {
  return {
    id: "7",
    userEmail: "tm@acw.example",
    status: "Active",
    customerRole: "TrainingManager",
    companyId: "423",
    companyDisplayName: "ACW Surfacing",
    customerNotificationsEnabled: true,
    candidateScopeName: "Wayne",
    ...overrides,
  };
}

test("assigned, active, same-company TM with notifications on is emailed", () => {
  const result = checkTrainingManagerContact(profile(), "423");
  assert.deepEqual(result, { contact: { email: "tm@acw.example", displayName: "Wayne" } });
});

test("a TM from another company is never emailed", () => {
  const result = checkTrainingManagerContact(profile({ companyId: "999", companyDisplayName: "Other Ltd" }), "423");
  assert.match(result.skip, /different company \(Other Ltd\)/);
});

test("inactive TM, notifications off, or unknown company → not emailed, with reason", () => {
  assert.match(checkTrainingManagerContact(profile({ status: "Inactive" }), "423").skip, /Inactive, not Active/);
  assert.match(
    checkTrainingManagerContact(profile({ customerNotificationsEnabled: false }), "423").skip,
    /notifications switched off/,
  );
  assert.match(checkTrainingManagerContact(profile(), null).skip, /no company/);
  assert.match(checkTrainingManagerContact(null, "423").skip, /missing or incomplete/);
});

test("a full Admin assigned as TM (all companies) may be emailed", () => {
  const result = checkTrainingManagerContact(profile({ customerRole: null, companyId: "0" }), "423");
  assert.equal(result.contact.email, "tm@acw.example");
});

// ---- what the admin is told -----------------------------------------------

test('"sent" only when Graph accepted the email', () => {
  const sent = summarizeTrainingManagerEmail({
    attempted: true,
    skipped: false,
    recipient: "tm@acw.example",
    result: { status: "sent", recipientEmail: "tm@acw.example", subject: "s", logged: true },
  });
  assert.equal(sent.status, "sent");

  const failed = summarizeTrainingManagerEmail({
    attempted: true,
    skipped: true,
    recipient: "tm@acw.example",
    result: { status: "failed", recipientEmail: "tm@acw.example", subject: "s", errorMessage: "Mailbox not found", logged: true },
  });
  assert.equal(failed.status, "failed");
  assert.match(failed.message, /FAILED: Mailbox not found/);

  const notConfigured = summarizeTrainingManagerEmail({
    attempted: true,
    skipped: true,
    result: { status: "not_configured", recipientEmail: "x", subject: "s", logged: true },
  });
  assert.equal(notConfigured.status, "not_configured");

  const turnedOff = summarizeTrainingManagerEmail({
    attempted: true,
    skipped: true,
    recipient: "tm@acw.example",
    result: { status: "skipped", recipientEmail: "x", subject: "s", errorMessage: "NPORS results notifications are turned off in Settings.", logged: true },
  });
  assert.equal(turnedOff.status, "skipped");
  assert.match(turnedOff.message, /turned off in Settings/);
  assert.doesNotMatch(turnedOff.message, /emailed \(/);
});

test("every switchable notification type defaults to on", () => {
  const defaults = defaultNotificationTypes();
  assert.equal(Object.keys(defaults).length, NOTIFICATION_CATEGORIES.length);
  assert.ok(Object.values(defaults).every(Boolean));
  assert.ok("training_npors" in defaults);
});
