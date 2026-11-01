/**
 * Calendar: event documents (folder link), customer visibility (both
 * columns), and the Outlook mailbox setting.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-event-documents.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { eventDocumentSubfolders, eventIdFromDocumentFolder } = await import(
  "@/lib/training/eventDocumentFolders"
);
const { isEventCustomerVisible, eventVisibilityFields } = await import(
  "@/lib/training/eventVisibility"
);
const { cleanEnvValue } = await import("@/lib/config/envValue");
const { resolveDocumentFolderSegments } = await import("@/lib/services/documentFolderPaths");

test("event documents go to {company}/Company Documents/Events/{id}", () => {
  const { segments } = resolveDocumentFolderSegments({
    companyNumber: "C00001",
    companyName: "ACW Surfacing",
    documentType: "Other",
    hasCandidate: false,
    companySubfolders: eventDocumentSubfolders("65"),
  });
  assert.deepEqual(segments, ["C00001 - ACW Surfacing", "Company Documents", "Events", "65"]);
});

test("existing company documents keep their exact path", () => {
  const { segments } = resolveDocumentFolderSegments({
    companyNumber: "C00001",
    companyName: "ACW Surfacing",
    documentType: "Other",
    hasCandidate: false,
  });
  assert.deepEqual(segments, ["C00001 - ACW Surfacing", "Company Documents"]);
});

test("a document's folder identifies its event (and only event folders)", () => {
  const base = "/sites/PaveTrainingOperationAdmin/Customer Documents/C00001 - ACW Surfacing/Company Documents";
  assert.equal(eventIdFromDocumentFolder(`${base}/Events/65`), "65");
  assert.equal(eventIdFromDocumentFolder(`${base}/Events/65/`), "65");
  assert.equal(eventIdFromDocumentFolder(base), null);
  assert.equal(eventIdFromDocumentFolder(`${base}/Events`), null);
  assert.equal(eventIdFromDocumentFolder(`${base}/Events/65/old`), null);
  assert.equal(eventIdFromDocumentFolder(null), null);
});

test("unticking Customer visible hides the event even if the duplicate column is still true", () => {
  // Live state before the fix: the app wrote only Customer_x0020_Visible.
  assert.equal(isEventCustomerVisible({ Customer_x0020_Visible: false, CustomerVisible: true }), false);
  assert.equal(isEventCustomerVisible({ Customer_x0020_Visible: true, CustomerVisible: true }), true);
  assert.equal(isEventCustomerVisible({ Customer_x0020_Visible: true }), true);
  assert.equal(isEventCustomerVisible({ CustomerVisible: "Yes" }), true);
  assert.equal(isEventCustomerVisible({}), false, "no flag = not visible");
  assert.deepEqual(eventVisibilityFields(false), { Customer_x0020_Visible: false, CustomerVisible: false });
});

test("Outlook mailbox setting survives the paste mistake found in production", () => {
  assert.equal(cleanEnvValue("OUTLOOK_USER_ID", "OUTLOOK_USER_ID=info@pavetraining.co.uk"), "info@pavetraining.co.uk");
  assert.equal(cleanEnvValue("OUTLOOK_USER_ID", ' "info@pavetraining.co.uk" '), "info@pavetraining.co.uk");
  assert.equal(cleanEnvValue("OUTLOOK_USER_ID", "info@pavetraining.co.uk"), "info@pavetraining.co.uk");
  assert.equal(cleanEnvValue("OUTLOOK_USER_ID", undefined), "");
});
