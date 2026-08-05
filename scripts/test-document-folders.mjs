/**
 * Unit tests for the Customer Documents folder path/naming rules.
 *
 * Runs with the built-in Node test runner + native TypeScript type-stripping
 * (Node >= 22); documentFolderPaths.ts is dependency-free so it loads with no
 * Graph and no "@/" resolver hook:
 *
 *   node --test scripts/test-document-folders.mjs
 *
 * These lock the folder STRUCTURE and NAMING that bulk workforce import relies
 * on (the Graph create/idempotency behaviour lives in
 * customerDocumentsFolderService and is covered by the manual steps):
 *
 *   Customer Documents
 *   └── {Company Number} - {Company Name}
 *       ├── Company Documents
 *       └── Candidates
 *           └── {Workforce Number} - {Candidate Name}
 *               ├── Certificates
 *               ├── Card Scans
 *               ├── NVQ Documents
 *               └── Other Documents
 */
import test from "node:test";
import assert from "node:assert/strict";

const BASE = new URL("../src/lib/", import.meta.url).pathname;
const {
  CANDIDATE_SUBFOLDERS,
  COMPANY_LEVEL_FOLDERS,
  companyDocumentsFolderName,
  candidateDocumentsFolderName,
  resolveDocumentFolderSegments,
} = await import(BASE + "services/documentFolderPaths.ts");

// --- Required folder set (requirement 2) ---------------------------------

test("candidate subfolders are exactly the four required, in order", () => {
  assert.deepEqual(
    [...CANDIDATE_SUBFOLDERS],
    ["Certificates", "Card Scans", "NVQ Documents", "Other Documents"],
  );
});

test("company-level folders are Company Documents + Candidates", () => {
  assert.deepEqual([...COMPANY_LEVEL_FOLDERS], [
    "Company Documents",
    "Candidates",
  ]);
});

// --- Exact naming: Number + " - " + Name (requirement 5) -----------------

test("company folder = Company Number + ' - ' + Company Name", () => {
  assert.equal(
    companyDocumentsFolderName("C00024", "Murphy Plant Ltd"),
    "C00024 - Murphy Plant Ltd",
  );
});

test("candidate folder = Workforce Number + ' - ' + Candidate Name", () => {
  assert.equal(
    candidateDocumentsFolderName("W00195", "John Murphy"),
    "W00195 - John Murphy",
  );
});

test("blank / missing Workforce Number falls back to Candidate Name only", () => {
  assert.equal(candidateDocumentsFolderName("", "Legacy Person"), "Legacy Person");
  assert.equal(candidateDocumentsFolderName(null, "Legacy Person"), "Legacy Person");
  assert.equal(
    candidateDocumentsFolderName("   ", "Legacy Person"),
    "Legacy Person",
  );
});

test("blank / missing Company Number falls back to Company Name only", () => {
  assert.equal(companyDocumentsFolderName("", "Solo Co"), "Solo Co");
  assert.equal(companyDocumentsFolderName(null, "Solo Co"), "Solo Co");
});

test("illegal folder characters are sanitised out of names", () => {
  // < > : " / \ | ? * and control chars become spaces (collapsed).
  assert.equal(
    companyDocumentsFolderName("C1", 'A/B:C"D'),
    "C1 - A B C D",
  );
});

// --- Structure: candidate goes UNDER the correct company (test 4) --------

test("candidate segments nest under the company's Candidates folder", () => {
  const { segments, destinationFolder } = resolveDocumentFolderSegments({
    companyNumber: "C00024",
    companyName: "Murphy Plant Ltd",
    workforceNumber: "W00195",
    candidateName: "John Murphy",
    documentType: "Certificate",
  });
  assert.deepEqual(segments, [
    "C00024 - Murphy Plant Ltd",
    "Candidates",
    "W00195 - John Murphy",
    "Certificates",
  ]);
  assert.equal(destinationFolder, "Certificates");
  // The candidate folder is a child of THIS company's folder.
  assert.equal(segments[0], "C00024 - Murphy Plant Ltd");
  assert.equal(segments[1], "Candidates");
});

test("company-level documents route under Company Documents (no candidate)", () => {
  const { segments, destinationFolder } = resolveDocumentFolderSegments({
    companyNumber: "C00024",
    companyName: "Murphy Plant Ltd",
    hasCandidate: false,
  });
  assert.deepEqual(segments, ["C00024 - Murphy Plant Ltd", "Company Documents"]);
  assert.equal(destinationFolder, "Company Documents");
});

test("unknown document type for a candidate routes to Other Documents", () => {
  const { destinationFolder } = resolveDocumentFolderSegments({
    companyNumber: "C1",
    companyName: "Co",
    workforceNumber: "W1",
    candidateName: "Jo",
    documentType: "Random",
  });
  assert.equal(destinationFolder, "Other Documents");
});
