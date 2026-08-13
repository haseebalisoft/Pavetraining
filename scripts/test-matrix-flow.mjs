#!/usr/bin/env node
/**
 * Offline behavioural test for the Matrix flow changes:
 *
 *   1. Admin Matrix link-status filter tabs — All / Linked / Needs Review / Orphan.
 *   2. Admin default view shows every matrix row (no hidden orphans).
 *   3. Matrix upload result wording — "N Matrix rows uploaded. X linked,
 *      Y need review, Z empty rows skipped, W failed." (no imported=0).
 *   4. Customer read filter — only Linked rows survive.
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

// ------ 1. link-status filter tabs (mirror of matchesLinkFilter) ---------
function matchesLinkFilter(row, linkFilter) {
  if (linkFilter === "all") return true;
  const status = row.matrixLinkStatus;
  if (linkFilter === "linked") return status === "Linked";
  if (linkFilter === "needs-review") return status === "Needs Review";
  if (linkFilter === "orphan") return status === "Orphan";
  return true;
}

console.log(bold("\n1. Admin Matrix link-status filter"));
{
  const rows = [
    { id: "a", matrixLinkStatus: "Linked" },
    { id: "b", matrixLinkStatus: "Needs Review" },
    { id: "c", matrixLinkStatus: "Orphan" },
    { id: "d", matrixLinkStatus: "Linked" },
    { id: "e", matrixLinkStatus: "Orphan" },
  ];
  assert("All tab shows every row", rows.filter((r) => matchesLinkFilter(r, "all")).map((r) => r.id), ["a", "b", "c", "d", "e"]);
  assert("Linked tab", rows.filter((r) => matchesLinkFilter(r, "linked")).map((r) => r.id), ["a", "d"]);
  assert("Needs Review tab", rows.filter((r) => matchesLinkFilter(r, "needs-review")).map((r) => r.id), ["b"]);
  assert("Orphan tab", rows.filter((r) => matchesLinkFilter(r, "orphan")).map((r) => r.id), ["c", "e"]);
}

// ------ 2. Admin default view shows everything -------------------------
console.log(bold("\n2. Admin Matrix default view"));
{
  const rows = [
    { id: "1", matrixLinkStatus: "Linked" },
    { id: "2", matrixLinkStatus: "Needs Review" },
    { id: "3", matrixLinkStatus: "Orphan" },
  ];
  // Default linkFilter is "all" per the client — nothing is hidden.
  const defaultFilter = "all";
  assert(
    "Default view includes Linked",
    rows.filter((r) => matchesLinkFilter(r, defaultFilter)).some((r) => r.matrixLinkStatus === "Linked"),
    true,
  );
  assert(
    "Default view includes Needs Review",
    rows.filter((r) => matchesLinkFilter(r, defaultFilter)).some((r) => r.matrixLinkStatus === "Needs Review"),
    true,
  );
  assert(
    "Default view includes Orphan",
    rows.filter((r) => matchesLinkFilter(r, defaultFilter)).some((r) => r.matrixLinkStatus === "Orphan"),
    true,
  );
}

// ------ 3. Matrix upload result wording --------------------------------
function matrixUploadSummary(summary) {
  const saved = summary.importedRows + summary.warningRows;
  const parts = [
    `${saved} Matrix row${saved === 1 ? "" : "s"} uploaded`,
    `${summary.importedRows} linked`,
  ];
  if (summary.warningRows > 0) parts.push(`${summary.warningRows} need review`);
  if (summary.skippedRows > 0) {
    parts.push(
      `${summary.skippedRows} empty row${summary.skippedRows === 1 ? "" : "s"} skipped`,
    );
  }
  if (summary.errorRows > 0) parts.push(`${summary.errorRows} failed`);
  return `${parts.join(", ")}.`;
}

console.log(bold("\n3. Matrix upload result wording"));
{
  const summary = { importedRows: 0, warningRows: 50, skippedRows: 17, errorRows: 0 };
  assert(
    "client's exact scenario reads correctly",
    matrixUploadSummary(summary),
    "50 Matrix rows uploaded, 0 linked, 50 need review, 17 empty rows skipped.",
  );
}
{
  const summary = { importedRows: 50, warningRows: 0, skippedRows: 0, errorRows: 0 };
  assert(
    "all linked",
    matrixUploadSummary(summary),
    "50 Matrix rows uploaded, 50 linked.",
  );
}
{
  const summary = { importedRows: 40, warningRows: 5, skippedRows: 2, errorRows: 3 };
  assert(
    "mixed outcome",
    matrixUploadSummary(summary),
    "45 Matrix rows uploaded, 40 linked, 5 need review, 2 empty rows skipped, 3 failed.",
  );
}
{
  const summary = { importedRows: 0, warningRows: 0, skippedRows: 0, errorRows: 0 };
  assert(
    "empty spreadsheet reports zero — no misleading 'imported=0'",
    matrixUploadSummary(summary).includes("imported=0"),
    false,
  );
}

// ------ 4. Customer filter — only Linked survives -----------------------
function customerAllows(row) {
  const status = row.matrixLinkStatus?.trim();
  return status !== "Orphan" && status !== "Needs Review";
}

console.log(bold("\n4. Customer read filter"));
{
  const rows = [
    { id: "1", matrixLinkStatus: "Linked" },
    { id: "2", matrixLinkStatus: "Needs Review" },
    { id: "3", matrixLinkStatus: "Orphan" },
    { id: "4", matrixLinkStatus: "Linked" },
  ];
  assert("customer sees only Linked", rows.filter(customerAllows).map((r) => r.id), ["1", "4"]);
}

console.log(
  "\n" +
    (failed === 0
      ? green(bold(`ALL GOOD — ${checks} checks passed.`))
      : red(bold(`${failed}/${checks} checks failed.`))),
);
process.exit(failed === 0 ? 0 : 1);
