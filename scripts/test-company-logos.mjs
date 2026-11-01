/**
 * Company bulk upload logos: which attached file goes to which imported
 * company, and that the report never claims an upload that did not happen.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-company-logos.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { uploadAttachedCompanyLogos, withLogoOutcome } = await import(
  "@/components/admin/bulkUpload/bulkCommitJob"
);

const file = (name) => new File([new Uint8Array([137, 80, 78, 71])], name, { type: "image/png" });

function imported(rowNumber, companyName, companyLogo, extra = {}) {
  return {
    rowNumber,
    status: "Imported",
    savedEntityId: String(100 + rowNumber),
    fields: { companyName, companyLogo },
    messages: [`Created company #${100 + rowNumber}.`],
    ...extra,
  };
}

function recorder() {
  const calls = [];
  return { calls, upload: async (companyId, f) => { calls.push([companyId, f.name]); } };
}

test("named logo file is uploaded and the waiting note is replaced", async () => {
  const { calls, upload } = recorder();
  const rows = [
    imported(2, "Alder & Finch Supplies Ltd", "alder-finch.png", {
      pendingLogoFile: "alder-finch.png",
      messages: ["Created company #102.", 'Company Logo "alder-finch.png" — waiting for the attached logo file.'],
    }),
  ];
  const report = await uploadAttachedCompanyLogos(rows, [file("Alder-Finch.PNG")], upload);
  assert.deepEqual(calls, [["102", "Alder-Finch.PNG"]]);
  assert.equal(report.uploaded, 1);
  assert.deepEqual(report.rows[0].messages, ["Created company #102.", 'Company Logo uploaded from "Alder-Finch.PNG".']);
});

test("named file not attached is reported as NOT uploaded, never as uploaded", async () => {
  const { calls, upload } = recorder();
  const rows = [
    imported(2, "Blue Lantern Logistics Ltd", "blue-lantern.png", {
      pendingLogoFile: "blue-lantern.png",
      messages: ['Company Logo "blue-lantern.png" — waiting for the attached logo file.'],
    }),
  ];
  const report = await uploadAttachedCompanyLogos(rows, [], upload);
  assert.equal(calls.length, 0);
  assert.deepEqual(report.notAttached, ["blue-lantern.png"]);
  assert.match(report.rows[0].messages.join(" "), /NOT uploaded — that file was not attached/);
  assert.doesNotMatch(report.rows[0].messages.join(" "), /waiting/);
});

test("placeholder cell falls back to a file named after the company", async () => {
  const { calls, upload } = recorder();
  const rows = [
    imported(3, "Cedar Grove Catering Ltd", "Logo placeholder", {
      messages: ['No logo ("Logo placeholder" is a placeholder) — the company\'s logo is left unchanged.'],
    }),
  ];
  const report = await uploadAttachedCompanyLogos(rows, [file("Cedar Grove Catering Ltd.jpg")], upload);
  assert.deepEqual(calls, [["103", "Cedar Grove Catering Ltd.jpg"]]);
  assert.deepEqual(report.rows[0].messages, ['Company Logo uploaded from "Cedar Grove Catering Ltd.jpg".']);
});

test("one file can serve several companies (group logo)", async () => {
  const { calls, upload } = recorder();
  const rows = [
    imported(2, "A Ltd", "group.png", { pendingLogoFile: "group.png" }),
    imported(3, "B Ltd", "group.png", { pendingLogoFile: "group.png" }),
  ];
  const report = await uploadAttachedCompanyLogos(rows, [file("group.png")], upload);
  assert.deepEqual(calls, [["102", "group.png"], ["103", "group.png"]]);
  assert.equal(report.uploaded, 2);
});

test("skipped duplicates and errors never get a logo", async () => {
  const { calls, upload } = recorder();
  const rows = [
    { rowNumber: 2, status: "Duplicate", matchedEntityId: "7", fields: { companyName: "A Ltd", companyLogo: "a.png" }, messages: [] },
    { rowNumber: 3, status: "Error", fields: { companyName: "B Ltd", companyLogo: "b.png" }, messages: [] },
  ];
  const report = await uploadAttachedCompanyLogos(rows, [file("a.png"), file("b.png")], upload);
  assert.equal(calls.length, 0);
  assert.deepEqual(report.unused, ["a.png", "b.png"]);
});

test("a failed upload is reported with the reason and counted", async () => {
  const rows = [imported(2, "A Ltd", "a.png", { pendingLogoFile: "a.png" })];
  const report = await uploadAttachedCompanyLogos(rows, [file("a.png")], async () => {
    throw new Error("Image must be 10 MB or smaller.");
  });
  assert.equal(report.failed, 1);
  assert.match(report.rows[0].messages.join(" "), /upload FAILED: Image must be 10 MB or smaller\./);
});

test("logo problems turn Completed into Completed with warnings", () => {
  const base = { outcome: "completed", reason: "Import finished: 2 imported, 0 skipped." };
  const clean = withLogoOutcome(base, { rows: [], uploaded: 2, failed: 0, notAttached: [], unused: [] });
  assert.deepEqual(clean, { outcome: "completed", reason: "Import finished: 2 imported, 0 skipped. 2 logo(s) uploaded." });
  const missing = withLogoOutcome(base, { rows: [], uploaded: 1, failed: 0, notAttached: ["b.png"], unused: [] });
  assert.equal(missing.outcome, "completed_with_warnings");
  assert.match(missing.reason, /1 logo\(s\) NOT uploaded \(not attached: b\.png\)/);
});
