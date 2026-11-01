/**
 * Workforce bulk upload candidate photos: each attached photo goes to the
 * profile the import saved, never to a same-named person, and a missing
 * photo never affects the candidate.
 *
 *   node --import ./scripts/register-server-only-stub.mjs --test scripts/test-candidate-photos.mjs
 */
import test from "node:test";
import assert from "node:assert/strict";

const { uploadAttachedCandidatePhotos, withImageOutcome } = await import(
  "@/components/admin/bulkUpload/bulkCommitJob"
);

const file = (name) => new File([new Uint8Array([255, 216, 255])], name, { type: "image/jpeg" });

function candidate(rowNumber, name, workforceNumber, candidatePhoto, extra = {}) {
  return {
    rowNumber,
    status: "Imported",
    savedEntityId: String(1600 + rowNumber),
    fields: { candidateName: name, workforceNumber, candidatePhoto },
    messages: ["Imported."],
    ...extra,
  };
}

function recorder() {
  const calls = [];
  return { calls, upload: async (id, f) => { calls.push([id, f.name]); } };
}

test("photo named in Candidate Photo goes to that candidate's profile", async () => {
  const { calls, upload } = recorder();
  const rows = [
    candidate(2, "Zaid", "PAVE-W0001", "zaid.jpg", {
      pendingPhotoFile: "zaid.jpg",
      messages: ["Imported.", 'Candidate Photo "zaid.jpg" — waiting for the attached photo file.'],
    }),
  ];
  const report = await uploadAttachedCandidatePhotos(rows, [file("Zaid.JPG")], upload);
  assert.deepEqual(calls, [["1602", "Zaid.JPG"]]);
  assert.deepEqual(report.rows[0].messages, ["Imported.", 'Candidate Photo uploaded from "Zaid.JPG".']);
});

test("blank cell matches a file named after the Workforce Number", async () => {
  const { calls, upload } = recorder();
  const rows = [candidate(3, "Amy Lee", "PAVE-W0002", null)];
  const report = await uploadAttachedCandidatePhotos(rows, [file("PAVE-W0002.jpg")], upload);
  assert.deepEqual(calls, [["1603", "PAVE-W0002.jpg"]]);
  assert.equal(report.uploaded, 1);
});

test("never matches by candidate name (same-name people exist)", async () => {
  const { calls, upload } = recorder();
  const rows = [
    candidate(4, "Aiden Mercer", "PAVE-W0010", null),
    candidate(5, "Aiden Mercer", "PAVE-W0011", null),
  ];
  const report = await uploadAttachedCandidatePhotos(rows, [file("Aiden Mercer.jpg")], upload);
  assert.equal(calls.length, 0);
  assert.deepEqual(report.unused, ["Aiden Mercer.jpg"]);
});

test("no photo: row is untouched and nothing is uploaded", async () => {
  const { calls, upload } = recorder();
  const rows = [candidate(6, "Ben", "PAVE-W0003", null)];
  const report = await uploadAttachedCandidatePhotos(rows, [], upload);
  assert.equal(calls.length, 0);
  assert.deepEqual(report.rows, rows);
  const outcome = withImageOutcome(
    { outcome: "completed", reason: "Import finished: 1 imported, 0 skipped." },
    report,
    "photo(s)",
    "candidate",
  );
  assert.deepEqual(outcome, { outcome: "completed", reason: "Import finished: 1 imported, 0 skipped." });
});

test("named photo not attached: candidate stays imported, row says NOT uploaded", async () => {
  const rows = [
    candidate(7, "Cara", "PAVE-W0004", "cara.png", {
      pendingPhotoFile: "cara.png",
      messages: ["Imported.", 'Candidate Photo "cara.png" — waiting for the attached photo file.'],
    }),
  ];
  const report = await uploadAttachedCandidatePhotos(rows, [], recorder().upload);
  assert.equal(report.rows[0].status, "Imported");
  assert.match(report.rows[0].messages.join(" "), /Candidate Photo "cara\.png" NOT uploaded — that file was not attached.*Upload photo on the candidate's profile/);
  const outcome = withImageOutcome({ outcome: "completed", reason: "Import finished: 1 imported, 0 skipped." }, report, "photo(s)", "candidate");
  assert.equal(outcome.outcome, "completed_with_warnings");
  assert.match(outcome.reason, /1 photo\(s\) NOT uploaded \(not attached: cara\.png\)/);
});

test("failed photo upload is reported; the candidate is still imported", async () => {
  const rows = [candidate(8, "Dan", "PAVE-W0005", "dan.jpg", { pendingPhotoFile: "dan.jpg" })];
  const report = await uploadAttachedCandidatePhotos(rows, [file("dan.jpg")], async () => {
    throw new Error("Image must be 10 MB or smaller.");
  });
  assert.equal(report.rows[0].status, "Imported");
  assert.equal(report.failed, 1);
  assert.match(report.rows[0].messages.join(" "), /Candidate Photo "dan\.jpg" upload FAILED: Image must be 10 MB or smaller\./);
});

test("skipped duplicates and errors never get a photo", async () => {
  const { calls, upload } = recorder();
  const rows = [
    { rowNumber: 9, status: "Duplicate", matchedEntityId: "77", fields: { workforceNumber: "PAVE-W0006", candidatePhoto: "e.jpg" }, messages: [] },
    { rowNumber: 10, status: "Error", fields: { workforceNumber: "PAVE-W0007" }, messages: [] },
  ];
  await uploadAttachedCandidatePhotos(rows, [file("e.jpg"), file("PAVE-W0007.jpg")], upload);
  assert.equal(calls.length, 0);
});
