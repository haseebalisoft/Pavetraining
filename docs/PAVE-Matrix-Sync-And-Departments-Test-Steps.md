# Manual test steps — Workforce ↔ Training Matrix sync + Departments Status/Notes + Candidate Categories

Three independent test packs. Part A covers the two-way Workforce ↔ Training
Matrix link. Part B covers the Departments `Status` / `Notes` wiring. Part C
covers the candidate-profile "All Training Categories" view.

Lists used: **Workforce List**, **Training Matrix Update** (the list the live
matrix UI reads), **Departments**.

---

## Part A — Workforce ↔ Training Matrix two-way sync

### Matching rule under test

A Matrix row is only ever attached to a Workforce record by one of these, in order:

1. `WorkforceItemId` (the SharePoint item id) — exact.
2. `WorkforceNumber` + `CompanyItemId` — legacy rows written before item ids.
3. Company + Candidate Name + DOB.
4. Candidate Name + DOB — **only if it matches exactly one** Workforce record.

**Candidate Name alone never links.** A row with a name but no DOB imports as
*Needs Review*; it never guesses a candidate.

### A1 — Workforce first, then Matrix (the normal path)

1. Admin → Bulk Upload → Workforce, pick `Workforce list.xlsx`, Preview.
2. Expect no Error rows. Commit.
3. Open Admin → Training Matrix. Every imported candidate has exactly one row.
4. In SharePoint, open **Training Matrix Update** and confirm each row has
   `WorkforceItemId`, `WorkforceNumber`, `CompanyItemId`, `CandidateName`, and
   `MatrixLinkStatus = Linked`.
5. Bulk Upload → Training Matrix, pick `Training matrix example.xlsx`, Preview.
   The Link column should read **Linked by Name + DOB** (or *Linked by Company +
   Name + DOB* where the sheet carries a company).
6. Commit. Row count in Training Matrix Update must be **unchanged** — matrix
   upload updates the linked rows, it does not add a second set.
7. Spot-check one candidate: their expiry dates now match the spreadsheet.

### A2 — Matrix first, then Workforce (the adoption path)

1. Start from an empty Workforce List and Training Matrix Update.
2. Bulk Upload → Training Matrix, `Training matrix example.xlsx`, Preview.
   Every data row should read **Needs Review - no Workforce match**.
3. Commit. Training Matrix Update now holds one row per candidate with
   `MatrixLinkStatus = Needs Review`, no `WorkforceItemId`, and the uploaded
   expiry dates present.
4. Log in to the **customer portal** as a customer for one of those companies.
   The Needs Review candidates must **not** appear anywhere in the matrix.
5. Bulk Upload → Workforce, `Workforce list.xlsx`, Preview. The Link column
   should read **Matrix Needs Review row will be linked**.
6. Commit, then re-open Training Matrix Update:
   - the same row ids as step 3 — **no new rows were created**;
   - each now carries `WorkforceItemId` / `WorkforceNumber` / `CompanyItemId`;
   - `MatrixLinkStatus = Linked`;
   - the expiry dates from step 3 are **still there** (adoption must not blank
     uploaded training data).
7. Customer portal: the candidates are now visible with their expiries.

### A3 — Blank cells must not erase data

1. After A1, edit `Training matrix example.xlsx` and clear one candidate's CSCS
   Expiry cell (leave it empty, don't delete the column).
2. Re-upload the Matrix sheet and commit.
3. That candidate's CSCS Expiry in Training Matrix Update must be **unchanged**.
   An empty cell means "no new value", never "delete the value".

### A4 — Same name, different DOB (must not collide)

1. Create two Workforce candidates with the same name and different DOBs, in the
   same company.
2. Each gets its own Matrix row. Confirm two rows, two different
   `WorkforceItemId`s.
3. Edit one candidate's profile. Only that candidate's Matrix row changes.

### A5 — Same name AND same DOB in two companies (needs Company)

1. Create the same name + same DOB under Company A and Company B.
2. Upload a Matrix sheet for that name/DOB **without** a Company column.
   Preview should read **Needs Review - multiple matches**, and commit must not
   attach it to either candidate.
3. Add the Company (or Company Number) column to the sheet and re-upload.
   Preview now reads **Linked by Company + Name + DOB** and the correct
   candidate's row is updated.

### A6 — Matrix row with no Workforce match at all

1. Add a row to the Matrix sheet for a candidate who does not exist in Workforce.
2. Preview: **Needs Review - no Workforce match**. Commit succeeds (not an error).
3. Admin → Training Matrix hides it until **Show all (incl. unlinked)** is on.
4. Customer portal never shows it.
5. Later create that candidate in Workforce → the existing Needs Review row is
   adopted and flips to Linked. No duplicate row.

### A7 — Rows with no candidate name are skipped, not failed

1. The supplied `Training matrix example.xlsx` has 17 trailing rows that are
   formatted but carry only a Face Fit date and no name.
2. Preview must report them as **Skipped** (not Error), with a message naming the
   values that were ignored. Total: 67 rows → 50 ready, 0 warnings, 17 skipped.

### A8 — Delete is link-aware

1. Delete a Workforce candidate who shares a name with another candidate.
2. Only that candidate's linked Matrix row is removed; the other candidate's row
   is untouched.
3. If SharePoint refuses the delete because another list still references the
   candidate, the error now **names the list, column and row ids** (for example
   `nporsRegister → CandidateName (1 row: #38)`) instead of saying "ask a Site
   Owner".

### A9 — Register save uses the same path

1. Open Admin → Registers (NPORS / EUSR / Streetworks) and save an expiry for a
   candidate.
2. That candidate's existing linked Matrix row is updated — no second row.

### Automated coverage

```bash
node --test scripts/test-workforce-matrix-sync.mjs \
             scripts/test-register-matrix-field-sync.mjs \
             scripts/test-matrix-link-scenarios.mjs
```

77 tests. `scripts/test-matrix-link-scenarios.mjs` parses the two real
spreadsheets and replays every scenario above against the real matching ladder.

---

## Part B — Departments Status + Notes

SharePoint prerequisites on the **Departments** list: `Status` (Choice: Active /
Inactive) and `Notes` (multiple lines of text).

### B1 — Create with defaults

1. Admin → Departments → Add.
2. The Status field is pre-filled **Active**; Notes is empty.
3. Save with Notes filled in. The table shows the new row with Status *Active*
   and the note.

### B2 — Edit status and notes

1. Edit the department, change Notes, save.
2. Admin → Activity Log shows a **DEPARTMENT_UPDATE** entry.
3. Edit again, set Status to **Inactive**, save.
4. Activity Log shows **DEPARTMENT_DEACTIVATE** (only on the Active → Inactive
   transition; saving an already-Inactive row logs an update, not a second
   deactivate).

### B3 — Deactivate button is the safe default

1. In the Departments table, use **Deactivate** on an Active department.
2. The row stays in the list with Status *Inactive* — nothing is deleted.
3. Activity Log shows DEPARTMENT_DEACTIVATE.
4. Hard **Delete** is still available for admins; use it on a throwaway
   department and confirm it is removed and that Workforce/Permissions rows that
   pointed at it are cleared rather than the delete failing.

### B4 — Dropdowns hide Inactive departments

1. With one department set Inactive, open Admin → Workforce → Add/Edit a
   candidate. The Department dropdown must **not** list it.
2. Admin → Permissions → Add/Edit an account. The "Departments allowed" picker
   must not list it either.
3. Open an **existing** permission whose allowed-departments already include the
   now-Inactive department and save it. The save must succeed — an existing
   assignment is never invalidated by a later deactivation.

### B5 — Admin Departments screen still shows Inactive

1. Admin → Departments lists both Active and Inactive rows with a Status column,
   so an Inactive department can be found and set back to Active.
2. Set it back to Active and confirm it reappears in the Workforce dropdown.

### B6 — Company page counts Active only

1. Note a company's department count on Admin → Companies.
2. Deactivate one of that company's departments.
3. The count drops by one.
4. Reactivate it — the count goes back up.

### B7 — Bulk workforce upload matches Active only

1. Deactivate a department that appears in the Department column of
   `Workforce list.xlsx`.
2. Bulk Upload → Workforce, Preview, with auto-create departments **off**.
   Rows for that department fail with:
   *Department "X" exists for this company but is Inactive — reactivate it under
   Admin → Departments, or enable auto-create.*
3. Turn **auto-create missing departments on** and Preview again. Rows now pass.
4. Commit, then open Admin → Departments: the department was **reactivated in
   place** — there must be exactly one row with that name for the company, not a
   second one.

### B8 — Duplicate and cap rules respect status

1. Try to create a department whose name matches an Inactive one in the same
   company. The error tells you it exists but is Inactive and to set it back to
   Active instead of duplicating.
2. With 10 Active departments on a company, creating an 11th is rejected
   ("Deactivate one first."). Deactivating one frees a slot — an Inactive
   department does not permanently consume one of the ten.

---

## Part C — Candidate profile: All Training Categories

A new **All Training Categories** table on the candidate profile (both
`/admin/workforce/<id>` and `/customer/candidates/<id>`) aggregates every
category the candidate holds across NPORS, EUSR, Streetworks/NRSWA, In-House,
NVQ and the Training Matrix's own CSCS/N-code columns into one sorted list.

### C1 — All sources appear, sorted expired-first

1. Pick a candidate with at least one record in each register (NPORS, EUSR,
   Streetworks, In-House) plus an NVQ row and a linked Training Matrix row.
2. Open their admin profile. The new **All Training Categories** section (just
   below the expiry summary cards, above the individual register tables) lists
   one row per category with Category, Source, Training Date, Expiry Date and
   Outcome.
3. Rows are ordered **Expired → Expiring soon → Compliant (Active) → Not
   applicable**, matching the same red/amber/green/grey colours used
   everywhere else (Training Matrix, register tables).
4. NVQ rows always sort into "Not applicable" (NVQ has no expiry concept) and
   show their Active/Completed status in the Outcome column instead of
   Pass/Fail.
5. Dates render as `dd/MM/yyyy`.

### C2 — Customer and admin views agree

1. Log in as a customer for that candidate's company and open the same
   candidate via `/customer/candidates/<id>`.
2. The All Training Categories table shows the **same rows** as the admin view
   (customer-visible records only — Needs Review Training Matrix rows and
   admin-only records are already filtered out upstream by the existing
   per-register fetchers, before this table ever sees them).

### C3 — No duplicate rows between registers and the Training Matrix

1. Pick a candidate whose Training Matrix row carries a CSCS expiry or an
   N-code plant expiry (e.g. N001).
2. Confirm CSCS / N-code rows appear **once**, sourced as "Training Matrix" —
   they are never repeated from a register, and a matrix column with no date
   does not produce an empty row.

### C4 — Candidate with no categories at all

1. Open the profile of a brand-new candidate with no register rows, no NVQ,
   and no linked Training Matrix row.
2. The section shows "No training categories recorded for this candidate."
   instead of an empty table.

### Automated coverage

```bash
node --import ./scripts/_register-e2e-hook.mjs --test scripts/test-candidate-categories.mjs
```

3 tests: sort order across all five sources, matrix-only categories appearing
only when dated (never duplicating a register), and the empty-candidate case.
