# PAVE Training Portal — Pre-Implementation Analysis

**Note on map conflicts (verified in code):** The schema-pa map claims `updateAdminCompany` "guards companyNumber edits". This is **wrong** — I read `adminCrudService.ts:299-318`: the update path runs only `requireText(input.companyNumber, "Company number")` with **no uniqueness check and no immutability guard**. The company map is correct. All other cross-map claims I spot-checked (bulkMode skipping matrix seed and folders at `adminCrudService.ts:1176-1185`, in-memory-only workforce-number clash check in bulk at `1032-1048`, workforceNumber clearable via `optionalText` in `updateAdminWorkforce`, zero code references to `CounterList`) are confirmed accurate.

---

## 1. Files that control this logic

### Company numbering & CRUD
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/companyNumber.ts` — `formatPortalCompanyNumber` (C#####), `parsePortalCompanyNumberSequence`, `allocateNextCompanyNumber` (max existing C-sequence + 1, snapshot-based, no persistent counter).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/adminCrudService.ts` — `createAdminCompany` (246-297: allocate + duplicate rejection + folder provisioning), `updateAdminCompany` (299-318: **no** number guard), `deleteAdminCompany`, `bulkDeleteAdminCompanies`, `companyWritePayload`.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/companyService.ts` — read-side mapping (`mapCompanyFields`, `getAllCompanies`).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/companyCascadeDeleteService.ts` — `deleteCompanyWithRelatedData`, `COMPANY_CASCADE_TARGETS` (12 lists + Training Manager Logs; no drive/folder cleanup, no number tombstone).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/components/admin/pages/AdminCompaniesClient.tsx` — create form pre-fills next number client-side (`getCreateDefaults`, 243-245); `companyNumber` readOnly is UI-only (line 128).

### Workforce numbering & CRUD
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/workforceNumber.ts` — `formatPortalWorkforceNumber` (W#####), `allocateNextWorkforceNumber` (same snapshot pattern).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/adminCrudService.ts` — `createAdminWorkforce` (942+: Graph clash check single / in-memory check bulk at 1032-1062, folder + matrix seed skipped in bulkMode at 1176-1185), `updateAdminWorkforce` (1239+: no number uniqueness; empty string clears WorkforceNumber), `deleteAdminWorkforce` (1385+), `ensurePermissionPerson` / `loadPermissionPeople`.

### Bulk upload
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/bulkUploadService.ts` — `previewBulkUpload` / `commitBulkUpload` dispatch; `MAX_BULK_UPLOAD_BYTES` = 15 MB (line 37); no row cap.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/candidateImporter.ts` — `previewCandidateImport`, `commitCandidateImport` (3-phase; `bulkMode: true` at 762; concurrency 5 at 421), `validateCandidateRow`, `ensureCompany` (399-418: auto-creates missing companies, `bulkMode: true` at 415).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/companyImporter.ts` — `previewCompanyImport`, `commitCompanyImport`, `findCompanyDuplicate` — **company bulk upload already exists** despite requirement 1 saying "later".
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/matching.ts` — `findCompanyByName` (name-only, Ltd/Limited-stripped), `findCandidateDuplicate` (number+company → name+DOB+company → name+company).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/parseSpreadsheet.ts`, `templates.ts`, `clientTemplateHeaders.ts` — parsing, template registry, canonical client headers.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/components/admin/pages/AdminBulkUploadClient.tsx` — posts all preview rows in one commit request (~313-321).

### Training matrix
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/trainingMatrixSyncService.ts` — `syncCandidateMatrix` / `syncCompanyMatrix` / `syncAllMatrix` / `syncAfterRegisterSave`, `shouldApplyPassExpiry`, `findMatrixRow` (name-first matching).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/trainingMatrixExampleService.ts` — `upsertTrainingMatrixExampleRow` (keyed by Title = candidate name; partial update), `listTrainingMatrixExampleRows`.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/matrixSyncHook.ts` — `triggerMatrixSyncAfterRegister` (best-effort, never fails the register save).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/bulkUpload/matrixImporter.ts` — bulk matrix spreadsheet import; refuses candidates not in Workforce.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/training/matrixManualOverrides.ts` — manual-override protection.
- `adminCrudService.ts` — `listAdminMatrix` (1694-1831: appends synthetic `workforce-only:<id>` rows), `updateAdminMatrix`, `createAdminMatrix` (writes to the **legacy** Training Matrix list — mismatched with GET).

### Documents/folders
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/customerDocumentsFolderService.ts` — `companyDocumentsFolderName` (`{CompanyNumber} - {CompanyName}`), `candidateDocumentsFolderName` (`{WorkforceNumber} - {CandidateName}`), `ensureCompanyDocumentFolders` (435-453), `ensureCandidateDocumentFolders` (460-497), `resolveChildFolderId` (277-301: `{number} -` prefix fallback makes folder identity number-stable), `resolveDocumentUploadFolder`.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/customerDocumentUploadService.ts` — `uploadCustomerDocument` (lazy folder ensure; writes `CompanyLookupId`/`CandidateLookupId`; 50 MB cap).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/customerDocumentsMigrationReportService.ts` — read-only path-vs-metadata audit ("automatic migration intentionally disabled").
- `adminCrudService.ts` — `listAdminDocumentsAtPath` (3699-3778: lazy folder ensure on browse), `parseNumberNameFolder`.

### Schema / config
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/schema/sharepointSchema.ts` — `SHAREPOINT_LISTS` single source of truth for list/field internal names (`companyNumber: 'CompanyNumber'` line 39, `workforceNumber: 'WorkforceNumber'` line 66).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/src/lib/services/sharePointListService.ts` — generic Graph CRUD by schema key; `buildSchemaFieldEqualsFilter`, `toSharePointFields`.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/pave-schema-export/sharepoint-schema.json` / `sharepoint-fields.csv` — live-tenant ground truth (proves no unique constraints; reveals dormant `CounterList`).

### Test/verification scripts (reusable)
- `/Users/masfae.kay/Desktop/Project/Pavetraining/scripts/test-company-bulk-e2e.mjs` — company bulk round-trip incl. Company Number preservation.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/scripts/test-workforce-bulk-e2e.mjs` — wipe + re-import of client Workforce list.xlsx via Graph (multi-company).
- `/Users/masfae.kay/Desktop/Project/Pavetraining/scripts/test-workforce-list-parse.mjs` — offline spreadsheet dry-parse.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/scripts/test-register-matrix-integration.mjs` — register→matrix flow.
- `/Users/masfae.kay/Desktop/Project/Pavetraining/scripts/report-customer-documents-migration.mjs`, `verify-matrix-hydrate.mjs`, `sync-training-matrix-columns.mjs`.

---

## 2. APIs that already exist

| Method | Route | Handler file | Purpose |
|---|---|---|---|
| GET | `/api/admin/companies` | `src/app/api/admin/companies/route.ts` | List companies (top 5000, name-sorted). |
| POST | `/api/admin/companies` | `src/app/api/admin/companies/route.ts` | Create one company; auto C##### allocation, duplicate rejection, folder provisioning, `folderWarning` surfaced. |
| PATCH | `/api/admin/companies/[id]` | `src/app/api/admin/companies/[id]/route.ts` | Update company — companyNumber writable, no uniqueness/immutability check. |
| DELETE | `/api/admin/companies/[id]` | `src/app/api/admin/companies/[id]/route.ts` | Cascade delete (12 lists + logs); folders and number history untouched. |
| POST | `/api/admin/companies/bulk-delete` | `src/app/api/admin/companies/bulk-delete/route.ts` | Bulk cascade delete by ids. |
| POST | `/api/admin/companies/[id]/logo` | `src/app/api/admin/companies/[id]/logo/route.ts` | Logo upload to CompanyLogo Thumbnail (10 MB). |
| GET | `/api/admin/workforce` | `src/app/api/admin/workforce/route.ts` | List workforce (optional `?companyName=`). |
| POST | `/api/admin/workforce` | `src/app/api/admin/workforce/route.ts` | Single candidate create: W##### check/allocation, folders, matrix seed; surfaces `folderWarning`/`matrixSeedWarning`. |
| PATCH | `/api/admin/workforce/[id]` | `src/app/api/admin/workforce/[id]/route.ts` | Update candidate — no number uniqueness; blank clears WorkforceNumber. |
| DELETE | `/api/admin/workforce/[id]` | `src/app/api/admin/workforce/[id]/route.ts` | Delete candidate; clears inbound lookups, removes name-matched matrix seeds. |
| POST | `/api/admin/workforce/[id]/photo` | `src/app/api/admin/workforce/[id]/photo/route.ts` | Candidate photo upload. |
| POST | `/api/admin/bulk-upload/preview` | `src/app/api/admin/bulk-upload/preview/route.ts` | Multipart preview for importType company/workforce/trainingMatrix/registers. |
| POST | `/api/admin/bulk-upload/commit` | `src/app/api/admin/bulk-upload/commit/route.ts` | Commit previewed rows, duplicateMode skip\|update\|create; maxDuration 300 s; single request, no chunking. |
| GET | `/api/admin/bulk-upload/templates` | `src/app/api/admin/bulk-upload/templates/route.ts` | Template list + xlsx/csv download (company & workforce templates exist). |
| GET/POST | `/api/admin/training-matrix` | `src/app/api/admin/training-matrix/route.ts` | GET reads Training Matrix Update + synthetic workforce-only rows; POST writes to the **legacy** list (unused by UI, mismatched with GET). |
| PATCH/DELETE | `/api/admin/training-matrix/[id]` | `src/app/api/admin/training-matrix/[id]/route.ts` | Upsert (`workforce-only:` seeds a real row) / delete matrix rows. |
| POST | `/api/admin/training-matrix/sync` (+ `/sync/candidate/[id]`, `/sync/company/[id]`) | `src/app/api/admin/training-matrix/sync/...` | Manual matrix resync (all / candidate / company), dryRun supported. |
| GET/POST, PATCH/DELETE | `/api/admin/training-records/{npors,eusr,streetworks,in-house}` (+`/[id]`) | `src/app/api/admin/training-records/...` | Register CRUD; create/update auto-sync the candidate's matrix row; DELETE does **not** sync. |
| POST | `/api/admin/documents/upload` | `src/app/api/admin/documents/upload/route.ts` | Single-file upload with lazy folder ensure + lookup metadata. |
| GET | `/api/admin/documents/browse` | `src/app/api/admin/documents/browse/route.ts` | Folder-tree browse; lazily ensures company/candidate folders. |
| GET / PATCH / DELETE | `/api/admin/documents`, `/bulk`, `/[id]`, `/[id]/download`, `/migration-report` | `src/app/api/admin/documents/...` | Metadata list/edit (files never moved), delete, download, read-only path audit. |
| GET | `/api/customer/documents` (+ `/[documentId]/download`, `/view`) | `src/app/api/customer/documents/...` | Customer-scoped document list/download/view via `CompanyLookupId` + `CustomerVisible`. |
| GET | `/api/customer/training-matrix` | `src/app/api/customer/training-matrix/route.ts` | Customer matrix view (name-keyed join). |

---

## 3. SharePoint lists/fields in use

| List | Key internal fields | Serves requirement |
|---|---|---|
| **Company List** (id 6079c6aa…) | `CompanyNumber` (**plain Text, Required=False, Indexed=False, NO EnforceUniqueValues** — `pave-schema-export/sharepoint-fields.csv:152`), `CompanyName` (required), `CompanySize`, `RegisteredAddress`, `Email`, `CompanyLogo` (Thumbnail), `Status` | Req 1, 2 |
| **Workforce List** | `WorkforceNumber` (**plain Text, not unique-enforced**), `CandidateName`, `CompanyName` (Lookup → Company, written as `CompanyNameLookupId`), projected `Company_x0020_Name_x003a__x0020_` (CompanyNumber, read-only), `Trainingmanager`/`Supervisor` (Lookups → Permissions), `Department0` (Lookup), `Email` (required), `Dateofbirth`, `CscsExpiry`, `EusrExpiry`, `SwqrExpiry`, `Photo`, `Status` | Req 3, 4, 7 |
| **Training Matrix Update** (live matrix store) | `Title` (**= candidate name — the ONLY link, no lookup, no company column**), `DOB`, `CSCSExpiry`, `EUSRExpiry`, `NRSWAExpiry`, dynamic `N###…Expiry` columns, `ManualOverrides`, optional `MatrixSyncedAt`/`AutomationStatus` | Req 5, 7 (weakly) |
| **Training Matrix** (legacy) | `CandidateName` (Lookup → Workforce), `MatrixCompany` (Lookup → Company), `N001Expiry…N100Expiry` | Written only by unused POST route |
| **Training Matrix Category Records** | `Candidate_x0020_Name`, `Company_x0020_Name`, `Category_x0020_Code`, `Expiry_x0020_Date` | Backup rows on matrix bulk import |
| **NPORS / EUSR / NRSWA / In-House / NVQ registers** | `CandidateName` + `CompanyName` Lookups, `TrainingOutcome`, `Expiry`/`Expirydate`/`ExpiryDate`, category columns | Req 5, 7 (matrix sync inputs) |
| **Customer Documents** (library) | `Company`/`CompanyLookupId`, `Candidate`/`CandidateLookupId`, `DocumentType`, `CustomerVisible`, `FileRef`/`FileDirRef`/`FSObjType`; folder convention `{C##### - Name}/{Company Documents\|Candidates/{W##### - Name}/{Certificates\|Card Scans\|NVQ Documents\|Other Documents}}` | Req 6, 7 |
| **Permissions List**, **Departments** | `Name`, `UserEmail`, `RoleType`, Company lookups; `Department0LookupId` target | Req 3 (lookup targets auto-created during import) |
| **Training Manager Logs** | audit rows incl. bulk-upload logging | Cross-cutting |
| **CounterList** (live, 132 items) | `CompanyCounter`, `WorkForceCounter` (Number) | **Referenced by ZERO code** (verified: grep across src/, scripts/, spfx/ returns nothing). Likely the data store of a retired out-of-repo Power Automate numbering flow. |

**Power Automate: explicitly — no Power Automate fields, flags, or integrations are used by any code in this repo.** Verified greps for "power automate", flow triggers, CounterList return zero code hits. The only Flow residue is on the live tenant: hidden built-ins (`TriggerFlowInfo`, `WorkflowVersion`) plus two stray unused Text columns (`Departments.flowInvolve`, `Permissions List.flowCreation`) and the orphaned `CounterList`. All numbering, folder creation, and matrix sync are app-side (Node/Graph). If a live Flow still increments CounterList, it is completely uncoordinated with the app's allocators — this must be confirmed dead before relying on requirement 2/4 guarantees.

---

## 4. What is missing — requirement-by-requirement

**Req 1 — Companies one at a time; bulk later: EXISTS (and bulk is already built).**
Single create: `POST /api/admin/companies` → `createAdminCompany` with auto-numbering, duplicate rejection, folder provisioning. The "later" bulk piece already exists too: `companyImporter.ts` + `Company-list-template.xlsx` + `scripts/test-company-bulk-e2e.mjs`. No build work; only hardening.

**Req 2 — Company numbers never lost/overwritten/duplicated: PARTIAL.**
- Duplicate check on create exists (`adminCrudService.ts:255-265`) but is app-side read-then-write with no lock — concurrent creates can both pass (no SharePoint EnforceUniqueValues per schema export).
- **Overwritable:** `updateAdminCompany` (299-318) accepts any `companyNumber` with only `requireText` — verified in code; UI readOnly is the only guard.
- **Reusable after delete:** no tombstone/registry; `allocateNextCompanyNumber` recomputes from the live list, so deleting the highest-numbered company re-issues its number, and the recreated company inherits the orphaned `{C##### - OldCo}` folder via the prefix match in `resolveChildFolderId`.

**Req 3 — Workforce bulk up to 50 records, multi-company, auto-assigned: PARTIAL (mostly EXISTS).**
Multi-company per file works: per-row `findCompanyByName` + `ensureCompany` auto-create (`candidateImporter.ts:399-418`). But: (a) no 50-row cap or chunking anywhere (only 15 MB / 300 s); (b) assignment is by fuzzy **name** only — the sheet's Company Number column is never used to disambiguate or verify an existing company, so spelling variants beyond Ltd/Limited silently create a duplicate company with a new number.

**Req 4 — Workforce numbers maintained: PARTIAL.**
Allocation + create-time clash checks exist (`adminCrudService.ts:1032-1067`). Missing: `updateAdminWorkforce` has no uniqueness check and empty string **clears** WorkforceNumber (verified, 1246-1251); intra-file explicit duplicate numbers can slip through commit (validation runs against pre-create state, known-numbers snapshot frozen at `candidateImporter.ts:748` before the 5-way-parallel create); no persistent counter (snapshot race across concurrent commits).

**Req 5 — Workforce auto-appears in Training Matrix with profile details: PARTIAL — MISSING for bulk.**
Single create seeds a Training Matrix Update row (`adminCrudService.ts:1201-1227`). **Bulk import never does**: `commitCandidateImport` always passes `bulkMode: true` (`candidateImporter.ts:762`) which skips the seed (`adminCrudService.ts:1181-1185`, verified), and there is no post-commit backfill. Bulk candidates appear in the matrix UI only as synthetic `workforce-only:` rows (no SharePoint row) until first edit/register-save/sync. Profile details in the matrix are joined at read time by candidate **name** — fragile (see Req 7).

**Req 6 — Folders auto-created: PARTIAL — MISSING for bulk.**
Single company/workforce create eagerly provisions the correct tree. Bulk (both company-auto-create during workforce import and workforce rows themselves) skips folders (`candidateImporter.ts:415, 762`; `adminCrudService.ts:275-283, 1176-1180`); folders appear only lazily on admin browse (`listAdminDocumentsAtPath:3712-3730`) or first upload. Note: `commitCompanyImport` (dedicated company bulk) **does** create folders per company — only the workforce-import path skips them.

**Req 7 — Everything stays linked: PARTIAL.**
Strong: Workforce↔Company, registers, documents all use SharePoint Lookup IDs (`CompanyNameLookupId`, `CandidateLookupId`, etc.); folder identity is number-prefix-stable. Weak: the Training Matrix Update list stores **only Title = candidate name** — no lookup, no company column, no workforce number — so same-named candidates across companies collide onto one row and a rename orphans the row (no matrix maintenance in `updateAdminWorkforce`, verified none). Also: company delete orphans the folder tree (no drive deletes anywhere in code); metadata edits never move files; register DELETE never rolls back matrix dates.

---

## 5. Safest implementation plan

Order chosen so each step is small, independently shippable, and protects data integrity before adding features.

**Step 1 — Server-side number immutability/uniqueness on update (smallest, highest value).**
Extend `updateAdminCompany` (`adminCrudService.ts:299-318`) and `updateAdminWorkforce` (1239+): if `companyNumber`/`workforceNumber` is supplied and differs from `existing`, reject (or, if change must be allowed, enforce uniqueness with the same scan/Graph-filter used on create — `buildSchemaFieldEqualsFilter` already exists). Also reject empty-string clears of `workforceNumber`. No new files.
*Test:* curl PATCH `/api/admin/companies/[id]` and `/api/admin/workforce/[id]` attempting number change/clear/duplicate; confirm 400s; run `scripts/test-company-bulk-e2e.mjs` and `scripts/test-workforce-bulk-e2e.mjs` to prove imports unaffected.

**Step 2 — Number tombstones so deleted numbers are never reissued.**
Smallest safe design: in `deleteCompanyWithRelatedData` (`companyCascadeDeleteService.ts`) and `deleteAdminWorkforce`, record the freed number (new small SharePoint list, e.g. "Retired Numbers", added to `sharepointSchema.ts` — or reconcile/repurpose the dormant live `CounterList` after confirming no Flow touches it). Extend `allocateNextCompanyNumber`/`allocateNextWorkforceNumber` call sites to pass retired numbers via the existing `extraUsed` parameter — the pure functions already support this, so **no allocator change is needed**, only the callers in `createAdminCompany`, `createAdminWorkforce`, `companyImporter.ts`, `candidateImporter.ts`.
*Test:* create → delete → create; assert new company gets a fresh number, not the deleted one; verify no orphaned-folder inheritance.

**Step 3 — Enforce the 50-row batch cap.**
Add a `MAX_BULK_ROWS = 50` (or configurable) check in `previewBulkUpload`/`commitBulkUpload` (`bulkUploadService.ts`) with a clear error, and mirror it in `AdminBulkUploadClient.tsx` before posting. Trivial, prevents the 300 s-timeout partial-failure mode.
*Test:* `scripts/test-workforce-list-parse.mjs` on the real client sheet; preview a 51-row file → expect rejection; 50-row file → passes.

**Step 4 — Close intra-file explicit-duplicate workforce numbers.**
In `commitCandidateImport`, before the parallel create phase (`candidateImporter.ts:748`), scan the batch for repeated explicit `workforceNumber` values (and values colliding with `allocatedWorkforceNumbers`) and mark repeats as errors; alternatively add each created number to a shared mutable set inside `mapPool` workers. Extend existing functions; no new files.
*Test:* craft a 3-row CSV repeating one explicit W-number across two companies; preview+commit; assert exactly one create.

**Step 5 — Post-bulk matrix seeding (Req 5).**
Do **not** flip `bulkMode` off (it exists to keep 50-row imports fast). Instead add a post-commit phase in `commitCandidateImport`: after Phase 3, loop created candidates and call the existing `upsertTrainingMatrixExampleRow` (same payload shape as `createAdminWorkforce:1204-1218` — Name, DOB, CSCS/EUSR/NRSWA expiry), best-effort with warnings collected into the result (mirroring `matrixSeedWarning`). Consider bounded concurrency via the existing `mapPool`.
*Test:* run bulk import of a small multi-company file, then GET `/api/admin/training-matrix` and assert rows are real `example:` rows, not `workforce-only:` synthetics; `scripts/verify-matrix-hydrate.mjs` for hydration; `scripts/test-register-matrix-integration.mjs` for regression.

**Step 6 — Post-bulk folder provisioning (Req 6).**
Same pattern: after commit, call existing `ensureCompanyDocumentFolders` for each auto-created company and `ensureCandidateDocumentFolders` for each created candidate (both already return `{ok, warning}` instead of throwing). Run serially or at low concurrency to respect Graph throttling; surface warnings in the commit summary. If 50 candidates × ~6 folder calls threatens the 300 s window, make it a follow-up endpoint (`POST /api/admin/bulk-upload/provision-folders`) triggered by the UI after commit — new route, reusing the same two functions.
*Test:* bulk-import then GET `/api/admin/documents/browse?path=...` WITHOUT relying on the lazy ensure (check drive children directly or via `scripts/report-customer-documents-migration.mjs`); re-run `test-workforce-bulk-e2e.mjs`.

**Step 7 — Harden matrix linkage (Req 7, larger change — schedule last).**
Add a `WorkforceLookupId` (or at minimum a `WorkforceNumber` text column) to the Training Matrix Update list (provision via a script modeled on `scripts/sync-training-matrix-columns.mjs` / `ensure-training-matrix-update-columns.mjs`), write it in `upsertTrainingMatrixExampleRow`, and prefer it over name in `findMatrixRow` (`trainingMatrixSyncService.ts:284-299`) and the read-time joins (`listAdminMatrix`, `getCustomerMatrixRecords`). Also add matrix-row maintenance on candidate rename in `updateAdminWorkforce`. Keep name matching as fallback for legacy rows.
*Test:* dry-run column script; rename a candidate and confirm matrix row follows; two same-named candidates in different companies get distinct rows; full `POST /api/admin/training-matrix/sync` with `dryRun: true` before/after.

**Deliberately deferred / flagged:** retiring the mismatched `createAdminMatrix` POST route (writes to legacy list; unused by UI — either delete or repoint at Training Matrix Update), folder cleanup/archival on company delete, and file re-filing on metadata change (migration is intentionally disabled — leave as-is until URL impact reviewed).

---

## 6. Risks before coding

1. **Number-uniqueness race on concurrent creates.** All duplicate protection is read-list-then-write in Node with no lock/etag and no SharePoint EnforceUniqueValues (`adminCrudService.ts:246-271` for companies; `1032-1067` for workforce; schema export confirms plain Text columns). Two admins, or a company bulk + workforce bulk running simultaneously, can both allocate the same C#####/W#####. Any fix short of a transactional counter only narrows the window — document this residual risk.

2. **Dormant live `CounterList` may still be fed by an out-of-repo Power Automate flow.** 132 items with `CompanyCounter`/`WorkForceCounter` (`pave-schema-export/sharepoint-lists.csv`), zero code references (verified). If a Flow still increments it, its numbering diverges from the app's snapshot allocators. Confirm the Flow is dead in the tenant **before** promising "never duplicated", and before repurposing the list in Step 2.

3. **Deleted numbers are silently reissued and inherit orphaned folders.** `deleteCompanyWithRelatedData` (`companyCascadeDeleteService.ts:275-361`) records nothing and never touches drive folders; `resolveChildFolderId`'s `{number} -` prefix fallback (`customerDocumentsFolderService.ts:277-301`) then attaches the reused number's new company to the old company's folder tree — cross-company document leakage. Step 2 (tombstones) must land before any delete/recreate happens in production data.

4. **Company matching in workforce bulk is name-only and fuzzy.** `findCompanyByName` (`matching.ts:25-39`) strips only Ltd/Limited/punctuation; the sheet's Company Number column is ignored for matching (`candidateImporter.ts:408-412`, used only on create). "P.A.V.E. Civils" vs "PAVE Civils Ltd" creates a duplicate company with a fresh number and splits the workforce across two records. Consider matching by Company Number first (mirroring `findCompanyDuplicate` in `companyImporter.ts:187-202`) as part of Step 3/4.

5. **PATCH endpoints currently allow overwriting/clearing numbers.** `updateAdminCompany` (`adminCrudService.ts:309-311`, verified) and `updateAdminWorkforce` (`1246-1251`) — protection is UI-only (`AdminCompaniesClient.tsx:128`). Any script, integration, or future UI change can violate requirement 2/4 today. Fix first (Step 1).

6. **Graph throttling / 300 s timeout on enriched 50-row batches.** Commit is one request (`AdminBulkUploadClient.tsx:313-321`, `commit/route.ts:8` maxDuration 300). The code comment at `adminCrudService.ts:1032` records that per-row Graph checks already made 50-row imports take ~100 s; adding per-candidate matrix upserts (Step 5) and ~6 folder calls per candidate (Step 6) could breach the window or trip 429s. Mitigate with `mapPool` bounded concurrency, or move folder provisioning to a follow-up call; company bulk commit is worse — it re-fetches the full company list after every row (`companyImporter.ts:415, 440, 459`), O(N²).

7. **Frozen snapshots inside one bulk commit.** `knownWorkforceNumbers` is captured once before the 5-way-parallel create phase (`candidateImporter.ts:748`), so intra-file explicit duplicates bypass the clash check entirely (Step 4). The same pattern means preview-assigned numbers can differ from committed ones if data changed between preview and commit — set expectations in the UI.

8. **Name-keyed matrix rows break "stay linked".** Training Matrix Update has no lookup/company/number column (`adminCrudService.ts:1714-1715`; `trainingMatrixExampleService.ts:358-380`); `findMatrixRow` prefers exact-name match (`trainingMatrixSyncService.ts:284-299`). Two same-named candidates collide onto one row; renames orphan rows (no matrix code in `updateAdminWorkforce`, verified). Step 5 built on top of this inherits the fragility until Step 7 lands — seed order matters.

9. **Mismatched matrix POST route can corrupt expectations.** `POST /api/admin/training-matrix` writes to the legacy Training Matrix list while GET reads Training Matrix Update (`adminCrudService.ts:1888-1891` vs `1694+`). Any new automation that "creates a matrix row" via this route will produce invisible rows. Retire or repoint before Step 5.

10. **Folder rename fallback fails when the number is empty.** `companyDocumentsFolderName` falls back to name-only when `companyNumber` is empty (`customerDocumentsFolderService.ts:51-58`), and the prefix fallback in `resolveChildFolderId` returns null without a number — a rename of a number-less (legacy) company silently forks a second folder tree. Audit legacy rows with empty `CompanyNumber` before Step 6's backfill.

11. **Upload metadata write is not atomic.** `uploadCustomerDocument` uploads bytes first, patches `CompanyLookupId` after (`customerDocumentUploadService.ts:196-245`); a failure between leaves an orphan file invisible to the customer portal and unmatchable by cascade delete. Any bulk folder/document work should not assume metadata is always present.

12. **This project's Next.js is a modified version.** Per `AGENTS.md`, APIs/conventions may differ from standard Next.js — read `node_modules/next/dist/docs/` before touching any route handler files (all Steps that add/modify routes: 3, 6, 9).
