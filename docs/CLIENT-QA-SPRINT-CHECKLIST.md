# PAVE Client QA Sprint Checklist (11 issues)

Live app: https://pave-training-portal-nu.vercel.app  
Admin only for CRUD / bulk / delete. Refresh SharePoint-backed pages after each action (hard refresh if needed).

## 1. Workforce bulk upload — no false error / warning
- [ ] Admin → Bulk Upload → Workforce
- [ ] Preview a valid sheet (≤50 rows): Ready rows must not show Warning only for missing optional DOB
- [ ] Soft name match → Warning with exact possible-match message
- [ ] Confirm Import: toast shows created/imported, skipped, failed; on partial failure toast is error not green success
- [ ] Table Messages show real row reasons (folder created/pending/failed, matrix sync notes)
- [ ] After import: Admin Workforce + Training Matrix list refresh and show new candidates

## 2. Company email accepts .org / organisation domains
- [ ] Admin → Companies → Add
- [ ] Email `test@example.org` — form continues / save works
- [ ] Also try `accounts@charity.org.uk` and `ops@sub.company.training`
- [ ] Blank Accounts email (optional) does not block save
- [ ] Invalid `not-an-email` shows clear field error

## 3. Bulk status warnings are explained
- [ ] Any Warning row shows an exact reason in Messages (e.g. company will be created, soft name match, folder pending/failed)
- [ ] No vague bare “warning” label without text

## 4. Training Matrix bulk + UK dates
- [ ] Bulk Upload → Training Matrix: import with `06/08/2026` and `13/08/2026`
- [ ] Preview/result and Matrix UI show **dd/MM/yyyy** (6 August → `06/08/2026`)
- [ ] Invalid `08/13/2026` fails that cell/row with a clear date error — does not US-swap
- [ ] Spot-check Admin Matrix, Customer Matrix, Candidate profile, Events, Logs — dates UK

## 5. Text + status colours
- [ ] Body/table text near-black, readable
- [ ] Expiry: missing=grey, expired/0–90=red, 91–180=amber, 181+=green — high contrast badges

## 6–7. Register save sync (NPORS + others)
- [ ] Add NPORS with **Pass**, mapped N### category, and Expiry
- [ ] Toast must not claim matrix updated if sync skipped; show sync warnings
- [ ] Open Workforce candidate + Training Matrix: category/expiry updated (link by WorkforceItemId)
- [ ] Repeat for EUSR, Streetworks, In-House (Asbestos → N031), NVQ on profile
- [ ] Fail outcome must not extend expiry
- [ ] Two candidates same name different companies do not collide

## 8. Workforce bulk folders
- [ ] After Workforce bulk import, Messages: “Document folders created or already existed…” or clear pending/failed text
- [ ] SharePoint / Documents browse: Company folder → Candidates → `{WorkforceNumber} - {Name}` (+ Certificates, Card Scans, NVQ Documents, Other Documents)
- [ ] Test 1 candidate and a multi-row batch

## 9. Customer Upcoming Events duration
- [ ] Customer → Dashboard / Events
- [ ] Same-day event shows hours (e.g. 09:00–16:00 → 7 hours)
- [ ] Multi-day shows “2 days” (or N days)
- [ ] No end → “Duration not set”
- [ ] Only company-visible events

## 10. Audit / Activity Log
- [ ] Admin → Audit / Activity Log
- [ ] If SharePoint list configured: newest first; filters by date/user/action/entity work
- [ ] Create a company or run bulk upload → new log appears
- [ ] If list missing: banner says console fallback (not a silent empty page)
- [ ] SharePoint load failure shows clear error toast (not blank success)

## 11. Departments
- [ ] Admin → Departments → Add under a Company (no duplicate name in same company)
- [ ] Company filter works on Departments page
- [ ] Workforce create/edit: Department dropdown filtered by selected Company
- [ ] Permissions: Training Manager / Supervisor can use department scope

## 12. Candidate profile — all categories + expiry
- [ ] Admin and Customer candidate profile: “All training categories” table
- [ ] Columns: category, source, training date (UK), expiry + colour, outcome
- [ ] Sorted expired → expiring soon → active
- [ ] Customer only sees their company/candidate data

## 13. Non-admin blocked
- [ ] Customer / Supervisor account cannot open `/admin` or call admin APIs (access denied)

## Remaining risks
- Matrix sync still requires **Pass + expiry + mapped category** for NPORS; other outcomes intentionally skip extending dates.
- Folder creation depends on Graph Sites permissions; failures should surface in row Messages.
- Audit list requires `SHAREPOINT_TRAINING_MANAGER_LOGS_LIST_ID` on Vercel for persisted logs.
- Live “fixed” only after client retests on production with refreshed SP data.
