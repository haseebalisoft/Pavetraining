# PAVE Customer Training Portal — how it works (for clients)

This note explains what customers see, where new workers appear, and how security works. You can share it with Training Managers and stakeholders.

---

## What the Customer Portal is

The Customer Training Portal is a **secure website** for your company only. It is **not** your SharePoint admin site.

- Your people log in with **email + one-time code** (they do **not** need a Microsoft account in PAVE’s tenancy).
- Access is controlled by PAVE’s **Permissions list** (Training Manager / Supervisor / Candidate).
- You can **view and download** company training information. You **cannot edit** records.

---

## Where to find things (main menu)

| Menu item | Purpose |
|-----------|---------|
| **Training Matrix** (home) | Compliance overview of your candidates — search, expiry colours, “Records to Review” |
| **Dashboard** | Snapshot of counts, upcoming expiries, company profile / logo |
| **Training Delivery** | Course list / brochures (NPORS, EUSR, Streetworks, NVQ) |
| **Candidates** | Your **workforce** list — who works for your company |
| **Training Records** | NPORS Training · EUSR Training · Streetworks Training · In-House Training |
| **NVQ Progress** | NVQ stages and dates (customer-facing fields only) |
| **Documents** | Card scans & certificates you can **download** (not edit) |
| **Events / Bookings** | Upcoming visits / bookings for **your company only** |
| **Offers** | Active promotions |
| **Support** | Contact PAVE |

---

## Where a new workforce candidate must appear

When PAVE adds someone to your company on the **Workforce** list, that person should appear for your company on:

1. **Candidates** — your workforce directory (name, DOB, training manager, supervisor, key card numbers & expiries, profile link).
2. **Training Matrix** — the same people as a colour-coded compliance grid (search and expiry filters).
3. **Candidate profile** — click the person for full details and linked training / documents.

That is the core of the portal: **your workforce + your matrix**, limited to your company.

### Training records (separate step)

NPORS / EUSR / Streetworks / In-House / NVQ rows appear **after** PAVE results that training (and marks it customer-visible).  
Documents appear **after** PAVE uploads certificates or card scans into your company/candidate folders.

So:

- **New starter** → Candidates + Matrix first.  
- **After a course is resulted** → Training Records (+ matrix expiries update where linked).  
- **After a scan is uploaded** → Documents (+ optional email to Training Managers).

---

## Security (plain English)

| Rule | Meaning |
|------|---------|
| Company isolation | You never see another company’s people, matrix, or diary |
| Role scope | Training Managers (Full Company) see the whole company; Supervisors see their departments; Candidates see themselves |
| No editing | Lists are read-only for customers |
| Downloads | Certificates / card scans can be downloaded when your permission allows download |
| Remove access | PAVE can turn a Permissions row **Inactive** — history stays; portal access stops |

“Missing Data” on the admin side is shown to you as **Records to Review** — meaning something (often an expiry) needs attention, not that the system is broken.

---

## Expiry colours (matrix)

| Colour | Meaning (approx.) |
|--------|-------------------|
| Red | Expired, or due within about **3 months** |
| Amber | Due within about **3–6 months** |
| Green | **6–9 months and beyond** — compliant |
| Grey / Records to Review | No usable expiry date yet |

---

## Simple end-to-end example

1. PAVE adds your company and makes you an Active Training Manager in Permissions → you get an invite email.  
2. You log in with OTP → home is **Training Matrix**.  
3. PAVE adds candidates under your company → they appear under **Candidates** and on the **Matrix**.  
4. PAVE results an NPORS Pass → you see it under **Training Records → NPORS Training**.  
5. PAVE uploads a card scan → you see it under **Documents** and can download it; Training Managers may get an email.  
6. If PAVE sets your Permissions to Inactive → you can no longer sign in.

---

## What customers do *not* do

- Bulk upload spreadsheets (PAVE admin only)  
- Edit companies, workforce, registers, or matrix  
- See PAVE’s full operations calendar or other companies’ bookings  
- Appear as users in Microsoft Entra / Azure AD for this app  

---

## Testing checklist

A detailed row-by-row UAT sheet (developer vs client columns) lives in:

**`docs/PAVE-Customer-Portal-UAT-Checklist.csv`**

Open it in Excel. Fill:

- **Client_Tested_Yes_No** / **Client_Pass_Fail** / **Client_Notes** as you walk through each line.

Questions about access or missing people: confirm the candidate’s **company** on Admin Workforce matches the Training Manager’s **company** on Permissions.
