# PAVE Training Portal — Simple Client Test Guide

## What this is

The portal is a **web app built on top of your SharePoint site**.

- **Source of truth:** SharePoint lists and the **Customer Documents** library  
- **Site:** `https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin`  
- **Portal:** users sign in with **Microsoft 365**, then see Admin or Customer screens based on the **Permissions List**

Nothing replaces SharePoint. The portal reads and writes the same lists you already use.

---

## How login works

1. Open the portal URL (Vercel / your live link).
2. Click **Sign in with Microsoft**.
3. Use a Microsoft account that exists in the SharePoint **Permissions List**.
4. Status must be **Active**.

| Permissions List Role | Where they go | What they see |
|---|---|---|
| **Admin** / PAVE staff (Training Manager on admin side) | `/admin` | All companies, full management |
| **Training Manager** (customer) | `/customer` | Their **company** — all candidates |
| **Supervisor** | `/customer` | Only **assigned departments / candidates** |
| **Candidate** | `/customer` | **Only their own** training info |

Access scope (Company / Department / Assigned candidates) is also set on the Permissions List.

---

## How to test each role

### 1) Admin (PAVE team)

1. In SharePoint **Permissions List**, set the user’s role so they can access **Admin**.
2. Sign in → you should land on **Admin**.
3. Try:
   - Companies / Workforce  
   - Training Matrix  
   - **Customer Documents** (folder browse like SharePoint)  
   - Events / Registers  

**Admin is for managing data** that customers then view.

### 2) Training Manager (customer company)

1. Permissions List:
   - Role ≈ **Training Manager**  
   - Company = that customer company  
   - Access = company-wide  
   - Status = Active  
2. Sign in → **Customer** portal.  
3. Home page = **Training Matrix** for that company.  
4. Check candidates, documents, events, NVQ as allowed.

### 3) Supervisor

1. Permissions List:
   - Role = **Supervisor**  
   - Company = correct company  
   - Departments / assigned candidates filled in  
   - Status = Active  
2. Sign in → **Customer** portal.  
3. They should **only** see people in their department / assignment — not the whole company.

**Tip:** Use two different Microsoft accounts so Admin and Training Manager / Supervisor don’t overwrite each other in the same browser. Or use a private/incognito window.

---

## What is working now (ready to test)

- Microsoft login  
- Role-based Admin vs Customer access  
- Training Matrix as customer home (filters, expiry colours, open candidate profile)  
- Training records / registers (visible records)  
- Customer Documents (metadata + Customer Visible)  
- Admin Customer Documents folder browse matching:

```
Customer Documents
└── Company Number - Company Name
    ├── Company Documents
    └── Candidates
        └── Candidate Number - Candidate Name
            ├── Certificates
            ├── Card Scans
            ├── NVQ Documents
            └── Other Documents
```

- Document upload into the correct folder by Document Type  
- Assign Company / Candidate / Customer Visible in Admin  
- Events and offers (customer-visible items)  
- Shared expiry colours (Expired / Urgent / Upcoming / Valid / Records to Review)

---

## What is still pending / limited

| Item | Notes |
|---|---|
| **Old document folders** | Older folders like `Company - Person` may still sit at the root. New structure is used for new companies/candidates. **No automatic file move** (to protect links). Migration report only. |
| **Bulk Upload / Notifications / Settings** | Admin menu stubs — not full features yet. |
| **Events “one per company”** | SharePoint may still enforce unique Event Company — check that setting if create fails. |
| **Full UAT polish** | Mobile tweaks, more edge-case testing with live customer data. |
| **SPFx in-SharePoint shell** | Separate from the main Next.js / Vercel portal; confirm which URL the client should use for UAT. |

---

## Quick test checklist

- [ ] Admin can sign in and open all companies  
- [ ] Training Manager sees only their company matrix  
- [ ] Supervisor sees fewer candidates than Training Manager  
- [ ] Click a matrix row → candidate profile opens; Back keeps filters  
- [ ] Admin Documents: open company → Candidates → candidate → Certificates / Card Scans  
- [ ] Upload a Certificate → lands in **Certificates**; set **Customer Visible = Yes**  
- [ ] Customer Documents page shows that file (if visible)  
- [ ] Expiry colours match legend (red / amber / green / grey)

---

## Simple mental model

```
SharePoint (lists + Customer Documents)
        ↑↓
   PAVE Portal (this app)
        ↑
   Microsoft login
        ↑
 Admin  |  Training Manager  |  Supervisor  |  Candidate
```

If something is wrong in the portal, check SharePoint first: Permissions, Company, Workforce, then the list or document you expected to see.
