# PAVE Training Portal — SPFx (inside SharePoint)

Same **Admin** + **Customer** portal concept as the Next.js app, running **in SharePoint** with **direct list REST** (no Next.js `/api`, no Vercel required for this path).

## Goals mapped to SPFx

| Goal | How |
|------|-----|
| Admin portal | Web part **PAVE Admin Portal** |
| Customer portal | Web part **PAVE Customer Portal** |
| Same nav / button clicks | Sidebar buttons switch views (Dashboard, Companies, Workforce, …) |
| Same design | Charcoal / lime shell (`portal.module.scss`) |
| Same lists / fields | `src/shared/schema/sharepointSchema.ts` |
| Same permission logic | Active Permissions List row; Training Manager → Admin, Supervisor → Customer |
| Skip API | `SPHttpClient` calls `/_api/web/lists/...` as the signed-in user |

## Prerequisites

- **Node.js 18.17+ or 20.11+ only** (SPFx 1.20 rejects Node 22). Prefer `20.16.0` or `20.11.0` via **fnm** (this PC does not have `nvm`).
- Site: `https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin`
- Active **Permissions List** row for each user

## Setup (step by step)

```powershell
cd spfx
# One-time per PowerShell window: enable fnm, then pick Node 20
fnm env --shell power-shell | Out-String | Invoke-Expression
fnm use 20.16.0
node -v   # must show v20.x — not v22
npm install
npx gulp trust-dev-cert
npx gulp serve
```

Workbench opens on the PAVE site. Add **PAVE Admin Portal** or **PAVE Customer Portal** from the toolbox.

## Package & deploy

```powershell
cd spfx
fnm env --shell power-shell | Out-String | Invoke-Expression
fnm use 20.16.0
npm run package
```

1. Upload `sharepoint/solution/pave-training-portal-spfx.sppkg` to the tenant **App Catalog**
2. Deploy the package
3. Create Site Pages:
   - **Admin Portal** → add **PAVE Admin Portal** (full width)
   - **Customer Portal** → add **PAVE Customer Portal** (full width)

## Shared modules (mirrors Next.js)

- `src/shared/schema/sharepointSchema.ts` — list titles + internal fields
- `src/shared/services/sharePointListService.ts` — REST list reads
- `src/shared/services/permissionService.ts` — Active permission + role normalize
- `src/shared/services/portalDataService.ts` — dashboard counts + tables
- `src/shared/ui/PortalShell.tsx` — button-click navigation + list tables
- `src/shared/ui/portal.module.scss` — PAVE charcoal / lime shell

## Security

SPFx runs as the **signed-in user**. Enforce company isolation with SharePoint permissions as well as `CustomerVisible` filters in code.

## Relation to Next.js app

The Next.js app in the repo root can stay on Vercel. This `spfx/` folder is the in-SharePoint path with the same concept and schema. Feature screens load list tables; Automation remains a stub until Outlook sync is implemented in SPFx.
