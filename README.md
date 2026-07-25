# PAVE Training Portal

Next.js foundation for the PAVE Training Portal. Microsoft Entra ID handles login. SharePoint (via Microsoft Graph) remains the database. All Graph/SharePoint calls run on the server only.

## Architecture

- Frontend: Next.js App Router pages
- Auth: Auth.js (`next-auth`) + Microsoft Entra ID
- Data: SharePoint lists through server-side Graph
- Role routing: Permissions List (`UserEmail` + `Status=Active`)

## Setup

1. Copy `.env.example` to `.env.local` and fill in values.
2. Register an Entra ID app for interactive login (redirect URI: `http://localhost:3000/api/auth/callback/microsoft-entra-id`).
3. Grant the Graph app application permissions needed for SharePoint list read (for example `Sites.Read.All` or `Sites.Selected`) and admin-consent them.
4. Install and run:

```bash
npm install
npm run dev
```

## Key routes

| Route | Purpose |
| --- | --- |
| `/login` | Microsoft sign-in |
| `/` | Resolves Permissions List and routes Admin → `/admin`, Customer → `/customer` |
| `/customer` | Customer dashboard (company from Permissions only) |
| `/admin` | Admin dashboard |
| `/access-denied` | No active permission |
| `GET /api/me` | Logged-in user role/context |
| `GET /api/customer/context` | Customer context (ignores any client `companyId`) |
| `GET /api/admin/context` | Admin context |

## SharePoint schema & models

- Site: `https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin`
- Schema mapping (list names + internal fields): `src/lib/schema/sharepointSchema.ts`
- TypeScript models: `src/types/models.ts`

Backend services must import list/field names from the schema file. Do not scatter SharePoint names across components.

## Server services

- `src/lib/graph/graphClient.ts`
- `src/lib/services/sharePointListService.ts`
- `src/lib/services/permissionService.ts`
- `src/lib/services/companyService.ts`
- `src/lib/services/customerContextService.ts`
