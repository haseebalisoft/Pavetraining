# PAVE Training Portal - Updated Package

## Matrix consistency update

- The admin, customer, sync, and dashboard matrix views now read the active SharePoint list **Training Matrix Update**.
- Admin matrix POST, PATCH, and DELETE operations now target the active wide matrix list through `SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID`.
- Legacy narrow-list IDs are rejected by the active matrix update endpoint instead of silently writing to the wrong list.
- Matrix creation prevents duplicate candidate rows in the active wide list.
- Matrix updates preserve unchanged category dates and only replace fields included in the request.
- Dashboard expiry counts and upcoming-expiry records now use the same active matrix data shown in the portal.

## Category mapping retained

- N202 and N216 are resolved by category code even when SharePoint labels differ from the template.
- The primary N016 column is preferred over the accidental `N016 ... Tonne2` duplicate.
- N020 and N114 remain blank until those columns exist in the active SharePoint list.

## Security and packaging

- Real secrets are intentionally excluded.
- `.env.example` contains the tenant, app ID, site, and SharePoint list IDs, with secret placeholders.
- Create a new Entra client secret and a new `AUTH_SECRET` before deployment.
- Generated/temporary content is excluded: `.env`, `.next`, `.npm-cache`, `node_modules`, `*.tsbuildinfo`, and generated service-worker output.

## Validation

- Changed TypeScript files passed syntax transpilation checks.
- A full clean dependency install/build could not be completed in the sandbox because the package registry returned repeated HTTP 503 responses.
