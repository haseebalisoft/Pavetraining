# PAVE Training Portal - Review Notes

Reviewed package: pave-training-portal-updated(1).zip

## Current status found in code

Implemented / present in source:

- Microsoft login and permission based routing
- Customer home page loads the Training Matrix
- Training Matrix filters are stored in the URL
- Candidate profile route checks company and access scope before showing data
- Shared expiry status logic exists
- Customer training record views for NPORS, EUSR, Streetworks, In-House
- Customer-facing NRSWA wording is changed to Streetworks Training
- Customer NVQ page has Active and Completed tabs
- Customer documents are filtered by company, CustomerVisible, file-only, and role scope
- Admin documents browse/upload/metadata logic exists
- Bulk upload page and backend exist for Companies, Workforce/Candidates, and Training Matrix rows
- NPORS/EUSR/Streetworks/In-House/NVQ import templates exist, but full import is marked coming next
- Training Matrix register sync service exists and is hooked after admin register save
- Notifications services, settings page, audit logs page, and calendar sync service structure exist
- Outlook sync is Phase 1: SharePoint Events to Outlook Calendar

## Important cleanup

The uploaded zip included private/build files that should not be shared with the client:

- .env
- .next
- __MACOSX
- tsconfig.tsbuildinfo

Use the cleaned zip package instead of the uploaded package for handover.

## Still needs live testing

- Admin login
- Training Manager login
- Supervisor scope
- Candidate-only scope
- Murphy vs Fast company isolation
- Event company visibility
- Customer document view/download permissions
- Bulk Workforce upload with real spreadsheet
- Bulk Training Matrix upload with real spreadsheet
- Matrix sync after NPORS/EUSR/Streetworks/In-House register save
- Notification test email
- Outlook calendar sync with OUTLOOK_USER_ID and Calendars.ReadWrite permission

## Notes

This review was source/static review. Full npm build was not completed inside the sandbox because dependency installation did not complete cleanly here. Run the production build locally or on Vercel with the correct environment variables.
