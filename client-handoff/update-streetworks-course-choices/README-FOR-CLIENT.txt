PAVE — Update Streetworks Course choices (SharePoint)
=====================================================

Docker is NOT required. Node.js IS required (once).

Why this is needed
------------------
The app Course list is:
  Operative
  Supervisor
  Operative Reassessment
  Supervisor Reassessment

SharePoint still has old values (Unit 1, Refresher, etc.).
When those do not match, Save can fail or the Course field will not stick.

Who runs this
-------------
A Site Owner on:
  https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin

Install Node.js first (once)
----------------------------
1. https://nodejs.org → download LTS Windows installer → install
2. Close Command Prompt, open a NEW one
3. Check:  node -v

Commands — one by one (Command Prompt)
--------------------------------------
Do NOT paste lines that start with #.
Do NOT use docker commands.

1. cd into this unzipped folder
2. npm install
3. node update-streetworks-course-choices.mjs --dry-run
4. node update-streetworks-course-choices.mjs
5. Sign in with the Site Owner Microsoft account when prompted
