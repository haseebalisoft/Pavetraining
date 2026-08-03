PAVE Training — Update NPORS Category choices in SharePoint
==========================================================

Send / zip this whole folder to a Site Owner.
They do NOT need Docker. They do NOT edit SharePoint by hand.

What it does
------------
Updates list "NPORS Register" column "NPORS Category" with the codes in:
  NPORS-Category-SharePoint-Choices.txt

Who can run it
--------------
Someone who is a Site Owner on:
  https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin

IMPORTANT — install Node.js first (once only)
---------------------------------------------
Your PC needs Node.js. Docker is NOT required.

1. Open: https://nodejs.org
2. Download the LTS Windows installer (.msi)
3. Run the installer → Next / Next / Install (keep default options)
4. CLOSE Command Prompt completely, then open a NEW Command Prompt
5. Check it worked (type this exactly, then press Enter):

   node -v

   You should see something like: v22.x.x   (or v20 / v24)

   If you still see "'node' is not recognized", restart the PC once
   and try "node -v" again in a new Command Prompt.

Easiest way to run (Windows)
----------------------------
1. Unzip this folder to Desktop (not inside OneDrive sync if possible)
2. Double-click:  RUN-ME.bat
3. Follow the on-screen prompts (dry-run first, then live apply)
4. When asked, open the Microsoft login URL, enter the code, sign in
   as a Site Owner

Manual commands — one by one (Command Prompt)
---------------------------------------------
Do NOT paste lines that start with #.
Do NOT use Docker commands.

1. Open Command Prompt

2. Go into this folder (change the path if yours is different):

   cd /d "%USERPROFILE%\Desktop\PAVE-Update-NPORS-Category-Choices"

3. Install dependencies (once — needs internet):

   npm install

4. SAFE preview (no SharePoint changes):

   node update-npors-category-choices.mjs --dry-run

   → Sign in when prompted. Confirm it lists ~118 codes and shows what
     would change. Close when done.

5. APPLY for real:

   node update-npors-category-choices.mjs

   → Sign in again as Site Owner. Wait until it says the update succeeded.

6. Check in SharePoint:
   NPORS Register → list settings → NPORS Category column → many N### choices.

Files in this pack
------------------
  RUN-ME.bat                            ← double-click helper (Windows)
  update-npors-category-choices.mjs     ← the script
  NPORS-Category-SharePoint-Choices.txt ← the codes
  package.json                          ← npm dependencies
  README-FOR-CLIENT.txt                 ← this file

No app secrets are included. Sign-in is interactive (device code).

If something fails
------------------
- "'node' is not recognized" → Node not installed / old terminal still open.
  Install from nodejs.org, close ALL Command Prompt windows, open a new one.
- "'npm' is not recognized" → same as above (npm comes with Node).
- Sign-in / 403 error → account is not a Site Owner on that SharePoint site.
- Do not follow Docker install guides for this pack.
