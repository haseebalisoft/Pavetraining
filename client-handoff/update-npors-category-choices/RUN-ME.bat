@echo off
setlocal
cd /d "%~dp0"

echo.
echo ============================================================
echo  PAVE — Update NPORS Category choices (SharePoint)
echo ============================================================
echo.
echo  This will NOT use Docker.
echo  You must be a SharePoint Site Owner.
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: Node.js is not installed on this PC.
  echo.
  echo 1. Open https://nodejs.org
  echo 2. Download the LTS Windows installer and install it
  echo 3. Close this window
  echo 4. Open RUN-ME.bat again
  echo.
  pause
  exit /b 1
)

echo Node found:
node -v
echo.
echo Step 1/3 — install dependencies (once^)...
echo.
call npm install
if errorlevel 1 (
  echo.
  echo npm install failed. Check internet / antivirus, then try again.
  pause
  exit /b 1
)

echo.
echo Step 2/3 — DRY RUN (safe preview, no changes^)...
echo.
pause
node update-npors-category-choices.mjs --dry-run
if errorlevel 1 (
  echo.
  echo Dry-run failed. Do not continue until this works.
  pause
  exit /b 1
)

echo.
echo Dry-run finished. If the preview looked correct, continue to APPLY.
echo Press Ctrl+C to cancel, or
pause

echo.
echo Step 3/3 — LIVE APPLY (writes to SharePoint^)...
echo.
node update-npors-category-choices.mjs
if errorlevel 1 (
  echo.
  echo Live update failed. Contact PAVE / your developer with the error above.
  pause
  exit /b 1
)

echo.
echo Done. Check SharePoint: NPORS Register → NPORS Category column choices.
echo.
pause
endlocal
