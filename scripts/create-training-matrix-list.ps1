# Create SharePoint list "Training Matrix" with ALL columns from Training matrix example.xlsx
# Run inside pwsh-preview as Site Owner:
#
#   Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#   Install-Module PnP.PowerShell -Scope CurrentUser -Force
#   cd /Users/masfae.kay/Downloads/pave-training-portal-updated
#   ./scripts/create-training-matrix-list.ps1
#
# After it finishes, copy the printed List ID into .env:
#   SHAREPOINT_TRAINING_MATRIX_LIST_ID=<id>
#
# Prefer (no PowerShell): node --env-file=.env scripts/create-training-matrix-list-delegated.mjs

param(
  [string]$SiteUrl = "https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin",
  [string]$ListTitle = "Training Matrix",
  [string]$ClientId = $env:AZURE_CLIENT_ID,
  [string]$ExcelPath = (Join-Path $PSScriptRoot "..\Training matrix example.xlsx")
)

$ErrorActionPreference = "Stop"

if (-not $ClientId) {
  # Fallback to portal app id from this project's .env
  $ClientId = "e0ed9fff-e507-4c35-ad0f-6b3f1c40dc34"
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  Write-Host "Installing PnP.PowerShell…"
  Install-Module PnP.PowerShell -Scope CurrentUser -Force
}

Import-Module PnP.PowerShell

Write-Host "Connecting to $SiteUrl (device login, ClientId=$ClientId)…"
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientId

# Read headers via Excel COM if available, else expect a CSV sibling
$headers = @()
if (Test-Path $ExcelPath) {
  Write-Host "Reading columns from $ExcelPath"
  # Use ImportExcel if present; otherwise fall back to bundled CSV generated next to this script
  if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $row = Import-Excel -Path $ExcelPath -StartRow 1 -EndRow 1 -NoHeader
    # Prefer reading first worksheet header row via Open-XML-ish approach:
  }
}

$csvFallback = Join-Path $PSScriptRoot "training-matrix-example-headers.csv"
if (-not (Test-Path $csvFallback)) {
  throw "Missing $csvFallback — regenerate with the Node helper first."
}
$rows = Import-Csv $csvFallback
Write-Host "Template columns: $($rows.Count)"

$existing = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
if ($existing) {
  $stamp = Get-Date -Format "yyyyMMddHHmm"
  $ListTitle = "$ListTitle $stamp"
  Write-Host "Existing list found — creating as '$ListTitle'"
}

Write-Host "Creating list '$ListTitle'…"
$list = New-PnPList -Title $ListTitle -Template GenericList -OnQuickLaunch
Write-Host "Created. Id=$($list.Id)"

foreach ($row in $rows) {
  $header = $row.Header
  $name = $row.InternalName
  $type = $row.Type
  Write-Host "  + $header ($name)"
  if ($type -eq "Text") {
    Add-PnPField -List $ListTitle -DisplayName $header -InternalName $name -Type Text -ErrorAction Stop | Out-Null
  } else {
    Add-PnPField -List $ListTitle -DisplayName $header -InternalName $name -Type DateTime -ErrorAction Stop | Out-Null
    Set-PnPField -List $ListTitle -Identity $name -Values @{ DisplayFormat = 1 } -ErrorAction SilentlyContinue | Out-Null
  }
}

Write-Host ""
Write-Host "DONE"
Write-Host "List title : $ListTitle"
Write-Host "List ID    : $($list.Id)"
Write-Host "Put this in .env:"
Write-Host "SHAREPOINT_TRAINING_MATRIX_LIST_ID=$($list.Id)"
