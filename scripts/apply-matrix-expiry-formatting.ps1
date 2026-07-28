# Apply portal expiry colour scheme to "Training matrix example" SharePoint columns.
# Colours match the portal:
#   Grey  = missing
#   Red   = expired or urgent (within 90 days)
#   Amber = upcoming (91–270 days)
#   Green = valid (271+ days)
#
# Run in pwsh-preview as Site Owner:
#   pwsh-preview
#   Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
#   Install-Module PnP.PowerShell -Scope CurrentUser -Force
#   cd /Users/masfae.kay/Downloads/pave-training-portal-updated
#   ./scripts/apply-matrix-expiry-formatting.ps1

param(
  [string]$SiteUrl = "https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin",
  [string]$ListTitle = "Training matrix example",
  [string]$ClientId = "e0ed9fff-e507-4c35-ad0f-6b3f1c40dc34",
  [string]$FormatPath = (Join-Path $PSScriptRoot "sharepoint-matrix-expiry-column-format.json")
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell

$json = Get-Content -Raw -Path $FormatPath
Write-Host "Connecting to $SiteUrl …"
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientId

$fields = Get-PnPField -List $ListTitle
$skip = @("Title", "ID", "ContentType", "Attachments", "Edit", "LinkTitle", "LinkTitleNoMenu")
$applied = 0

foreach ($field in $fields) {
  if ($skip -contains $field.InternalName) { continue }
  if ($field.Hidden -or $field.ReadOnlyField) { continue }
  # Excel-import date columns are Number (serial) or DateTime
  if ($field.TypeAsString -notin @("Number", "Currency", "DateTime")) { continue }

  Write-Host "  format $($field.Title) ($($field.InternalName))"
  Set-PnPField -List $ListTitle -Identity $field.InternalName -Values @{
    CustomFormatter = $json
  } | Out-Null
  $applied++
}

Write-Host ""
Write-Host "DONE — applied expiry colours to $applied columns on '$ListTitle'."
Write-Host "Open the list in SharePoint and refresh to see grey/red/amber/green."
