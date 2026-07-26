param(
    [string]$SiteUrl = "https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin",
    [string]$ClientId = "56e8a0ec-eafa-4a2b-96f3-7254baad746f",
    [switch]$WhatIfMode,
    [switch]$CreateTestEvent,
    [string]$TestCompanyName = "Murphy plant"
)

$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Skip {
    param([string]$Message)
    Write-Host "[SKIP] $Message" -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function Get-ListSafe {
    param([string]$ListTitle)

    try {
        return Get-PnPList -Identity $ListTitle -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-FieldSafe {
    param(
        [string]$ListTitle,
        [string]$InternalName
    )

    try {
        return Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Ensure-EventsList {
    $eventsList = Get-ListSafe -ListTitle "Events"

    if ($eventsList) {
        Write-Ok "Events list already exists."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would create Events calendar list."
        return
    }

    Write-Info "Creating Events calendar list..."
    New-PnPList -Title "Events" -Template Events -EnableVersioning | Out-Null
    Write-Ok "Events calendar list created."
}

function Ensure-TextField {
    param(
        [string]$ListTitle,
        [string]$InternalName,
        [string]$DisplayName
    )

    $field = Get-FieldSafe -ListTitle $ListTitle -InternalName $InternalName

    if ($field) {
        Write-Skip "$ListTitle already has $InternalName."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would add text field $InternalName to $ListTitle."
        return
    }

    Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Text -AddToDefaultView | Out-Null
    Write-Ok "Added text field $InternalName to $ListTitle."
}

function Ensure-NoteField {
    param(
        [string]$ListTitle,
        [string]$InternalName,
        [string]$DisplayName
    )

    $field = Get-FieldSafe -ListTitle $ListTitle -InternalName $InternalName

    if ($field) {
        Write-Skip "$ListTitle already has $InternalName."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would add note field $InternalName to $ListTitle."
        return
    }

    Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type Note -AddToDefaultView | Out-Null
    Write-Ok "Added note field $InternalName to $ListTitle."
}

function Ensure-DateTimeField {
    param(
        [string]$ListTitle,
        [string]$InternalName,
        [string]$DisplayName
    )

    $field = Get-FieldSafe -ListTitle $ListTitle -InternalName $InternalName

    if ($field) {
        Write-Skip "$ListTitle already has $InternalName."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would add datetime field $InternalName to $ListTitle."
        return
    }

    Add-PnPField -List $ListTitle -DisplayName $DisplayName -InternalName $InternalName -Type DateTime -AddToDefaultView | Out-Null
    Write-Ok "Added datetime field $InternalName to $ListTitle."
}

function Ensure-BooleanFieldFromXml {
    param(
        [string]$ListTitle,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$DefaultValue = "0"
    )

    $field = Get-FieldSafe -ListTitle $ListTitle -InternalName $InternalName

    if ($field) {
        Write-Skip "$ListTitle already has $InternalName."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would add boolean field $InternalName to $ListTitle."
        return
    }

    $xml = "<Field Type='Boolean' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName'><Default>$DefaultValue</Default></Field>"

    Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml | Out-Null
    Write-Ok "Added boolean field $InternalName to $ListTitle."
}

function Ensure-ChoiceFieldFromXml {
    param(
        [string]$ListTitle,
        [string]$InternalName,
        [string]$DisplayName,
        [string[]]$Choices,
        [string]$DefaultChoice
    )

    $field = Get-FieldSafe -ListTitle $ListTitle -InternalName $InternalName

    if ($field) {
        Write-Skip "$ListTitle already has $InternalName."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would add choice field $InternalName to $ListTitle."
        return
    }

    $choicesXml = ""
    foreach ($choice in $Choices) {
        $choicesXml += "<CHOICE>$choice</CHOICE>"
    }

    $xml = @"
<Field Type='Choice' DisplayName='$DisplayName' Name='$InternalName' StaticName='$InternalName' Format='Dropdown'>
  <CHOICES>
    $choicesXml
  </CHOICES>
  <Default>$DefaultChoice</Default>
</Field>
"@

    Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml | Out-Null
    Write-Ok "Added choice field $InternalName to $ListTitle."
}

function Ensure-EventCompanyLookup {
    $listTitle = "Events"
    $internalName = "EventCompany"

    $field = Get-FieldSafe -ListTitle $listTitle -InternalName $internalName

    if ($field) {
        Write-Skip "Events already has EventCompany."
        # Unique values must be OFF — otherwise only one event per company can be saved.
        try {
            if ($field.EnforceUniqueValues -eq $true) {
                if ($WhatIfMode) {
                    Write-Skip "Would disable EnforceUniqueValues on EventCompany."
                }
                else {
                    Set-PnPField -List $listTitle -Identity $internalName -Values @{ EnforceUniqueValues = $false; Indexed = $true } -ErrorAction Stop
                    Write-Ok "Disabled Enforce unique values on EventCompany (required for multiple events per company)."
                }
            }
        }
        catch {
            Write-Err "Could not disable EnforceUniqueValues on EventCompany: $($_.Exception.Message). Do this manually in SharePoint column settings."
        }
        return
    }

    $companyList = Get-ListSafe -ListTitle "Company List"

    if (-not $companyList) {
        Write-Err "Company List does not exist. Cannot create EventCompany lookup."
        return
    }

    $companyNameField = Get-FieldSafe -ListTitle "Company List" -InternalName "CompanyName"
    $showField = "Title"

    if ($companyNameField) {
        $showField = "CompanyName"
    }

    if ($WhatIfMode) {
        Write-Skip "Would add EventCompany lookup to Events using Company List field $showField."
        return
    }

    $companyListId = $companyList.Id.ToString("B")

    $xml = "<Field Type='Lookup' DisplayName='Event Company' Name='EventCompany' StaticName='EventCompany' List='$companyListId' ShowField='$showField' EnforceUniqueValues='FALSE' Indexed='TRUE' />"

    Add-PnPFieldFromXml -List "Events" -FieldXml $xml | Out-Null
    Write-Ok "Added EventCompany lookup to Events."
}

function Ensure-CoreEventsFields {
    $listTitle = "Events"

    Ensure-DateTimeField -ListTitle $listTitle -InternalName "EventDate" -DisplayName "Start Time"
    Ensure-DateTimeField -ListTitle $listTitle -InternalName "EndDate" -DisplayName "End Time"
    Ensure-TextField -ListTitle $listTitle -InternalName "Location" -DisplayName "Location"
    Ensure-NoteField -ListTitle $listTitle -InternalName "Description" -DisplayName "Description"
    Ensure-NoteField -ListTitle $listTitle -InternalName "TrainingAddress" -DisplayName "Training Address"

    $customerVisibleEncoded = Get-FieldSafe -ListTitle $listTitle -InternalName "Customer_x0020_Visible"
    $customerVisibleClean = Get-FieldSafe -ListTitle $listTitle -InternalName "CustomerVisible"

    if ($customerVisibleEncoded) {
        Write-Skip "Events already has Customer_x0020_Visible."
    }
    elseif ($customerVisibleClean) {
        Write-Skip "Events has CustomerVisible instead of Customer_x0020_Visible. Use CustomerVisible in app schema or create encoded field manually."
    }
    else {
        Ensure-BooleanFieldFromXml -ListTitle $listTitle -InternalName "Customer_x0020_Visible" -DisplayName "Customer Visible" -DefaultValue "0"
    }

    Ensure-EventCompanyLookup
}

function Ensure-CalendarSyncFields {
    $listTitle = "Events"

    Ensure-TextField -ListTitle $listTitle -InternalName "OutlookEventId" -DisplayName "Outlook Event ID"
    Ensure-TextField -ListTitle $listTitle -InternalName "OutlookCalendarId" -DisplayName "Outlook Calendar ID"
    Ensure-TextField -ListTitle $listTitle -InternalName "OutlookICalUid" -DisplayName "Outlook iCal UID"

    Ensure-ChoiceFieldFromXml `
        -ListTitle $listTitle `
        -InternalName "SyncStatus" `
        -DisplayName "Sync Status" `
        -Choices @("Not Synced", "Pending", "Synced", "Failed", "Skipped") `
        -DefaultChoice "Not Synced"

    Ensure-ChoiceFieldFromXml `
        -ListTitle $listTitle `
        -InternalName "SyncDirection" `
        -DisplayName "Sync Direction" `
        -Choices @("None", "SharePointToOutlook", "OutlookToSharePoint", "TwoWay") `
        -DefaultChoice "SharePointToOutlook"

    Ensure-DateTimeField -ListTitle $listTitle -InternalName "LastSyncedAt" -DisplayName "Last Synced At"

    Ensure-ChoiceFieldFromXml `
        -ListTitle $listTitle `
        -InternalName "LastSyncSource" `
        -DisplayName "Last Sync Source" `
        -Choices @("None", "SharePoint", "Outlook", "System") `
        -DefaultChoice "None"

    Ensure-NoteField -ListTitle $listTitle -InternalName "SyncError" -DisplayName "Sync Error"
    Ensure-BooleanFieldFromXml -ListTitle $listTitle -InternalName "DoNotSync" -DisplayName "Do Not Sync" -DefaultValue "0"
    Ensure-TextField -ListTitle $listTitle -InternalName "SyncHash" -DisplayName "Sync Hash"
}

function Create-TestEvent {
    if (-not $CreateTestEvent) {
        return
    }

    $companyItems = Get-PnPListItem -List "Company List" -PageSize 500 -Fields "Title", "CompanyName"

    $companyItem = $companyItems | Where-Object {
        $_["CompanyName"] -eq $TestCompanyName -or $_["Title"] -eq $TestCompanyName
    } | Select-Object -First 1

    if (-not $companyItem) {
        Write-Err "Test company '$TestCompanyName' not found in Company List. Test event not created."
        return
    }

    if ($WhatIfMode) {
        Write-Skip "Would create test event for company $TestCompanyName."
        return
    }

    $start = (Get-Date).AddDays(7).Date.AddHours(10)
    $end = $start.AddHours(2)

    $values = @{
        "Title" = "Test Training Event"
        "EventDate" = $start
        "EndDate" = $end
        "Location" = "Test Location"
        "TrainingAddress" = "Test Training Address"
        "EventCompany" = $companyItem.Id
        "Customer_x0020_Visible" = $true
        "Description" = "Test event created for portal testing."
        "SyncStatus" = "Not Synced"
        "SyncDirection" = "SharePointToOutlook"
        "LastSyncSource" = "SharePoint"
        "DoNotSync" = $true
    }

    Add-PnPListItem -List "Events" -Values $values | Out-Null
    Write-Ok "Created test event for $TestCompanyName."
}

Write-Info "Connecting to SharePoint..."
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
Write-Ok "Connected."

Ensure-EventsList
Ensure-CoreEventsFields
Ensure-CalendarSyncFields
Create-TestEvent

Write-Host ""
Write-Ok "Events schema check/update completed."

Write-Host ""
Write-Host "Run this to verify:" -ForegroundColor Cyan
Write-Host 'Get-PnPField -List "Events" | Where-Object { $_.InternalName -in @("Title","EventDate","EndDate","Location","Description","TrainingAddress","EventCompany","Customer_x0020_Visible","OutlookEventId","OutlookCalendarId","OutlookICalUid","SyncStatus","SyncDirection","LastSyncedAt","LastSyncSource","SyncError","DoNotSync","SyncHash") } | Select Title, InternalName, TypeAsString' -ForegroundColor Gray