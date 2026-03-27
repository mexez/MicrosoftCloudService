#### MAIN SCRIPT
## SharePoint file access audit on a specific site


# ================================================================
# SharePoint File Access Status Audit – Accounting Site
# ================================================================

function Clean-Text($s) {
    if ($null -eq $s) { return "None" }
    return $s.ToString() -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", ""
}

$tenantId  = "xxx"
$clientId  = "yyy"
$certPath  = "C:\Users\bcadmin\Documents\e.pfx"
$siteUrl   = "https://domain.sharepoint.com/sites/accounting"
$excelPath = "C:\Users\bcadmin\Downloads\Reports\DAReportSheet\zz.xlsx"

$certPassword = Read-Host "Enter PFX password" -AsSecureString

Connect-PnPOnline `
  -Url $siteUrl `
  -Tenant $tenantId `
  -ClientId $clientId `
  -CertificatePath $certPath `
  -CertificatePassword $certPassword

# ------------------------------------------------
# Collect site-level groups for inheritance context
# ------------------------------------------------

$siteGroups = Get-PnPGroup | Where-Object { $_.Title -match "Owners|Members|Visitors" }
$groupNames = ($siteGroups.Title -join "; ")

# ------------------------------------------------
# File / Folder Access Status Sheet
# ------------------------------------------------

$fileAccessStatus = @()

$lists = Get-PnPList | Where-Object {
    $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false
}

foreach ($list in $lists) {

    Write-Host "Scanning library:" $list.Title -ForegroundColor Yellow

    $query = @"
<View Scope='RecursiveAll'>
  <ViewFields>
    <FieldRef Name='FileRef'/>
    <FieldRef Name='FileSystemObjectType'/>
    <FieldRef Name='HasUniqueRoleAssignments'/>
  </ViewFields>
  <RowLimit>200</RowLimit>
</View>
"@

    $items = Get-PnPListItem -List $list -Query $query

    foreach ($item in $items) {

        $path = $item["FileRef"]
        if (-not $path) { continue }

        $itemType = if ($item["FileSystemObjectType"] -eq 1) { "Folder" } else { "File" }

        $inheritance = if ($item["HasUniqueRoleAssignments"] -eq $true) {
            "Unique"
        } else {
            "Inherited"
        }

        # -------------------------------
        # Sharing link check (status only)
        # -------------------------------

        $sharingStatus = "No share info – inherited access only"
        $sharingDetails = "N/A"

        try {
            $endpoint = "/_api/web/GetFileByServerRelativeUrl('$path')/ListItemAllFields/ShareLink"
            $resp = Invoke-PnPSPRestMethod -Method Get -Url $endpoint -ErrorAction Stop

            if ($resp.value.Count -gt 0) {
                $sharingStatus = "Sharing link exists"
                $sharingDetails = ($resp.value | ForEach-Object {
                    "$($_.linkKind) / $($_.scope)"
                }) -join "; "
            }
        } catch { }

        $fileAccessStatus += [PSCustomObject]@{
            Library                 = $list.Title
            ItemPath               = Clean-Text $path
            ItemType               = $itemType
            InheritanceStatus      = $inheritance
            InheritedFrom          = if ($inheritance -eq "Inherited") { "Site / Library" } else { "Direct Permission" }
            EffectiveAccessSource  = $groupNames
            SharingStatus          = $sharingStatus
            SharingDetails         = $sharingDetails
        }
    }
}

# ------------------------------------------------
# Export result
# ------------------------------------------------

$fileAccessStatus | Export-Excel `
    -Path $excelPath `
    -WorksheetName "File Access Status" `
    -AutoSize `
    -TableStyle Medium6

Write-Host "`n FILE ACCESS STATUS AUDIT COMPLETE" -ForegroundColor Green
Write-Host "📄 Output:" $excelPath -ForegroundColor Cyan

Disconnect-PnPOnline
``




###NOTE:
#* SharingDetails is a descriptive column that is only populated when SharePoint finds an actual sharing link on a file or folder eg View / Anonymous, Edit / SpecificPeople , NA (No sharing links exist)


