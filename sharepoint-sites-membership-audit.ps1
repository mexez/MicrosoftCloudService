###This script  audits all SharePoint Online sites in your tenant and exports a membership report to Excel.

# ================================================================
# SharePoint Sites Membership Audit
# ================================================================

function Clean-ExcelString($inputString) {
    if ($null -eq $inputString) { return "None" }
    return $inputString.ToString() -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", ""
}

$tenantId     = "0758b553-2d3f-40ec-bf92-f660c064ac83"
$clientId     = "5ffac2a8-c1e1-41f2-994d-efbc9e9c33fb"
$certPath     = "C:\Users\bcadmin\Documents\ParklaneClientAsset.pfx"
$ExcelPath    = "C:\Users\bcadmin\Downloads\Reports\DAReportSheet\Parklane_SPSites_Membership_Mar182026.xlsx"
$adminUrl     = "https://9058445020-admin.sharepoint.com"

$certPassword = Read-Host "Enter PFX password" -AsSecureString

# Connect to PnP Admin Center
Write-Host "Connecting to PnP Admin Center..." -ForegroundColor Cyan
Connect-PnPOnline -Url $adminUrl `
    -ClientId $clientId `
    -Tenant $tenantId `
    -CertificatePath $certPath `
    -CertificatePassword $certPassword

# Verify connection
$ctx = Get-PnPContext
if ($null -eq $ctx) {
    Write-Host "CONNECTION FAILED - stopping script." -ForegroundColor Red
    return
}
Write-Host "Connected successfully." -ForegroundColor Green

# Fetch all sites (no OneDrive)
Write-Host "Fetching all SharePoint sites..." -ForegroundColor Yellow
$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction SilentlyContinue
Write-Host "Found $($allSites.Count) sites." -ForegroundColor Green

$sharepointReport = @()

foreach ($site in $allSites) {
    Write-Host "  Processing: $($site.Url)" -ForegroundColor Gray

    try {
        Connect-PnPOnline -Url $site.Url `
            -ClientId $clientId `
            -Tenant $tenantId `
            -CertificatePath $certPath `
            -CertificatePassword $certPassword `
            -ErrorAction Stop

        $admins     = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
        $adminNames = if ($admins) { ($admins | Select-Object -ExpandProperty LoginName) -join "; " } else { "None" }

        $ownerGroup   = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue
        $ownerMembers = if ($ownerGroup) {
            (Get-PnPGroupMember -Group $ownerGroup -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty LoginName) -join "; "
        } else { "None" }

        $memberGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction SilentlyContinue
        $siteMembers = if ($memberGroup) {
            (Get-PnPGroupMember -Group $memberGroup -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty LoginName) -join "; "
        } else { "None" }

        $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -ErrorAction SilentlyContinue
        $siteVisitors = if ($visitorGroup) {
            (Get-PnPGroupMember -Group $visitorGroup -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty LoginName) -join "; "
        } else { "None" }

        $sharepointReport += [PSCustomObject]@{
            "Site Title"         = Clean-ExcelString $site.Title
            "Site URL"           = $site.Url
            "Template"           = $site.Template
            "Storage Used (MB)"  = [math]::Round($site.StorageUsageCurrent, 2)
            "Site Admins"        = Clean-ExcelString $adminNames
            "Site Owners"        = Clean-ExcelString $ownerMembers
            "Site Members"       = Clean-ExcelString $siteMembers
            "Site Visitors"      = Clean-ExcelString $siteVisitors
            "Sharing Capability" = $site.SharingCapability
            "Locked State"       = $site.LockState
        }

    } catch {
        Write-Host "  SKIPPED $($site.Url): $($_.Exception.Message)" -ForegroundColor DarkYellow
        $sharepointReport += [PSCustomObject]@{
            "Site Title"         = Clean-ExcelString $site.Title
            "Site URL"           = $site.Url
            "Template"           = $site.Template
            "Storage Used (MB)"  = 0
            "Site Admins"        = "Access Error"
            "Site Owners"        = "Access Error"
            "Site Members"       = "Access Error"
            "Site Visitors"      = "Access Error"
            "Sharing Capability" = $site.SharingCapability
            "Locked State"       = $site.LockState
        }
    }
}

# Make sure Excel file is CLOSED before this runs
if ($sharepointReport) {
    $sharepointReport | Export-Excel -Path $ExcelPath -AutoSize -TableStyle "Medium2" -WorksheetName "SharePoint Sites"
    Write-Host "TEST COMPLETE: $ExcelPath" -ForegroundColor Green
} else {
    Write-Host "No data collected." -ForegroundColor Red
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
