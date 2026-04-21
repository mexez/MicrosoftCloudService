
#This script is a **tenant-wide SharePoint governance audit tool** that scans all SharePoint sites and produces an Excel report of site ownership and access.
#It connects to each site in the tenant and extracts:

#* Site collection admins, owners, members, and visitors
#* Site metadata (URL, template, storage, sharing settings, lock state)

# A key feature is **identity resolution**, where it converts SharePoint’s raw technical identifiers (e.g. tenant claims, Azure AD GUIDs, system accounts) into **human-readable names** such as user emails or Microsoft 365 group display names using Azure AD lookups.
# It also standardises system-generated accounts (e.g. SharePoint system principals) and handles unresolved or deleted objects gracefully.
# Overall, the script transforms complex SharePoint permission data into a **readable, business-friendly governance report** that can be used for security reviews, ownership validation, and compliance auditing.



# ================================================================
# SharePoint Sites Membership Audit (FULL RESOLVED VERSION)
# ================================================================

function Clean-ExcelString($inputString) {
    if ($null -eq $inputString) { return "None" }
    return $inputString.ToString() -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", ""
}

function Resolve-MemberName($member) {

    if ($null -eq $member) { return "None" }

    # System account cleanup
    if ($member.LoginName -like "*SHAREPOINT\\system*") {
        return "System Account"
    }

    # Claims-based identities (groups/users/service principals)
    if ($member.LoginName -match "c:.*\|.*\|") {

        $raw  = $member.LoginName.Split("|")[-1]
        $guid = $raw -replace "_o$", ""

        try {
            # Try Azure AD Group
            return (Get-PnPAzureADGroup -Identity $guid -ErrorAction Stop).DisplayName
        }
        catch {
            try {
                # Try Azure AD User
                return (Get-PnPAzureADUser -Identity $guid -ErrorAction Stop).DisplayName
            }
            catch {
                return "Unknown ($guid)"
            }
        }
    }

    # Standard M365 user
    elseif ($member.LoginName -like "i:0#.f|membership|*") {
        if ($member.Email) { return $member.Email }
        return $member.Title
    }

    return $member.Title
}

# =======================
# AUTH CONFIG
# =======================
$tenantId  = "0758b553-2d3f-40ec-bf92-f660c064ac83"
$clientId  = "5ffac2a8-c1e1-41f2-994d-efbc9e9c33fb"
$certPath  = "C:\Users\bcadmin\Documents\ParklaneClientAsset.pfx"
$ExcelPath = "C:\Users\bcadmin\Downloads\Reports\DAReportSheet\Parklane_SPSites_Membership_Apr212026.xlsx"
$adminUrl  = "https://9058445020-admin.sharepoint.com"

$certPassword = Read-Host "Enter PFX password" -AsSecureString

# =======================
# CONNECT TO ADMIN CENTER
# =======================
Write-Host "Connecting to PnP Admin Center..." -ForegroundColor Cyan

Connect-PnPOnline -Url $adminUrl `
    -ClientId $clientId `
    -Tenant $tenantId `
    -CertificatePath $certPath `
    -CertificatePassword $certPassword

Write-Host "Connected successfully." -ForegroundColor Green

# =======================
# GET ALL SITES
# =======================
Write-Host "Fetching all SharePoint sites..." -ForegroundColor Yellow

$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction SilentlyContinue

Write-Host "Found $($allSites.Count) sites." -ForegroundColor Green

$sharepointReport = @()

# =======================
# MAIN LOOP
# =======================
foreach ($site in $allSites) {

    Write-Host "Processing: $($site.Url)" -ForegroundColor Gray

    try {

        Connect-PnPOnline -Url $site.Url `
            -ClientId $clientId `
            -Tenant $tenantId `
            -CertificatePath $certPath `
            -CertificatePassword $certPassword `
            -ErrorAction Stop

        # =======================
        # SITE ADMINS
        # =======================
        $admins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue

        $adminNames = if ($admins) {
            ($admins | ForEach-Object { Resolve-MemberName $_ }) -join "; "
        } else { "None" }

        # =======================
        # SITE OWNERS
        # =======================
        $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction SilentlyContinue

        $ownerMembers = if ($ownerGroup) {
            (Get-PnPGroupMember -Group $ownerGroup -ErrorAction SilentlyContinue |
                ForEach-Object { Resolve-MemberName $_ }) -join "; "
        } else { "None" }

        # =======================
        # SITE MEMBERS
        # =======================
        $memberGroup = Get-PnPGroup -AssociatedMemberGroup -ErrorAction SilentlyContinue

        $siteMembers = if ($memberGroup) {
            (Get-PnPGroupMember -Group $memberGroup -ErrorAction SilentlyContinue |
                ForEach-Object { Resolve-MemberName $_ }) -join "; "
        } else { "None" }

        # =======================
        # SITE VISITORS
        # =======================
        $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -ErrorAction SilentlyContinue

        $siteVisitors = if ($visitorGroup) {
            (Get-PnPGroupMember -Group $visitorGroup -ErrorAction SilentlyContinue |
                ForEach-Object { Resolve-MemberName $_ }) -join "; "
        } else { "None" }

        # =======================
        # OUTPUT ROW
        # =======================
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

    }
    catch {
        Write-Host "SKIPPED $($site.Url): $($_.Exception.Message)" -ForegroundColor DarkYellow

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

# =======================
# EXPORT TO EXCEL
# =======================
if ($sharepointReport) {

    $sharepointReport | Export-Excel `
        -Path $ExcelPath `
        -AutoSize `
        -TableStyle "Medium2" `
        -WorksheetName "SharePoint Sites"

    Write-Host "EXPORT COMPLETE: $ExcelPath" -ForegroundColor Green
}
else {
    Write-Host "No data collected." -ForegroundColor Red
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue
