<#
.SYNOPSIS
    Analyzes team/SharePoint site memberships against reference Entra ID security groups.

.DESCRIPTION
    This script pulls all members from specified teams/SharePoint sites and compares them
    against reference Entra ID security groups. It categorizes sites based on whether their
    members belong to a single reference group or multiple reference groups. The output is
    generated as a CSV file with members listed in a single cell separated by a configurable
    delimiter (default: semicolon).

.PARAMETER TenantUrl
    The SharePoint tenant URL (e.g., https://contoso.sharepoint.com)

.PARAMETER ClientId
    The Entra ID App Client ID for PnP connection

.PARAMETER CertificateThumbprint
    The certificate thumbprint for app-only authentication (optional, use this OR ClientSecret)

.PARAMETER ClientSecret
    The client secret for app-only authentication (optional, use this OR CertificateThumbprint)

.PARAMETER ReferenceGroupsFile
    Path to the file containing reference Entra ID security groups (default: referencegroups.txt)

.PARAMETER TeamSitesFile
    Path to the file containing teams/sites to analyze (default: teamsites.txt)

.PARAMETER OutputFile
    Path to the output CSV report file (default: membership-report.csv)

.PARAMETER MemberSeparator
    The separator character to use between members in the CSV output (default: ;)
    Common options: ";" (semicolon), "|" (pipe), "," (comma - may cause CSV issues), or any other character

.EXAMPLE
    .\Analyze-TeamMembership.ps1 -TenantUrl "https://contoso.sharepoint.com" -ClientId "your-app-id" -ClientSecret "your-secret"

.EXAMPLE
    .\Analyze-TeamMembership.ps1 -TenantUrl "https://contoso.sharepoint.com" -ClientId "your-app-id" -CertificateThumbprint "ABC123..."

.EXAMPLE
    .\Analyze-TeamMembership.ps1 -TenantUrl "https://contoso.sharepoint.com" -ClientId "your-app-id" -ClientSecret "your-secret" -MemberSeparator "|"
#>

[CmdletBinding()]
param(
   

    [Parameter(ParameterSetName = 'Certificate')]
    [string]$CertificateThumbprint,

    

    [Parameter(Mandatory = $false)]
    [string]$ReferenceGroupsFile = "referencegroups.txt",

    [Parameter(Mandatory = $false)]
    [string]$TeamSitesFile = "teamssites.txt",

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "membership-report.csv",

    [Parameter(Mandatory = $false)]
    [string]$MemberSeparator = ";"
)

#Requires -Modules PnP.PowerShell, Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users

$TenantUrl       = "https://wincanton-admin.sharepoint.com"
$tenantDomain  = "wincanton.onmicrosoft.com"  # Replace with your actual tenant domain
$ClientId     = "0d662184-f5b0-475d-a3f2-5b38de95b716"  # GUID format# Connect to the SharePoint Admin site interactively

$CertificateThumbprint= (Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=SPOSetter" }).Thumbprint





# Color output helpers
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Read file and filter out comments and empty lines
function Read-ConfigFile {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        Write-ColorOutput "ERROR: File not found: $FilePath" -Color Red
        return @()
    }

    Get-Content $FilePath | Where-Object {
        $_ -notmatch '^\s*#' -and $_ -notmatch '^\s*$'
    }
}

# Get all members of an Entra ID group (recursive)
function Get-EntraGroupMembers {
    param(
        [string]$GroupId,
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey($GroupId)) {
        return $Cache[$GroupId]
    }

    try {
        $members = @()
        $groupMembers = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop

        foreach ($member in $groupMembers) {
            if ($member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                $userDetails = Get-MgUser -UserId $member.Id -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction SilentlyContinue
                if ($userDetails) {
                    $members += [PSCustomObject]@{
                        Id                = $userDetails.Id
                        DisplayName       = $userDetails.DisplayName
                        UserPrincipalName = $userDetails.UserPrincipalName
                        Email             = if ($userDetails.Mail) { $userDetails.Mail } else { $userDetails.UserPrincipalName }
                    }
                }
            }
            elseif ($member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                # Recursive call for nested groups
                $nestedMembers = Get-EntraGroupMembers -GroupId $member.Id -Cache $Cache
                $members += $nestedMembers
            }
        }

        $Cache[$GroupId] = $members
        return $members
    }
    catch {
        Write-ColorOutput "WARNING: Failed to get members for group $GroupId - $_" -Color Yellow
        return @()
    }
}

# Get SharePoint group members
function Get-SharePointGroupMembers {
    param(
        [string]$GroupId,
        [string]$SiteUrl
    )

    try {
        $members = @()

        # Try to get group members using Microsoft Graph (for M365 Groups/Teams)
        try {
            $graphMembers = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop

            foreach ($member in $graphMembers) {
                if ($member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                    $userDetails = Get-MgUser -UserId $member.Id -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction SilentlyContinue
                    if ($userDetails) {
                        $members += [PSCustomObject]@{
                            Id                = $userDetails.Id
                            DisplayName       = $userDetails.DisplayName
                            UserPrincipalName = $userDetails.UserPrincipalName
                            Email             = if ($userDetails.Mail) { $userDetails.Mail } else { $userDetails.UserPrincipalName }
                        }
                    }
                }
            }
        }
        catch {
            Write-ColorOutput "  Info: Group $GroupId is not an M365 Group, trying SharePoint..." -Color Cyan

            # If not an M365 Group, try SharePoint site
            if ($SiteUrl) {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $tenantDomain -Thumbprint $CertificateThumbprint

                $siteUsers = Get-PnPUser | Where-Object { $_.PrincipalType -eq 'User' }
                foreach ($user in $siteUsers) {
                    $members += [PSCustomObject]@{
                        Id                = $user.Id
                        DisplayName       = $user.Title
                        UserPrincipalName = $user.LoginName -replace '.*\|', ''
                        Email             = $user.Email
                    }
                }
            }
        }

        return $members
    }
    catch {
        Write-ColorOutput "  ERROR: Failed to get members for group $GroupId - $_" -Color Red
        return @()
    }
}

# Main script
try {
    Write-ColorOutput "`n=== Team/SharePoint Site Membership Analyzer ===" -Color Cyan
    Write-ColorOutput "Starting analysis...`n" -Color Cyan

    # Connect to Microsoft Graph
    Write-ColorOutput "Connecting to Microsoft Graph..." -Color Yellow

   
Connect-MgGraph
    Write-ColorOutput "Connected to Microsoft Graph successfully!`n" -Color Green

    # Read reference groups
    Write-ColorOutput "Reading reference groups from $ReferenceGroupsFile..." -Color Yellow
    $referenceGroupIds = Read-ConfigFile -FilePath $ReferenceGroupsFile

    if ($referenceGroupIds.Count -eq 0) {
        Write-ColorOutput "ERROR: No reference groups found in $ReferenceGroupsFile" -Color Red
        exit 1
    }

    Write-ColorOutput "Found $($referenceGroupIds.Count) reference group(s)`n" -Color Green

    # Read teams/sites
    Write-ColorOutput "Reading teams/sites from $TeamSitesFile..." -Color Yellow
    $teamSitesLines = Read-ConfigFile -FilePath $TeamSitesFile

    if ($teamSitesLines.Count -eq 0) {
        Write-ColorOutput "ERROR: No teams/sites found in $TeamSitesFile" -Color Red
        exit 1
    }

    Write-ColorOutput "Found $($teamSitesLines.Count) team(s)/site(s)`n" -Color Green

    # Parse teams/sites configuration
    $teamSites = @()
    foreach ($line in $teamSitesLines) {
        $parts = $line -split '\|'
        $teamSites += [PSCustomObject]@{
            Name    = $parts[0].Trim()
            GroupId = $parts[1].Trim()
            SiteUrl = if ($parts.Count -gt 2) { $parts[2].Trim() } else { $null }
        }
    }

    # Cache for group memberships
    $groupMembershipCache = @{}

    # Prepare reference group metadata (don't load members yet - lazy load on demand)
    Write-ColorOutput "Preparing reference group metadata..." -Color Yellow
    $referenceGroupData = @{}

    foreach ($refGroupId in $referenceGroupIds) {
        Write-ColorOutput "  Loading reference group: $refGroupId" -Color Gray

        try {
            # Try to get group by ID first, then by display name
            $group = $null
            try {
                $group = Get-MgGroup -GroupId $refGroupId -ErrorAction Stop
            }
            catch {
                # Try to find by display name
                $groups = Get-MgGroup -Filter "displayName eq '$refGroupId'" -ErrorAction Stop
                if ($groups) {
                    $group = $groups[0]
                }
            }

            if ($group) {
                # Store only metadata, not members (lazy load later)
                $referenceGroupData[$group.DisplayName] = @{
                    Id      = $group.Id
                    Members = $null  # Will be loaded on-demand
                }
                Write-ColorOutput "    Reference group ready: '$($group.DisplayName)'" -Color Green
            }
            else {
                Write-ColorOutput "    WARNING: Group not found: $refGroupId" -Color Yellow
            }
        }
        catch {
            Write-ColorOutput "    ERROR: Failed to process reference group $refGroupId - $_" -Color Red
        }
    }

    Write-ColorOutput "`nReference groups ready: $($referenceGroupData.Keys.Count)`n" -Color Green

    # Analyze each team/site and prepare CSV data
    $csvData = @()

    foreach ($teamSite in $teamSites) {
        Write-ColorOutput "`nAnalyzing: $($teamSite.Name) (GroupID: $($teamSite.GroupId))" -Color Cyan

        # Get site members
        $siteMembers = Get-SharePointGroupMembers -GroupId $teamSite.GroupId -SiteUrl $teamSite.SiteUrl

        if ($siteMembers.Count -eq 0) {
            Write-ColorOutput "  No members found or unable to access group" -Color Yellow
            $csvData += [PSCustomObject]@{
                'Site/Team Name'    = $teamSite.Name
                'Group ID'          = $teamSite.GroupId
                'Status'            = 'NO MEMBERS FOUND OR ACCESS DENIED'
                'Total Members'     = 0
                'Reference Groups'  = ''
                'Members'           = ''
            }
            continue
        }

        Write-ColorOutput "  Found $($siteMembers.Count) member(s)" -Color Green

        # Lazy load reference group members only now (on-demand)
        Write-ColorOutput "  Loading reference group members for comparison..." -Color Gray
        foreach ($refGroupName in $referenceGroupData.Keys) {
            if ($null -eq $referenceGroupData[$refGroupName].Members) {
                $refGroupId = $referenceGroupData[$refGroupName].Id
                Write-ColorOutput "    Loading members from: $refGroupName" -Color DarkGray
                $members = Get-EntraGroupMembers -GroupId $refGroupId -Cache $groupMembershipCache
                $referenceGroupData[$refGroupName].Members = $members
                Write-ColorOutput "      Loaded $($members.Count) member(s)" -Color DarkGray
            }
        }

        # Categorize members by reference group
        $membersByRefGroup = @{}
        $membersNotInAnyRefGroup = @()

        foreach ($member in $siteMembers) {
            $foundInGroups = @()

            foreach ($refGroupName in $referenceGroupData.Keys) {
                $refGroupMembers = $referenceGroupData[$refGroupName].Members
                if ($refGroupMembers.Id -contains $member.Id) {
                    $foundInGroups += $refGroupName
                }
            }

            if ($foundInGroups.Count -eq 0) {
                $membersNotInAnyRefGroup += $member
            }
            else {
                foreach ($groupName in $foundInGroups) {
                    if (-not $membersByRefGroup.ContainsKey($groupName)) {
                        $membersByRefGroup[$groupName] = @()
                    }
                    $membersByRefGroup[$groupName] += $member
                }
            }
        }

        if ($membersByRefGroup.Keys.Count -eq 0) {
            # No members belong to any reference group
            Write-ColorOutput "  Status: NO REFERENCE GROUP MATCH" -Color Yellow
            $membersList = ($siteMembers | ForEach-Object { "$($_.DisplayName) ($($_.Email))" }) -join $MemberSeparator
            $csvData += [PSCustomObject]@{
                'Site/Team Name'    = $teamSite.Name
                'Group ID'          = $teamSite.GroupId
                'Status'            = 'NO REFERENCE GROUP MATCH'
                'Total Members'     = $siteMembers.Count
                'Reference Groups'  = 'None'
                'Members'           = $membersList
            }
        }
        elseif ($membersByRefGroup.Keys.Count -eq 1 -and $membersNotInAnyRefGroup.Count -eq 0) {
            # All members belong to exactly one reference group
            $refGroupName = $membersByRefGroup.Keys[0]
            Write-ColorOutput "  Status: SINGLE REFERENCE GROUP - $refGroupName" -Color Green
            $uniqueMembers = $membersByRefGroup[$refGroupName] | Sort-Object Id -Unique
            $membersList = ($uniqueMembers | ForEach-Object { "$($_.DisplayName) ($($_.Email))" }) -join $MemberSeparator
            $csvData += [PSCustomObject]@{
                'Site/Team Name'    = $teamSite.Name
                'Group ID'          = $teamSite.GroupId
                'Status'            = 'SINGLE REFERENCE GROUP'
                'Total Members'     = $siteMembers.Count
                'Reference Groups'  = $refGroupName
                'Members'           = $membersList
            }
        }
        else {
            # Mixed membership - create a single row with reference groups as delimited string
            Write-ColorOutput "  Status: MIXED MEMBERSHIP" -Color Magenta

            # Build a string of all reference groups found
            $refGroupsList = ($membersByRefGroup.Keys | Sort-Object) -join $MemberSeparator

            # Add "Not in any reference group" if applicable
            if ($membersNotInAnyRefGroup.Count -gt 0) {
                $refGroupsList += $MemberSeparator + "Not in any reference group"
            }

            # Collect all unique members with their reference groups
            $allMembersWithGroups = @()

            # Add members from each reference group with group label
            foreach ($refGroupName in ($membersByRefGroup.Keys | Sort-Object)) {
                $uniqueMembers = $membersByRefGroup[$refGroupName] | Sort-Object Id -Unique
                foreach ($member in $uniqueMembers) {
                    $allMembersWithGroups += "[${refGroupName}] $($member.DisplayName) ($($member.Email))"
                }
            }

            # Add members not in any reference group
            if ($membersNotInAnyRefGroup.Count -gt 0) {
                foreach ($member in $membersNotInAnyRefGroup) {
                    $allMembersWithGroups += "[Not in any reference group] $($member.DisplayName) ($($member.Email))"
                }
            }

            $membersList = $allMembersWithGroups -join $MemberSeparator

            $csvData += [PSCustomObject]@{
                'Site/Team Name'    = $teamSite.Name
                'Group ID'          = $teamSite.GroupId
                'Status'            = 'MIXED MEMBERSHIP'
                'Total Members'     = $siteMembers.Count
                'Reference Groups'  = $refGroupsList
                'Members'           = $membersList
            }
        }
    }

    # Save CSV report
    $csvData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-ColorOutput "`n=== Analysis Complete ===" -Color Cyan
    Write-ColorOutput "Report saved to: $OutputFile`n" -Color Green

    # Display summary
    Write-ColorOutput "Summary:" -Color Cyan
    Write-ColorOutput "  Teams/Sites Analyzed: $($teamSites.Count)" -Color White
    Write-ColorOutput "  Reference Groups: $($referenceGroupData.Keys.Count)" -Color White
    Write-ColorOutput "`nReport file: $OutputFile" -Color Yellow
}
catch {
    Write-ColorOutput "`nERROR: $($_.Exception.Message)" -Color Red
    Write-ColorOutput "Stack Trace: $($_.ScriptStackTrace)" -Color Red
    exit 1
}
finally {
    # Disconnect
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnect errors
    }
}
