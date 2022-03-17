[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [String]
    $OutputFile
)
Begin {
    $ErrorActionPreference = "Stop"
    Write-Host "Getting unified groups..." -ForegroundColor Cyan
    $UnifiedGroups = Get-UnifiedGroup -ResultSize unlimited
}
PROCESS {
    $Total = $UnifiedGroups | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Results = @()
    foreach ($Group in $UnifiedGroups) {
        Write-Progress -Activity "Getting planner plans..." -Status "Group: [$i / $Total] | Planners Found: $($Results | Measure-Object | Select-Object -ExpandProperty Count) | Group: $($Group.DisplayName)" -PercentComplete (($i / $Total) * 100)
        $i++

        $GroupDisplayName = $Group.DisplayName
        $GroupName = $Group.Name
        $GroupObjectID = $Group.ExternalDirectoryObjectId

        $Planners = Get-PnPPlannerPlan -Group $GroupObjectID -ResolveIdentities
        if (-not $Planners) {
            continue
        }
        $SharePointSiteUrl = $Group.SharePointSiteUrl
        $SPOSite = Get-SPOSite $SharePointSiteUrl

        $GroupOwners = Get-UnifiedGroupLinks $GroupName -LinkType Owner | Select-Object -ExpandProperty WindowsLiveID
        $GroupMembers = Get-UnifiedGroupLinks $GroupName -LinkType Member | Select-Object -ExpandProperty WindowsLiveID
        foreach ($Plan in $Planners) {
            $Results += [PSCustomObject]@{
                "Group Name"                = $GroupDisplayName
                "Plan Title"                = $Plan.Title
                "Plan Created Date"         = $Plan.CreatedDateTime
                # "Plan Owner Group"         = $Plan.Owner
                "Plan Created By"           = $Plan.CreatedBy.User.UserPrincipalName
                "Group Object ID"           = $GroupObjectID
                "Group SharePoint Site URL" = $SharePointSiteUrl
                "Group Email Address"       = $Group.PrimarySmtpAddress
                "Group Is Teams Connected"  = $SPOSite.IsTeamsConnected
                "Group Owners"              = $GroupOwners -join " | "
                "Group Members"             = $GroupMembers -join " | "
            }
        }
    }
}
END {
    if ($Results) {
        $Results | Export-Excel -AutoSize -AutoFilter -FreezeTopRow -Path $OutputFile
    }
    Write-Host "File saved to: $OutputFile" -ForegroundColor Green
}
