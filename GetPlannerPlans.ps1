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
                GroupName              = $GroupDisplayName
                PlanTitle              = $Plan.Title
                PlanCreatedDate        = $Plan.CreatedDateTime
                PlanOwnerGroup         = $Plan.Owner
                PlanCreatedBy          = $Plan.CreatedBy.User.UserPrincipalName
                GroupObjectID          = $GroupObjectID
                GroupSharePointSiteUrl = $SharePointSiteUrl
                GroupEmailAddress      = $Group.PrimarySmtpAddress
                GroupIsTeamsConnected  = $SPOSite.IsTeamsConnected
                GroupOwners            = $GroupOwners -join " | "
                GroupMembers           = $GroupMembers -join " | "
            }
        }
    }
}
END {
    if($Results){
        $Results | Export-Excel -AutoSize -AutoFilter -FreezeTopRow -Path $OutputFile
    }
    Write-Host "File saved to: $OutputFile" -ForegroundColor Green
}
