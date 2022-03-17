[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $OutputFile
)
BEGIN {
    Write-Host "Getting SharePoint sites..." -ForegroundColor Cyan
    $Sites = Get-SPOSite -Limit All | Where-Object { $_.Title } | Sort-Object Title
}
PROCESS {
    $Total = $Sites | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Results = @()
    foreach ($Site in $Sites) {
        Write-Progress -Activity "Getting SharePoint site stats..." -Status "[ $i / $Total ] | Site: $($Site.Title)" -PercentComplete (($i / $Total) * 100)
        $i++
        $Title = $Site.Title
        $SizeMb = $Site.StorageUsageCurrent
        $Template = $Site.Template
        $GroupObjectID = $Site.GroupId.guid
        if ($GroupObjectID -eq '00000000-0000-0000-0000-000000000000') {
            $MemberCount = "n/a"
            $Owner = $Site.Owner
            $HasPlanner = $false
            $GroupObjectID = "n/a"
            $GroupCreatedDate = "n/a"
        }
        else {
            $MemberCount = Get-UnifiedGroupLinks -Identity $GroupObjectID -LinkType Member | Measure-Object | Select-Object -ExpandProperty Count
            $Owner = Get-UnifiedGroupLinks -Identity $GroupObjectID -LinkType Owner | Select-Object -ExpandProperty WindowsLiveID
            $Planners = Get-PnPPlannerPlan -Group $GroupObjectID -ResolveIdentities
            if ($Planners) {
                $HasPlanner = $true
            }
            else {
                $HasPlanner = $false
            }
            $GroupCreatedDate = Get-UnifiedGroup $GroupObjectID | Select-Object -ExpandProperty WhenCreated
        }
        $IsTeamsConnected = $Site.IsTeamsConnected

        $Results += [PSCustomObject]@{
            Title                   = $Title
            "Remove Don't Migrate"  = $null
            Org                     = $null
            "Last Content Modified" = $Site.LastContentModifiedDate
            "Created Date"          = $GroupCreatedDate
            SizeMB                  = $SizeMb
            Template                = $Template
            "Is Teams Connected"    = $IsTeamsConnected
            "Has Planner"           = $HasPlanner
            "Member Count"          = $MemberCount
            Owner                   = $Owner -join " | "
        }
    }
}
END {
    if ($Results) {
        $Results | Export-Excel -AutoSize -AutoFilter -FreezeTopRow -Path $OutputFile
    }
    Write-Host "File saved to: $OutputFile" -ForegroundColor Green
}
