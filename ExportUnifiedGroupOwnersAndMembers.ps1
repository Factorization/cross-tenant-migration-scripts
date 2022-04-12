[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]
    [string]
    $File,

    [Parameter(Mandatory = $False)]
    [string[]]
    $ExcludedUsers = @(),

    [Parameter(Mandatory = $false)]
    [string]
    $OutputDir = $PSScriptRoot
)
BEGIN{
    $ErrorActionPreference = "Stop"
    $CSV = Import-Csv $File
}
PROCESS{
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Results = @()

    foreach ($Group in $CSV){
        Write-Progress -Activity "Exporting Unified Group Permissions..." -Status "Groups: [$i / $Total] | Group: $($Group.SourceMailNickName)" -PercentComplete (($i / $Total) * 100)
        $i++

        $Members = Get-UnifiedGroupLinks $Group.SourceMailNickname -LinkType Member | Where-Object Name -NotIn $ExcludedUsers
        $Owners = Get-UnifiedGroupLinks $Group.SourceMailNickname -LinkType Owner | Where-Object Name -NotIn $ExcludedUsers

        foreach ($Member in $Members){
            $Results += [PSCustomObject]@{
                SourceName = $Group.SourceName
                SourceMailNickName = $Group.SourceMailNickName
                TargetName = $Group.TargetName
                TargetMailNickName = $Group.TargetMailNickName
                User = $Member.WindowsLiveId
                Permission = "Member"
            }
        }

        foreach ($Owner in $owners){
            $Results += [PSCustomObject]@{
                SourceName = $Group.SourceName
                SourceMailNickName = $Group.SourceMailNickName
                TargetName = $Group.TargetName
                TargetMailNickName = $Group.TargetMailNickName
                User = $Owner.WindowsLiveId
                Permission = "Owner"
            }
        }
    }
}
END {
    $Date = Get-Date -Format yyyy-MM-dd_HH.mm
    if ($Results) {
        $OutputFile = "Unified-Group-Owners-And-Members-$Date.csv"
        $ResultsFile = Join-Path $OutputDir $OutputFile
        $Results | Export-Csv -NoTypeInformation $ResultsFile
        Write-Host "Results saved to CSV: '$ResultsFile'" -ForegroundColor Green
    }
}
