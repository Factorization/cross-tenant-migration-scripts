[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $File,

    [Parameter(Mandatory = $false)]
    [string]
    $OutputDir = $PSScriptRoot
)
BEGIN {
    $ErrorActionPreference = "stop"
    $CSV = Import-Csv $File
}
PROCESS {
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Failed = @()
    $Success = @()
    foreach ($Site in $CSV) {
        Write-Progress "Creating SharePoint and Teams sites..." -Status "Total: [$i / $Total] | Site: $($Site.TargetName)" -PercentComplete (($i / $Total) * 100)
        $i++
        if ($Site.Teams -eq "TRUE") {
            try {
                $NewTeam = New-Team -DisplayName $Site.TargetName -MailNickName $Site.TargetName -Visibility "Private"
                $Site.TargetMailNickName = $NewTeam.MailNickName
                $Success += $Site
            }
            catch {
                $err = $_
                Write-Host "Failed to create Teams Site $($Site.TargetName). Error: $err" -ForegroundColor Red
                $Failed += $Site
            }
        }
        else {
            try {
                $NewSPSite = New-PNPSite -Type TeamSite -Title $Site.TargetName -Alias $Site.TargetName
                $Site.TargetMailNickName = $Site.TargetName + "|" + $NewSPSite
                $Success += $Site
            }
            catch {
                $err = $_
                Write-Host "Failed to create SharePoint Site $($Site.TargetName). Error: $err" -ForegroundColor Red
                $Failed += $Site
            }
        }
    }
}
END {
    $Date = Get-Date -Format yyyy-MM-dd_HH.mm
    if ($Success) {
        $SuccessFile = Join-Path $OutputDir "SharePoint-And-Teams-Site-Created-$Date.csv"
        $Success | Export-Csv -NoTypeInformation $SuccessFile
        Write-Host "Created sites saved to CSV: '$SuccessFile'" -ForegroundColor Green
    }
    if ($Failed) {
        $FailedFile = Join-Path $OutputDir "Failed-SharePoint-And-Teams-Sites-$Date.csv"
        $Failed | Export-Csv -NoTypeInformation $FailedFile
        Write-Host "Failed sites saved to CSV: '$FailedFile'" -ForegroundColor Red
    }
}
