[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $File
)
BEGIN{
    $ErrorActionPreference = "stop"
    $CSV = Import-Csv $File
}
PROCESS{
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Failed = @()
    foreach ($Site in $CSV){
        Write-Progress "Creating SharePoint and Teams sites..." -Status "Total: [$i / $Total] | Site: $($Site.TargetName)" -PercentComplete (($i / $Total) * 100)
        $i++
        if($Site.Teams -eq "TRUE"){
            try{
                $NewTeam = New-Team -DisplayName -eq $Site.TargetName
                $Site.TargetMailNickName = $NewTeam.MailNickName
            }
            catch{
                $err = $_
                Write-Host "Failed to create Teams Site $($Site.TargetName). Error: $err" -ForegroundColor Red
                $Failed += $Site
            }
        }
        else{
            try{
                $NewSPSite = New-PNPSite -Type TeamSite -Title $Site.TargetName -Alias $Site.TargetName
                $Site.TargetMailNickName = $Site.TargetName
            }
            catch{
                $err = $_
                Write-Host "Failed to create SharePoint Site $($Site.TargetName). Error: $err" -ForegroundColor Red
                $Failed += $Site
            }
        }
    }
}
END{
    $Date = Get-Date -Format yyyy-MM-dd_HH.mm
    $CSV | Export-Csv -NoTypeInformation SharePoint-And-Teams-Site-Created-$Date.csv
    Write-Host "Created sites saved to CSV: 'SharePoint-And-Teams-Site-Created-$Date.csv'"
    if($Failed){
        $Failed | Export-Csv -NoTypeInformation Failed-SharePoint-And-Teams-Sites-$Date.csv
        Write-Host "Failed sites saved to CSV: 'Failed-SharePoint-And-Teams-Sites-$Date.csv'"
    }
}
