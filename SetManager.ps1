[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $InputFiles,

    [Parameter(Mandatory = $false)]
    [string]
    $Server = "dcc-h-dc01.ad.cannabis.ca.gov"
)
BEGIN {
    $Exports = $InputFiles | ForEach-Object { Import-Csv $_ }
}
PROCESS {
    foreach ($User in $Exports) {
        $UPN = ($User.UserPrincipalName -split '@')[0] + '@cannabis.ca.gov'

        $Manager = $User.Manager
        if ([string]::IsNullOrWhiteSpace($Manager)) {
            Write-Host "User $UPN has no manager defined." -ForegroundColor Yellow
            Continue
        }

        $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Server $Server
        if (-not $ADUser) {
            Write-Host "User $UPN not found." -ForegroundColor Red
            Continue
        }
        if ( ($ADUser | Measure-Object | Select-Object -ExpandProperty Count) -ne 1 ) {
            Write-Host "User $UPN found multiple." -ForegroundColor Red
            Continue
        }

        Write-Host $Manager
    }
}
