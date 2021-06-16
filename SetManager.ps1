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
    $Exports_Name_Hash = $Exports | Group-Object -Property Name -AsHashTable

    function CheckManager([string]$manager){
        $manager_upn = $Exports_Name_Hash[$manager].UserPrincipalName
        if(-not $manager_upn){
            return $false
        }
        $manager_upn = ($manager_upn -split '@')[0] + "@cannabis.ca.gov"
        $Manager_ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$manager_upn'" -Server $Server
        if (-not $Manager_ADUser) {
            Write-Host "Manager $manager_upn not found." -ForegroundColor Red
            return $false
        }
        if ( ($Manager_ADUser | Measure-Object | Select-Object -ExpandProperty Count) -ne 1 ) {
            Write-Host "Manager $manager_upn found multiple." -ForegroundColor Red
            return $false
        }
        return $Manager_ADUser.DistinguishedName
    }
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

        $manager_dn = CheckManager -manager $Manager
        if (-not $manager_dn){
            continue
        }
        Write-Host $manager_dn -ForegroundColor Green
    }
}
