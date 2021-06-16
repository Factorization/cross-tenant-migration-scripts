[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $InputFiles,

    [Parameter(Mandatory = $true)]
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
    }
}
