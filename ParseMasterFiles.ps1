[CmdletBinding()]
param (
    [Parameter(Mandatory = $True)]
    [string]
    $Path
)
BEGIN {
    $DOMAINS = @{
        CDFA = @('cdfa.ca.gov', 'cdfao365.onmicrosoft.com')
        CDPH = @('cdph.ca.gov', 'cdph.onmicrosoft.com')
        DCA  = @('dca.ca.gov', 'dcao365.onmicrosoft.com')
    }
}
PROCESS {
    $Master_files = Get-ChildItem -Path $Path -File -Filter "Master_*.csv" | Where-Object { $_.name -notlike "*original*" }

    $Mailbox_Types = @(
        'Equipment',
        'Room',
        'Shared',
        'User'
    )

    $Results = @{
        CDFA = @{
            Equipment = @()
            Room      = @()
            Shared    = @()
            User      = @()
        }
        CDPH = @{
            Equipment = @()
            Room      = @()
            Shared    = @()
            User      = @()
        }
        DCA  = @{
            Equipment = @()
            Room      = @()
            Shared    = @()
            User      = @()
        }
    }

    foreach ($T in $Mailbox_Types) {
        $File = $Master_files | Where-Object { $_.Name -like "*$T*" }
        $CSV = Import-Csv $File.FullName

        foreach ($A in @("CDFA", "CDPH", "DCA")){
            $Tenant_Domains = $DOMAINS[$A]

            $Result = $CSV | Where-Object {$_.OldUPN -like "*$($Tenant_Domains[0])" -or $_.OldUPN -like "*$($Tenant_Domains[1])"}
            $Result = $Result | Select-Object @{n="Source User";e={$_.OldUPN}}, @{n="Target User";e={$_.UPN}}
            $Results[$A][$T] = $Result
        }
    }
    return $Results
}
END {}
