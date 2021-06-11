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

    $DATE = Get-Date -Format yyyy_MM_dd
}
PROCESS {
    $Master_files = Get-ChildItem -Path $Path -File -Filter "Master_*.csv" | Where-Object { $_.name -notlike "*original*" }

    $Mailbox_Types = @(
        'User',
        'Shared',
        'Room',
        'Equipment'
    )

    foreach ($T in $Mailbox_Types) {
        $File = $Master_files | Where-Object { $_.Name -like "*$T*" }
        $CSV = Import-Csv $File.FullName

        foreach ($A in @("CDFA", "CDPH", "DCA")){
            $OutputFile = "$($A)_Master_Mailbox_List_$DATE.xlsx"
            $SheetName = "$T Mailboxes"
            $Tenant_Domains = $DOMAINS[$A]

            $Result = $CSV | Where-Object {$_.OldUPN -like "*$($Tenant_Domains[0])" -or $_.OldUPN -like "*$($Tenant_Domains[1])"}
            $Result = $Result | Select-Object @{n="Source Mailbox";e={$_.OldUPN}}, @{n="Target Mailbox";e={$_.UPN}} | Sort-Object -Property "Source Mailbox"
            $Result | Export-Excel -Path $OutputFile -WorksheetName $SheetName -AutoSize -FreezeTopRow -AutoFilter
        }
    }

}
END {}
