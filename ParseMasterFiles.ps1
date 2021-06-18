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

    $Total = $Master_files | ForEach-Object {Import-Csv $_.FullName} | Measure-Object | Select-Object -ExpandProperty Count
    $Count = 0
    $Output_Results = $()
    foreach ($T in $Mailbox_Types) {
        $File = $Master_files | Where-Object { $_.Name -like "*$T*" }
        $CSV = Import-Csv $File.FullName
        $Count = $Count + ($CSV | Measure-Object | Select-Object -ExpandProperty Count)

        foreach ($A in @("CDFA", "CDPH", "DCA")){
            $OutputFile = "$($A)_Master_Mailbox_List_$DATE.xlsx"
            $SheetName = "$T Mailboxes"
            $Tenant_Domains = $DOMAINS[$A]

            $Result = $CSV | Where-Object {$_.OldUPN -like "*$($Tenant_Domains[0])" -or $_.OldUPN -like "*$($Tenant_Domains[1])"}
            $Result = $Result | Select-Object @{n="Source Mailbox";e={$_.OldUPN}}, @{n="Target Mailbox";e={$_.UPN}} | Sort-Object -Property "Source Mailbox"
            $Result | Export-Excel -Path $OutputFile -WorksheetName $SheetName -AutoSize -FreezeTopRow -AutoFilter
            $Output_Results += "$A | $T - $($Result | Measure-Object| Select-Object -ExpandProperty Count)"
        }
    }
    $CDFA_Total = 0
    $CDPH_Total = 0
    $DCA_Total = 0
    $Output_Results = $Output_Results | Sort-Object
    foreach ($R in $Output_Results){
        $A = ($R -split '|')[0]
        $t = ($R -split ' - ')[1] -as [int]
        if($A -eq 'CDFA'){$CDFA_Total += $t}
        elseif ($A -eq 'CDPH') {$CDPH_Total += $t}
        elseif($A -eq 'DCA'){$DCA_Total += $t}

        Write-Host $R
    }
    Write-Host "CDFA Total: $CDFA_Total"
    Write-Host "CDPH Total: $CDPH_Total"
    Write-Host "DCA Total: $DCA_Total"

    if($Total -ne $Count){
        Write-Host "Numbers don't match. Total $Total. Processed $Count" -ForegroundColor Red
    }
    else{
        Write-Host "All mailboxes processed. Total $Total" -ForegroundColor Green
    }
}
END {}
