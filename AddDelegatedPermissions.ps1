[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $File,

    [Parameter(Mandatory = $true)]
    [string]
    $ErrorFile
)
BEGIN {
    $ErrorActionPreference = "Stop"
    $CSV = Import-Csv $File

    function AddGrantSendOnBehalfTo($Mailbox, $GrantTo) {
        Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{Add = "$GrantTo" } -ErrorAction Stop | Out-Null
    }
    function AddSendAs($Mailbox, $GrantTo) {
        Add-RecipientPermission $Mailbox -Trustee $GrantTo -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
    }
    function AddFullAs($Mailbox, $GrantTo) {
        Add-MailboxPermission $Mailbox -User $GrantTo -AccessRights FullAccess -ErrorAction Stop | Out-Null
    }
}
PROCESS {
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $ErrorList = @()
    foreach ($Line in $CSV) {
        Write-Progress -Activity "Adding delegated permissions..." -Status "Total: [$i / $Total] | Mailbox: $($Line.NewEmail)" -PercentComplete (($i / $Total) * 100)
        $i++
        $NewEmail = $line.NewEmail
        $NewUserGrantedPermission = $Line.NewUserGrantedPermission
        $Permission = $Line.Permission

        if ($Permission -eq "GrantSendOnBehalfTo") {
            try {
                AddGrantSendOnBehalfTo -Mailbox $NewEmail -GrantTo $NewUserGrantedPermission
            }
            Catch {
                $err = $_
                Write-Host "Failed to add $Permission on $NewEmail to $NewUserGrantedPermission. Error: $err" -ForegroundColor Red
                $ErrorList += $Line
                continue
            }
        }
        elseif ($Permission -eq "SendAs") {
            try {
                AddSendAs -Mailbox $NewEmail -GrantTo $NewUserGrantedPermission
            }
            Catch {
                $err = $_
                Write-Host "Failed to add $Permission on $NewEmail to $NewUserGrantedPermission. Error: $err" -ForegroundColor Red
                $ErrorList += $Line
                continue
            }
        }
        elseif ($Permission -eq "FullAccess") {
            try {
                AddFullAs -Mailbox $NewEmail -GrantTo $NewUserGrantedPermission
            }
            Catch {
                $err = $_
                Write-Host "Failed to add $Permission on $NewEmail to $NewUserGrantedPermission. Error: $err" -ForegroundColor Red
                $ErrorList += $Line
                continue
            }
        }
    }
}
END {
    if($ErrorList){
        $ErrorList | Export-Csv -NoTypeInformation -LiteralPath $ErrorFile
        Write-Host "Error file saved to: '$ErrorFile'" -ForegroundColor Red
    }
}

