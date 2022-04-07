[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $MasterFile,

    [Parameter(Mandatory = $false)]
    [string[]]
    $ExcludeUsers = @("MigrationWiz@cabcsh.onmicrosoft.com"),

    [Parameter(Mandatory = $false)]
    [string]
    $OutputDir = $PSScriptRoot,

    [Parameter(Mandatory = $false)]
    [string]
    $ResultFilePrefix = $Null
)
BEGIN {
    $ErrorActionPreference = "Stop"
    $CSV = Import-Csv $MasterFile

    function GetGrantSendOnBehalfTo($Email) {
        try {
            $MB = Get-EXOMailbox $Email -Properties GrantSendOnBehalfTo -ErrorAction "Stop"
            if ($MB.GrantSendOnBehalfTo) {
                $Results = @()
                foreach ($User in $MB.GrantSendOnBehalfTo) {
                    try {
                        $Recipient = Get-EXORecipient $User -Properties PrimarySMTPAddress -ErrorAction "Stop"
                    }
                    catch {
                        Write-Host "Failed to find user '$User' for GrantSendOnBehalfTo. Skipping user." -ForegroundColor Cyan
                        continue
                    }
                    if ($Recipient.PrimarySMTPAddress -in $ExcludeUsers) {
                        continue
                    }
                    if ($Recipient.PrimarySMTPAddress -eq $Email) {
                        continue
                    }
                    $Results += [PSCustomObject]@{
                        Email                 = $Email
                        UserGrantedPermission = $Recipient.PrimarySMTPAddress
                        Permission            = "GrantSendOnBehalfTo"
                    }
                }
                return $Results
            }
        }
        catch {
            $err = $_
            $Result = [PSCustomObject]@{
                Email = $Email
                Task  = "Getting 'GrantSendOnBehalfTo' from Get-EXOMailbox"
                Error = $Err
            }
            $ErrorLog += $Result
        }
    }
    function GetFullAccess($Email) {
        try {
            $MBPermission = Get-EXOMailboxPermission $Email -ErrorAction "Stop" | Where-Object { $_.user -notlike "NT AUTHORITY\*" -and $_.user -notlike "s-1-*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "FullAccess" }
            $Results = @()
            foreach ($User in $MBPermission) {
                Try {
                    $Recipient = Get-EXORecipient $User.User -Properties PrimarySMTPAddress -ErrorAction "Stop"
                }
                catch {
                    Write-Host "Failed to find user '$($User.User)' for FullAccess permission. Skipping user." -ForegroundColor Cyan
                    continue
                }
                if ($Recipient.PrimarySMTPAddress -in $ExcludeUsers) {
                    continue
                }
                if ($Recipient.PrimarySMTPAddress -eq $Email) {
                    continue
                }
                $Results += [PSCustomObject]@{
                    Email                 = $Email
                    UserGrantedPermission = $Recipient.PrimarySMTPAddress
                    Permission            = "FullAccess"
                }
            }
            return $Results
        }
        catch {
            $err = $_
            $Result = [PSCustomObject]@{
                Email = $Email
                Task  = "Getting 'FullAccess' from Get-EXOMailboxPermission"
                Error = $Err
            }
            $ErrorLog += $Result
        }
    }
    function GetSendAs($Email) {
        try {
            $RecipientPermission = Get-EXORecipientPermission $Email -ErrorAction "Stop" | Where-Object { $_.trustee -notlike "NT AUTHORITY\*" -and $_.trustee -notlike "s-1-*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "SendAs" }
            $Results = @()
            foreach ($User in $RecipientPermission) {
                try {
                    $Recipient = Get-EXORecipient $User.trustee -Properties PrimarySMTPAddress -ErrorAction "Stop"
                }
                catch {
                    Write-Host "Failed to find user '$($User.trustee)' for SendAs permission. Skipping user." -ForegroundColor Cyan
                    continue
                }
                if ($Recipient.PrimarySMTPAddress -in $ExcludeUsers) {
                    continue
                }
                if ($Recipient.PrimarySMTPAddress -eq $Email) {
                    continue
                }
                $Results += [PSCustomObject]@{
                    Email                 = $Email
                    UserGrantedPermission = $Recipient.PrimarySMTPAddress
                    Permission            = "SendAs"
                }
            }
            return $Results
        }
        catch {
            $err = $_
            $Result = [PSCustomObject]@{
                Email = $Email
                Task  = "Getting 'SendAs' from Get-EXORecipientPermission"
                Error = $Err
            }
            $ErrorLog += $Result
        }
    }
}
PROCESS {
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Results = @()
    $ErrorLog = @()
    foreach ($User in $CSV) {
        Write-Progress -Activity "Exporting delegated permissions..." -Status "Total: [ $i/$Total] | User: $($User.OldUPN)" -PercentComplete (($i / $Total) * 100)
        $i++

        $Email = $User.OldUPN

        $GrantSendOnBehalfTo = GetGrantSendOnBehalfTo $Email
        if ($GrantSendOnBehalfTo) {
            $Results += $GrantSendOnBehalfTo
        }

        $FullAccess = GetFullAccess $Email
        if ($FullAccess) {
            $Results += $FullAccess
        }

        $SendAs = GetSendAs $Email
        if ($SendAs) {
            $Results += $SendAs
        }
    }
}
END {
    $Date = Get-Date -Format yyyy-MM-dd_HH.mm
    if ($Results) {
        $file = "Mailbox-Delegated-Permissions-Export-$Date.csv"
        if ($ResultFilePrefix) {
            $file = $ResultFilePrefix + "-" + $file
        }
        $ResultsFile = Join-Path $OutputDir $file
        $Results | Export-Csv -NoTypeInformation $ResultsFile
        Write-Host "Created sites saved to CSV: '$ResultsFile'" -ForegroundColor Green
    }
    if ($ErrorLog) {
        $file = "Mailbox-Delegated-Permissions-Export-Errors-$Date.csv"
        if ($ResultFilePrefix) {
            $file = $ResultFilePrefix + "-" + $file
        }
        $ErrorFile = Join-Path $OutputDir $file
        $ErrorLog | Export-Csv -NoTypeInformation $ErrorFile
        Write-Host "Failed sites saved to CSV: '$ErrorFile'" -ForegroundColor Red
    }
}

