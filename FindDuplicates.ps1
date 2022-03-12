[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $Domain,

    [Parameter(Mandatory = $true)]
    [string]
    $ExportCSV
)
BEGIN {
    $ErrorActionPreference = "Stop"
    function GetUPN($OldUPN, $Number = $null) {
        $Prefix = ($OldUPN -split "@")[0]
        if (-not $Prefix) {
            Throw "UPN prefix error for '$OldUPN'."
        }
        if ($Number) {
            $Prefix = $Prefix + "$Number"
        }
        return $Prefix + "@dca.ca.gov"
    }
    function GetLocation($OldUPN) {
        $Suffix = ($OldUPN -split "@")[1]
        if ($Suffix -match "^bcsh\.ca\.gov$|^cabcsh\.onmicrosoft\.com$") {
            $Location = "BCSH"
        }
        elseif ($Suffix -eq "ccap.ca.gov") {
            $Location = "CCAP"
        }
        else {
            Throw "Failed to parse OU from old UPN for '$OldUPN', mailbox type '$MailboxType'."
        }
        return $Location
    }
    Write-Host "Getting AD users..." -ForegroundColor Cyan
    $TargetADUsers = Get-Aduser -Filter * -Server $Domain -Properties CanonicalName, Mail, ProxyAddresses | Select-Object *

    Write-Host "Importing CSV..." -ForegroundColor Cyan
    $Export = Import-Csv $ExportCSV
}
PROCESS {
    $Duplicate_UPNs = @()
    $Duplicate_SamAccountName = @()
    $Duplicate_Mail = @()
    $Duplicate_ProxyAddresses = @()
    $DuplicateCount = 0

    $Total = $Export | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    foreach ($User in $Export) {
        Write-Progress -Activity "Finding duplicate users..." -Status "Users: [$i / $Total] | Duplicates: $DuplicateCount | User: $($User.UserPrincipalName)" -PercentComplete (($i / $Total) * 100)
        $i++

        $MailboxType = $User.RecipientTypeDetails
        $UPN = GetUPN -OldUPN $User.UserPrincipalName
        $SamAccountName = $User.SamAccountName
        if (-not $SamAccountName) {
            $SamAccountName = ($UPN -split '@')[0]
        }

        if ($MailboxType -ne "UserMailbox") {
            $Location = GetLocation -OldUPN $User.UserPrincipalName
            if (-not $UPN.StartsWith($Location)) {
                $UPN = $Location + '.' + $UPN
            }
            if (-not $SamAccountName.StartsWith($Location)) {
                $SamAccountName = $Location + "." + $SamAccountName
            }
        }

        if ($SamAccountName.Length -gt 20) {
            $SamAccountName = $SamAccountName.Substring(0, 20)
        }

        Write-Verbose "Finding duplicates for:"
        Write-Verbose "`tUPN: $UPN"
        Write-Verbose "`tSamAccountName: $SamAccountName"

        # UPN
        $Matching_DCA_User = $TargetADUsers | Where-Object { $_.UserPrincipalName -eq $UPN }
        If ($Matching_DCA_User) {
            $DuplicateCount++
            $Duplicate_UPNs += [PSCustomObject]@{
                Name              = $User.Name
                BCSH_UPN          = $User.UserPrincipalName
                Matching_DCA_User = ($Matching_DCA_User -split "/" | Select-Object -Skip 1) -join "/"
            }
        }

        # SamAccountName
        $Matching_DCA_User = $TargetADUsers | Where-Object { $_.SamAccountName -eq $SamAccountName }
        If ($Matching_DCA_User) {
            $DuplicateCount++
            $Duplicate_SamAccountName += [PSCustomObject]@{
                Name                = $User.Name
                BCSH_SamAccountName = $User.SamAccountName
                Matching_DCA_User   = ($Matching_DCA_User -split "/" | Select-Object -Skip 1) -join "/"
            }
        }

        # Mail
        $Matching_DCA_User = $TargetADUsers | Where-Object { $_.mail -eq $UPN }
        If ($Matching_DCA_User) {
            $DuplicateCount++
            $Duplicate_Mail += [PSCustomObject]@{
                Name              = $User.Name
                BCSH_Mail         = $User.UserPrincipalName
                Matching_DCA_User = ($Matching_DCA_User -split "/" | Select-Object -Skip 1) -join "/"
            }
        }

        # ProxyAddresses
        $Matching_DCA_User = $TargetADUsers | Where-Object { $_.ProxyAddresses -match $UPN }
        If ($Matching_DCA_User) {
            $DuplicateCount++
            $Duplicate_ProxyAddresses += [PSCustomObject]@{
                Name              = $User.Name
                BCSH_UPN          = $User.UserPrincipalName
                Matching_DCA_User = ($Matching_DCA_User -split "/" | Select-Object -Skip 1) -join "/"
            }
        }
    }
}
END {
    if ($Duplicate_UPNs) {
        Write-Host "Duplicate UPNs:" -ForegroundColor Cyan
        $Duplicate_UPNs | Format-Table -AutoSize -Wrap
    }
    if ($Duplicate_SamAccountName) {
        Write-Host "Duplicate SamAccountNames:" -ForegroundColor Cyan
        $Duplicate_SamAccountName | Format-Table -AutoSize -Wrap
    }
    if ($Duplicate_Mail) {
        Write-Host "Duplicate Mail:" -ForegroundColor Cyan
        $Duplicate_Mail | Format-Table -AutoSize -Wrap
    }
    if ($Duplicate_ProxyAddresses) {
        Write-Host "Duplicate ProxyAddresses:" -ForegroundColor Cyan
        $Duplicate_ProxyAddresses | Format-Table -AutoSize -Wrap
    }
}
