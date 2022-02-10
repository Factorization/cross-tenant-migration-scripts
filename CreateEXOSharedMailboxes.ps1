[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $InputFile,

    [Parameter(Mandatory = $true)]
    [string]
    $Prefix,

    [Parameter(Mandatory = $true)]
    [string]
    $TargetEmailDomain,

    [Parameter(Mandatory = $false)]
    [string]
    $ResultsFile = "$($Prefix)_MailboxCreation_$(Get-Date -Format yyyy-MM-dd_HH.mm).csv"
)
BEGIN {
    function SaveResult($Result) {
        $Result.Results = $Result.Results -join "`r`n"
        $Result | Export-Csv -NoTypeInformation -Append -LiteralPath $ResultsFile
    }
    $Global:ErrorActionPreference = "Stop"
    try {
        Get-Command "Get-Mailbox" | Out-Null
    }
    catch {
        Write-Error "Not connected to Exchange Online. You must connect to Exchange Online first."
        exit
    }

    if (-not $Prefix.EndsWith("_")) {
        $Prefix = $Prefix + "_"
    }

    $CSV = Import-Csv $InputFile

}
PROCESS {
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty count
    $i = 0
    $Error_Count = 0
    foreach ($Line in $CSV) {
        Write-Progress -Activity "Creating EXO mailbox" -Status "Mailbox: $($Line.PrimarySmtpAddress) | Status: ($i / $total) | Error Count: $Error_Count" -PercentComplete "$(($i / $Total)*100)"
        $i++

        # Name
        $Name = $Prefix + $Line.Name
        # Alias
        $Alias = $Prefix + $Line.Alias
        # DisplayName
        $DisplayName = $Line.DisplayName
        # FirstName
        $FirstName = $Line.FirstName
        # LastName
        $LastName = $Line.LastName
        # Office
        $Office = $Line.Office
        # HiddenFromAddressList
        $HiddenFromAddressListsEnabled = $Line.HiddenFromAddressListsEnabled -eq $true
        # Archive Mailbox
        $ArchiveEnabled = $Line.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000"
        # PrimarySmtpAddress
        $PrimarySmtpAddress = $Prefix + ($Line.PrimarySmtpAddress -split "@")[0] + "@" + $TargetEmailDomain
        # ProxyAddresses
        $ProxyAddresses = @()
        $ProxyAddresses += "x500:$($Line.LegacyExchangeDN)"
        $ProxyAddresses += ($Line.EmailAddresses -split ";") | Where-Object { $_ -like "x500:*" } | ForEach-Object { $_ -creplace "^X500:", "x500:" }
        $ProxyAddresses = $ProxyAddresses | Sort-Object -Unique
        # FutureAliases
        $FutureAliases = @()
        $FutureAliases += ($Line.EmailAddresses -split ";") | Where-Object { $_ -like "smtp:*" } | Where-Object { $_ -notlike "*.onmicrosoft.com" }
        #Type
        $MailboxType = $Line.RecipientTypeDetails

        # Results Log
        $ResultLog = [PSCustomObject]@{
            Source_Name                 = $Line.Name
            Target_Name                 = $Name
            Source_Alias                = $line.Alias
            Target_Alias                = $Alias
            Source_PrimarySmtpAddress   = $Line.PrimarySmtpAddress
            Target_PrimarySmtpAddress   = $PrimarySmtpAddress
            Source_RecipientTypeDetails = $Line.RecipientTypeDetails
            Target_RecipientTypeDetails = $null
            FutureAliases               = $FutureAliases -join ";"
            Results                     = @()
        }

        # Check if mailbox already exists
        try {
            $NewMailbox = Get-mailbox $PrimarySmtpAddress -ErrorAction stop
            $ResultLog.Target_RecipientTypeDetails = $NewMailbox.RecipientTypeDetails
            $ResultLog.Results += "Mailbox already exists."
        }
        catch {
            $NewMailbox = $null
        }

        # Define properties
        $Properties = @{
            Name               = $Name
            Alias              = $Alias
            PrimarySmtpAddress = $PrimarySmtpAddress
            DisplayName        = $DisplayName
        }
        if ($FirstName) {
            $Properties.FirstName = $FirstName
        }
        if ($LastName) {
            $Properties.LastName = $LastName
        }

        # Create Mailbox
        if (-not $NewMailbox) {
            try {
                if ($MailboxType -eq "SharedMailbox" -or $MailboxType -eq "UserMailbox") {
                    $ResultLog.Target_RecipientTypeDetails = "SharedMailbox"
                    New-Mailbox -Shared @Properties | Out-Null
                }
                elseif ($MailboxType -eq "EquipmentMailbox") {
                    $ResultLog.Target_RecipientTypeDetails = "EquipmentMailbox"
                    New-Mailbox -Equipment @Properties | Out-Null
                }
                elseif ($MailboxType -eq "RoomMailbox") {
                    $ResultLog.Target_RecipientTypeDetails = "RoomMailbox"
                    New-Mailbox -Room @Properties | Out-Null
                }
                else {
                    $ResultLog.Results += "Unknown mailbox type."
                    SaveResult -Result $ResultLog
                    $ResultLog | Out-Host
                    $Error_Count++
                    continue
                }
                $ResultLog.Results += "Mailbox created."
            }
            Catch {
                $err = $_
                $ResultLog.Results += "Failed to create mailbox. Error: $err"
                SaveResult -Result $ResultLog
                $ResultLog | Out-Host
                $Error_Count++
                continue
            }
        }

        # Get Mailbox after creation
        $try = 0
        $MaxTries = 10
        $NewMailbox = $null
        while ($True) {
            if ($try -ge $MaxTries) { break }
            try {
                $NewMailbox = Get-mailbox $PrimarySmtpAddress -ErrorAction stop
                break
            }
            Catch {
                Start-Sleep -Seconds 10
            }
        }
        if (-not $NewMailbox) {
            $ResultLog.Results += "Unable to get new mailbox for $PrimarySmtpAddress. Skipping mailbox."
            SaveResult -Result $ResultLog
            $ResultLog | Out-Host
            $Error_Count++
            continue
        }

        # Set Mailbox Proxy Addresses
        try {
            Set-Mailbox $PrimarySmtpAddress -EmailAddresses @{add = $ProxyAddresses } | Out-Null
            $ResultLog.Results += "Updated proxy addresses."
        }
        Catch {
            $err = $_
            $ResultLog.Results += "Unable to update mailbox proxy addresses. Error: $err"
            SaveResult -Result $ResultLog
            $ResultLog | Out-Host
            $Error_Count++
            Continue
        }

        # Set Office
        if ($Office) {
            try {
                Set-Mailbox $PrimarySmtpAddress -Office $Office | Out-Null
                $ResultLog.Results += "Set office."
            }
            Catch {
                $err = $_
                $ResultLog.Results += "Unable to update mailbox office. Error: $err"
                SaveResult -Result $ResultLog
                $ResultLog | Out-Host
                $Error_Count++
                Continue
            }
        }

        # Set HiddenFromAddressListsEnabled
        if ($HiddenFromAddressListsEnabled) {
            try {
                Set-Mailbox $PrimarySmtpAddress -HiddenFromAddressListsEnabled:$True | Out-Null
                $ResultLog.Results += "Set HiddenFromAddressListsEnabled."
            }
            Catch {
                $err = $_
                $ResultLog.Results += "Unable to update mailbox HiddenFromAddressListsEnabled. Error: $err"
                SaveResult -Result $ResultLog
                $ResultLog | Out-Host
                $Error_Count++
                Continue
            }
        }

        # Set ArchiveEnabled
        if ($ArchiveEnabled) {
            if ($NewMailbox.ArchiveGuid -ne "00000000-0000-0000-0000-000000000000") {
                $ResultLog.Results += "Archive mailbox already enabled."
            }
            else {
                try {
                    Enable-Mailbox $PrimarySmtpAddress -Archive | Out-Null
                    $ResultLog.Results += "Enabled archive mailbox."
                }
                Catch {
                    $err = $_
                    $ResultLog.Results += "Unable to enable archive mailbox. Error: $err"
                    SaveResult -Result $ResultLog
                    $ResultLog | Out-Host
                    $Error_Count++
                    Continue
                }
            }
        }

        # Log Results
        SaveResult -Result $ResultLog
    }
}
END {
    Write-Host "Result file saved to: $ResultsFile" -ForegroundColor Green
}
