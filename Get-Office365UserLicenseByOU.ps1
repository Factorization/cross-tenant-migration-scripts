#Requires -version 3
#Requires -Modules MSOnline
#Requires -Modules ActiveDirectory
#Requires -Modules ExchangeOnlineManagement

<#
 .Synopsis
	Get Office 365 Users license by OU.

 .Description
	This script connects to Office 365 and local AD, gets all users, and generates an O365 license report. Must be ran as a user that has full read permissions to all AD objects. The credentials entered must have read access to Azure AD and Exchange Online.

 .Parameter Credential
	A PSCredential value parameter. Credential for Office 365.

 .Parameter Path
	A string value parameter. Path to the location where to save the output file.

 .Parameter FileName
	A string value parameter. Name for output file.

 .Parameter SendTo
    A string parameter. To address for email of output file.

 .Parameter SendFrom
    A string parameter. From address for email of output file.

 .Parameter SmtpServer
    A string parameter. SMTP server for email of output file.

 .Parameter DeleteOutputFile
    A switch value parameter. Deletes output file at end of script.

 .Example
	.\Get-Office365UserLicenseByOU.PS1

    Generates CSV and saves output to default location with the default file name.

 .Example
	.\Get-Office365UserLicenseByOU.PS1 -Credential $UserCredential

    Logs into MS Online using credentials saved in $UserCredential, generates CSV and saves output to default location with the default file name.

 .Example
	.\Get-Office365UserLicenseByOU.PS1 -FileName Output.csv

    Generates CSV and saves output to same location as the script with the file name Output.csv.

 .Example
	.\Get-Office365UserLicenseByOU.PS1 -Path C:\temp

    Generates CSV and saves output to c:\temp with the default file name (Note: C:\temp must already exist).

 .Example
	.\Get-Office365UserLicenseByOU.PS1 -SendTo 'user@domain.com' -SendFrom 'report@domain.com' -SmtpServer 'mail.domain.com' -DeleteOutputFile

    Generates CSV and saves output to default location with the default file name. Then emails the output file and deletes the output file.

 .Notes

	#######################################################
	#  .                                               .  #
	#  .                Written By:                    .  #
	#.....................................................#
	#  .              Jeffrey Kraemer                  .  #
	#.....................................................#
	#  .                                               .  #
	#######################################################
#>

[CmdletBinding(DefaultParameterSetName = 'NoEmail')]
Param(
    # Credential parameter.
    [Parameter(ParameterSetName = 'NoEmail', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Email', Mandatory = $false)]
    [PSCredential]
    $Credential = (Get-Credential -Message "Office 365 Login"),

    # Export path parameter.
    [Parameter(ParameterSetName = 'NoEmail', Mandatory = $FALSE)]
    [Parameter(ParameterSetName = 'Email', Mandatory = $false)]
    [ValidateScript( { Test-Path $_ -PathType 'Container' })]
    [String]
    $Path = $PSScriptRoot,

    # File name parameter.
    [Parameter(ParameterSetName = 'NoEMail', Mandatory = $FALSE)]
    [Parameter(ParameterSetName = 'Email', Mandatory = $false)]
    [String]
    $FileName = "Office365UserLicenseByOU_$(Get-Date -Format 'MM.dd.yy_H.mm').csv",

    #Delete output file
    [Parameter(ParameterSetName = 'NoEMail', Mandatory = $FALSE)]
    [Parameter(ParameterSetName = 'Email', Mandatory = $false)]
    [switch]
    $DeleteOutputFile,

    #SendTO
    [Parameter(ParameterSetName = 'Email', Mandatory = $True)]
    [String[]]
    $SendTo,

    #SendFrom
    [Parameter(ParameterSetName = 'Email', Mandatory = $True)]
    [String]
    $SendFrom,

    #SmtpServer
    [Parameter(ParameterSetName = 'Email', Mandatory = $True)]
    [string]
    $SmtpServer
)

BEGIN {
    #Check if emailing
    if ($SendTo -and $SendFrom -and $SmtpServer) {
        $Email = $true
        Write-Verbose -Message "Email flag set."
    }
    else {
        $Email = $false
        Write-Verbose -Message "Email flag not set."
    }


    # Append .csv to file name if missing.
    Write-Verbose -Message "Testing output file."
    if (-not $FileName.EndsWith('.csv')) {
        $FileName = "$FileName.csv"
    }

    # Build save file location.
    $SaveFilePath = "$Path\$FileName"

    # Test if output file already exists
    if (Test-Path -Path $SaveFilePath -PathType Leaf) {
        $Choice = ""
        while ($Choice -notmatch '^y$|^n$') {
            $Choice = Read-Host "File '$FileName' already exists. Do you want to overwrite? (Y/N)"
        }
        if ($Choice -eq 'n') {
            Write-Warning "Ending script. User chose not to overwrite output file."
            Break
        }
    }

    # Connect to MSOnline.
    Write-Verbose -Message "Connecting to MSOL Service."
    Try {
        Connect-MsolService -Credential $Credential -ErrorAction Stop
    }
    Catch {
        Write-Error "Can't connect to MS Online service with user '$($Credential.UserName)'. Please check password and/or permissions and try again."
        Break
    }

    # Connect to Exchange Online.
    Write-Verbose -Message "Connecting to Exchange Online."
    #Try {
        # $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -ErrorAction Stop
        # Import-PSSession -Session $Session -DisableNameChecking -AllowClobber -Name Get-Mailbox, Get-MailboxStatistics -Prefix O365 -FormatTypeName * | Out-Null
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$False  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
    #}
    #Catch {
    #    Write-Error "Can't connect to Exchange Online service with user '$($Credential.UserName)'. Please check password and/or permissions and try again."
    #    Break
    #}
}

PROCESS {
    # Get MSOL Users.
    Write-Verbose -Message "Getting MSOL Users."
    Try {
        $MSOLUsers = Get-MsolUser -All -ErrorAction Stop
    }
    Catch {
        Write-Error "Can't retrieve MS Online users. Please check your permissions and try again."
        Break
    }

    # Get shared mailboxes
    Write-Verbose -Message "Getting shared mailboxes."
    $SharedMailboxes = (Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox).UserPrincipalName
    Write-Verbose -Message "Getting mailbox statistics."
    $UserMailboxStats = Get-EXOMailbox -ResultSize Unlimited | Select-Object Name, UserPrincipalName, @{n = 'LastLogonTime'; e = { (Get-EXOMailboxStatistics -Identity $_.Identity -Properties LastLogonTIme).LastLogonTime } }

    # Get AD Users
    Write-Verbose -Message "Getting AD users."
    Try {
        $ADUsers = Get-ADUser -Filter * -Properties CanonicalName, Enabled, LastLogonDate, extensionAttribute5 -ErrorAction Stop
    }
    Catch {
        Write-Error "Can't retrieve AD users. Please check your permissions and try again."
        Break
    }

    # Build AD User dictionary
    $ADUsers_Hash = @{}
    foreach ($ADUser in $ADUsers) {
        $ADUsers_Hash["$($ADUser.UserPrincipalName)"] = @($ADUser | Select-Object CanonicalName, Enabled, LastLogonDate, extensionAttribute5)
    }

    # Build User Mailbox dictionary
    $UserMailbox_Hash = @{}
    foreach ($UserMailbox in $UserMailboxStats) {
        $UserMailbox_Hash["$($UserMailbox.UserPrincipalName)"] = @($UserMailbox | Select-Object LastLogonTime)
    }

    # Loop through users and build object with license info and OU.
    Write-Verbose -Message "Building licensing info for $($MSOLUsers.count) users."
    $Results = @()
    foreach ($User in $MSOLUsers) {
        # Build variables for user
        $UPN = $User.UserPrincipalName
        $DisplayName = $User.DisplayName
        $SharedMailbox = $UPN -in $SharedMailboxes

        # Test if user is licensed
        If ($User.isLicensed) {
            $License = $User.Licenses.AccountSku.SkuPartNumber -join ', '
            $ServicePlan = (($User.Licenses.ServiceStatus | Where-Object -FilterScript { $_.ProvisioningStatus -ne "Disabled" }).ServicePlan.ServiceName | Sort-Object) -join ', '
        }
        Else {
            $License = 'None'
            $ServicePlan = 'None'
        }

        # Test if user is in AD, if so get OU and Enabled
        $ADUser = $ADUsers_Hash[$UPN]
        If ($ADUser) {
            $CanonicalName = $ADUser.CanonicalName -split '/'
            $CanonicalNameLength = $CanonicalName.length
            $OULength = $CanonicalNameLength - 2
            $OU = $CanonicalName[1..$OULength] -join '/'
            $ADEnabled = $ADUser.Enabled
            if ($ADUser.LastLogonDate) {
                $ADLastLogonDate = $ADUser.LastLogonDate
            }
            else {
                $ADLastLogonDate = "No logon time in AD."
            }
            $LeaveReturnDate = $ADUser.extensionAttribute5
        }
        Else {
            Write-Warning "User '$UPN' not found in Active Directory"
            $OU = "User not in AD."
            $ADEnabled = "User not in AD."
            $ADLastLogonDate = "User not in AD."
            $LeaveReturnDate = $null
        }

        $UserMailbox = $UserMailbox_Hash[$UPN]
        If ($UserMailbox) {
            if ($UserMailbox.LastLogonTime) {
                $MailboxLastLogon = $UserMailbox.LastLogonTime
            }
            else {
                $MailboxLastLogon = "No logon time in EXO."
            }
        }
        else {
            $MailboxLastLogon = "User not in EXO."
        }

        # Build hash table for object
        $Properties = [ordered]@{
            'UserPrincipalName' = $UPN
            'DisplayName'       = $DisplayName
            'OU'                = $OU
            'AD_Enabled'        = $ADEnabled
            'AD_LastLogon'      = $ADLastLogonDate
            'MailboxLastLogon'  = $MailboxLastLogon
            'LeaveReturnDate'   = $LeaveReturnDate
            'SharedMailbox'     = $SharedMailbox
            'License'           = $License
            'ServicePlan'       = $ServicePlan
        }

        # Build object and add to results array
        $Object = New-Object -TypeName PSObject -Property $Properties
        $Results += $Object
    }

    # Sort the results by OU and UPN.
    Write-Verbose -Message "Saving output file."
    $OrderedResults = $Results | Sort-Object -Property OU, UserPrincipalName

    # Export CSV
    $OrderedResults | Export-Csv -NoTypeInformation -Path $SaveFilePath

    # Display output location.
    Write-Host "CSV file exported to: $SaveFilePath"
}
End {
    Write-Verbose -Message "Removing session to Exchange Online."
    #Remove-PSSession -Session $Session -ErrorAction SilentlyContinue -Confirm:$FALSE | Out-Null
    Disconnect-ExchangeOnline -Confirm:$False -ErrorAction SilentlyContinue | Out-Null

    if ($Email) {
        Write-Verbose -Message "Sending email."
        $date = (Get-Date -Format MM/dd/yyyy)
        $MailSettings = @{
            To          = $SendTo
            From        = $SendFrom
            SmtpServer  = $SmtpServer
            Subject     = "Office 365 License Report - $date"
            Body        = "Attached is the Office 365 license report as of $date."
            Attachments = $SaveFilePath
        }
        Send-MailMessage @MailSettings
    }

    if ($DeleteOutputFile) {
        Write-Verbose -Message "Removing output file."
        Remove-Item -Path $SaveFilePath -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
}

# End of Script
