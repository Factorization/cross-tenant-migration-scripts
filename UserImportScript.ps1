[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]
    $MasterUserFile = $null,

    [Parameter(Mandatory = $false)]
    [string]
    $MasterSharedMailboxFile = $null,

    [Parameter(Mandatory = $false)]
    [string]
    $MasterEquipmentMailboxFile = $null,

    [Parameter(Mandatory = $false)]
    [string]
    $MasterRoomMailboxFile = $null,

    [Parameter(Mandatory = $false)]
    [string]
    $OutputFolder = $PSScriptRoot,

    [Parameter(Mandatory = $true)]
    [string[]]
    $InputFiles,

    [Parameter(Mandatory = $true)]
    [string]
    $Server
)
BEGIN {
    #########################
    #        GLOBALS        #
    #########################
    $ErrorActionPreference = "stop"
    $DATE = Get-Date -Format yyyy-MM-dd_HH.mm
    $Root = Join-Path $OutputFolder "User_Import_$Date"
    $BackupDir = Join-Path $Root "Backup"
    $LogFile = Join-Path $Root "Log_File_$DATE.log"
    $MasterUserFileOutput = Join-Path $Root "Master_Users_File_$DATE.csv"
    $MasterSharedMailboxFileOutput = Join-Path $Root "Master_Shared_Mailbox_File_$DATE.csv"
    $MasterEquipmentMailboxFileOutput = Join-Path $Root "Master_Equipment_Mailbox_File_$DATA.csv"
    $MasterRoomMailboxFileOutput = Join-Path $Root "Master_Room_Mailbox_File_$DATE.csv"
    $ErrorFileOutput = Join-Path $Root "Error_File_$DATE.csv"
    # $SuccessFileOutput = Join-Path $Root "Success_File_$DATE.csv"

    # Map of user objects to specific OUs
    $OU_MAP = @{
        "UserMailbox"      = @{
            "DCA"  = "USER/DCA"
            "CDFA" = "USER/CDFA"
            "CDPH" = "USER/CDPH"
        }
        "SharedMailbox"    = @{
            "DCA"  = "Shared/DCA"
            "CDFA" = "Shared/CDFA"
            "CDPH" = "Shared/CDPH"
        }
        "EquipmentMailbox" = @{
            "DCA"  = "Equip/DCA"
            "CDFA" = "Equip/CDFA"
            "CDPH" = "Equip/CDPH"
        }
        "RoomMailbox"      = @{
            "DCA"  = "Room/DCA"
            "CDFA" = "Room/CDFA"
            "CDPH" = "Room/CDPH"
        }
    }

    # Required properties for input files
    $Required_Input_File_Properties = @(
        'AcceptMessagesOnlyFrom',
        'AcceptMessagesOnlyFromDLMembers',
        'AcceptMessagesOnlyFromSendersOrMembers',
        'Alias',
        'ArchiveGuid',
        'City',
        'Company',
        'Department',
        'Description',
        'DisplayName',
        'Division',
        'EmailAddresses',
        'EmployeeID',
        'EmployeeNumber',
        'ExchangeGuid',
        'extensionAttribute1',
        'extensionAttribute10',
        'extensionAttribute11',
        'extensionAttribute12',
        'extensionAttribute13',
        'extensionAttribute14',
        'extensionAttribute15',
        'extensionAttribute2',
        'extensionAttribute3',
        'extensionAttribute4',
        'extensionAttribute5',
        'extensionAttribute6',
        'extensionAttribute7',
        'extensionAttribute8',
        'extensionAttribute9',
        'Fax',
        'FirstName',
        'HiddenFromAddressListsEnabled',
        'HomePhone',
        'LastName',
        'LegacyExchangeDN',
        'LitigationHoldDate',
        'LitigationHoldDuration',
        'LitigationHoldEnabled',
        'LitigationHoldOwner',
        'MailboxItemCount',
        'MailboxSize',
        'Manager',
        'MobilePhone',
        'Name',
        'Office',
        'Organization',
        'OtherFax',
        'OtherHomePhone',
        'OtherTelephone',
        'Phone',
        'PostalCode',
        'PrimarySmtpAddress',
        'RecipientType',
        'RecipientTypeDetails',
        'SamAccountName',
        'State',
        'StreetAddress',
        'Title',
        'UserPrincipalName'
    )

    # Required properties for master files
    $Required_Master_File_Properties = @(
        "OldUPN",
        "UPN",
        "FirstName",
        "LastName",
        "Password",
        "CreatedDate"
    )

    #########################
    #       FUNCTIONS       #
    #########################
    Function WriteLog {
        Param ([string]$Message, [Switch]$isError, [Switch]$isWarning)

        # Enure log file exists
        if (-not (Test-Path -LiteralPath $LogFile -PathType Leaf)) {
            New-Item -Path $LogFile -ItemType File -Force | Out-Null
        }

        # Get the current date
        [string]$date = Get-Date -Format G

        # Write everything to our log file
        ( "[" + $date + "]`t" + $Message) | Out-File -FilePath $LogFile -Append

        # Write verbose output
        if ($isError) {
            Write-Error -Message $Message
        }
        elseif ($isWarning) {
            Write-Warning -Message $Message
        }
        else {
            Write-Verbose -Message $Message
        }
    }
    function WriteCSVOutput($Data, $File) {
        $Data | Export-Csv -NoTypeInformation -LiteralPath $File -Append
    }
    function WriteMasterFile($Data, $MailboxType) {
        if ($MailboxType -eq "UserMailbox") {
            $File = $MasterUserFileOutput
        }
        elseif ($MailboxType -eq "SharedMailbox") {
            $File = $MasterSharedMailboxFileOutput
        }
        elseif ($MailboxType -eq "EquipmentMailbox") {
            $File = $MasterEquipmentMailboxFileOutput
        }
        elseif ($MailboxType -eq "RoomMailbox") {
            $File = $MasterRoomMailboxFileOutput
        }
        WriteCSVOutput -Data $Data -File $File
    }
    function ImportMasterFile($File, $Name, $OutputFile) {
        if (-not $File) {
            WriteLog "No $Name master file provided."
            return @()
        }
        WriteLog "Importing $Name master file $File..."
        $Results = Import-Csv -LiteralPath $File
        if (TestRequiredProperties -Data $Results -RequiredProperties $Required_Master_File_Properties) {
            $Results | Export-Csv -NoTypeInformation -LiteralPath $OutputFile
            return $Results
        }
        WriteLog "$Name Master file '$File' incorrect CSV columns. Exiting..." -isError
        Exit
    }

    function ImportInputFiles() {
        $Results = @()
        foreach ($File in $InputFiles) {
            WriteLog "Importing input file '$File'..."
            $Result = Import-Csv -LiteralPath $File
            if (TestRequiredProperties -Data $Result -RequiredProperties $Required_Input_File_Properties) {
                $Results += $Result
            }
            else {
                WriteLog "Input file '$File' incorrect CSV columns. Exiting..." -isError
                Exit
            }
        }
        return $Results
    }
    function TestRequiredProperties($Data, $RequiredProperties) {
        foreach ($Item in $Data) {
            $Properties = $Item | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | Sort-Object
            $isEqual = @(Compare-Object $Properties $RequiredProperties).Length -eq 0
            if ( -not $isEqual) {
                return $false
            }
        }
        return $true
    }
    function CreateDirectory($Path) {
        WriteLog "Creating directory '$Path'..."
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    function CreateFile($Path) {
        WriteLog "Creating file '$Path'..."
        New-Item -Path $Path -ItemType File -Force | Out-Null
    }
    function ImportUser($Data) {
        $OldUPN = $Data.UserPrincipalName
        $MailboxType = $Data.RecipientTypeDetails
        if ($MailboxType -eq "UserMailbox") {
            $NewUPN = CheckUserInMasterFile -OldUPN $OldUPN -MasterFile $MasterUsers
        }
        elseif ($MailboxType -eq "SharedMailbox") {
            $NewUPN = CheckUserInMasterFile -OldUPN $OldUPN -MasterFile $MasterSharedMailboxes
        }
        elseif ($MailboxType -eq "EquipmentMailbox") {
            $NewUPN = CheckUserInMasterFile -OldUPN $OldUPN -MasterFile $MasterEquipmentMailboxes
        }
        elseif ($MailboxType -eq "RoomMailbox") {
            $NewUPN = CheckUserInMasterFile -OldUPN $OldUPN -MasterFile $MasterRoomMailboxes
        }
        else {
            $err = "Invalid mailbox type for '$OldUPN' and type '$MailboxType'."
            WriteLog $err
            $Data | Add-Member -MemberType NoteProperty -Name Error -Value $err
            WriteCSVOutput -Data $Data -File $ErrorFileOutput
            return $false
        }

        # Create user
        if (-not $NewUPN) {
            WriteLog "Creating AD user for $MailboxType..."
            $NewUPN = CreateADUser -Data $Data
            WriteLog "Done creating AD user."
            if ($NewUPN -eq $false) {
                return $false
            }
        }
        # Update user attributes
        # $Result = SetUserAttributes -Data $Date -NewUPN $NewUPN
        # return $Result

    }
    function CheckUserInMasterFile($OldUPN, $MasterFile) {
        If (-not $MasterFile) {
            return
        }
        $NewUPN = $MasterFile | Where-Object { $_.OldUPN -eq "$OldUPN" } | Select-Object -ExpandProperty UPN
        return $NewUPN
    }
    function SetUserAttributes($Data, $NewUPN) {
        $ADUser = GetADUser -NewUPN $NewUPN
        if ($ADUser -eq $false) {
            $err = "Failed to find AD user with UPN '$NewUPN'."
            WriteLog $err
            $Data | Add-Member -MemberType NoteProperty -Name Error -Value $err
            WriteCSVOutput -Data $Data -File $ErrorFileOutput
            return $false
        }
    }
    function CreateADUser($Data) {
        Try {
            $OldUPN = $Data.UserPrincipalName
            $MailboxType = $Data.RecipientTypeDetails
            $NewUPN = GetUPN -OldUPN $OldUPN
            $NewSamAccountName = ($NewUPN -split "@")[0]
            $OU = GetOU -OldUPN $OldUPN -MailboxType $MailboxType
            $NewDisplayName = GetDisplayName -OldDisplayName $Data.DisplayName
            $Password = GetPassword
            $FirstName = $Data.FirstName
            $LastName = $Data.LastName
            $Attributes = @{
                Name              = $NewDisplayName
                DisplayName       = $NewDisplayName
                UserPrincipalName = $NewUPN
                Path              = $OU
                SamAccountName    = $NewSamAccountName
                AccountPassword   = (ConvertTo-SecureString -AsPlainText $Password -Force)
                OtherAttributes   = @{'msExchHideFromAddressLists' = "$true" }
            }
            if ($MailboxType -eq "UserMailbox") {
                $Attributes.Enabled = $true
                $Attributes.ChangePasswordAtLogon = $true
            }
            else {
                $Attributes.Enabled = $false
            }
            if ($FirstName) {
                $Attributes.GivenName = $FirstName
            }
            if ($LastName) {
                $Attributes.SurName = $LastName
            }
            # New-ADUser @Attributes -Server $Server

            $Result = [PSCustomObject]@{
                OldUPN      = $OldUPN
                UPN         = $NewUPN
                FirstName   = $FirstName
                LastName    = $LastName
                Password    = $Password
                CreatedDate = $DATE
            }
            WriteLog "Writing master file..."
            WriteMasterFile -Data $Result -MailboxType $MailboxType
            WriteLog "Done writing master file."

            return $NewUPN
        }
        Catch {
            $err = $_
            WriteLog "Error creating AD user. Error: $err"
            $Data | Add-Member -MemberType NoteProperty -Name Error -Value $err
            WriteCSVOutput -Data $Data -File $ErrorFileOutput
            return $false
        }
    }
    function GetUPN($OldUPN) {
        $Prefix = ($OldUPN -split "@")[0]
        if (-not $Prefix) {
            Throw "UPN prefix error for '$OldUPN'."
        }
        return $Prefix + "@cannabis.ca.gov"
    }
    function GetOU($OldUPN, $MailboxType) {
        $Suffix = ($OldUPN -split "@")[1]
        if ($Suffix -match "^dca\.ca\.gov$|^dcao365\.onmicrosoft\.com$") {
            $Location = "DCA"
        }
        elseif ($Suffix -eq "cdph.ca.gov") {
            $Location = "CDPH"
        }
        elseif ($Suffix -eq "cdfa.ca.gov") {
            $Location = "CDFA"
        }
        else {
            Throw "Failed to parse OU from old UPN for '$OldUPN', mailbox type '$MailboxType'."
        }
        $OU = $OU_MAP.$MailboxType.$Location
        if (-not $OU) {
            Throw "Failed to parse OU from old UPN for '$OldUPN', mailbox type '$MailboxType'."
        }
        return $OU
    }
    function GetDisplayName($OldDisplayName) {
        $NewDisplayName = $OldDisplayName
        if ($NewDisplayName -like "*@*") {
            $NewDisplayName = ($NewDisplayName -split "@")[0] + "@Cannabis"
        }
        if ($NewDisplayName -like "*DCA*") {
            $NewDisplayName = $NewDisplayName -replace "DCA", "DCC"
        }
        if ($NewDisplayName -like "*CDFA*") {
            $NewDisplayName = $NewDisplayName -replace "CDFA", "DCC"
        }
        if ($NewDisplayName -like "*CDPH*") {
            $NewDisplayName = $NewDisplayName -replace "CDPH", "DCC"
        }
        if (-not $NewDisplayName) {
            Throw "Failed to parse display name '$OldDisplayName'."
        }
        return $NewDisplayName
    }
    function GetPassword() {
        $characters = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!@#$%')
        $Password = Get-Password -PasswordLength 15 -Count 1 -InputStrings $characters
        return $Password
    }
    function GetADUser($NewUPN) {
        Try {
            $ADUser = Get-ADUser -Filter "UserPrincipalName -eq '$NewUPN'" -Server $Server -Properties *
            if (-not $ADUser) {
                throw "Failed to find AD User with UPN '$NewUPN'"
            }
            return $ADUser
        }
        Catch {
            $err = $_
            WriteLog "Error getting AD user. Error: $err"
            return $false
        }
    }

}
PROCESS {
    # Script start
    WriteLog "Script starting..."

    # Create directories
    WriteLog "Creating directories..."
    CreateDirectory -Path $Root
    CreateDirectory -Path $BackupDir
    WriteLog "Done creating directories."

    # Create output files
    WriteLog "Creating output files..."
    CreateFile -Path $MasterUserFileOutput
    CreateFile -Path $MasterSharedMailboxFileOutput
    CreateFile -Path $MasterEquipmentMailboxFileOutput
    CreateFile -Path $MasterRoomMailboxFileOutput
    CreateFile -Path $ErrorFileOutput
    WriteLog "Done creating output files."

    # Import master files for User, Share, Equipment and Room
    WriteLog "Importing master files..."
    $MasterUsers = ImportMasterFile -File $MasterUserFile -Name Users -OutputFile $MasterUserFileOutput
    $MasterSharedMailboxes = ImportMasterFile -File $MasterSharedMailboxFile -Name Shared -OutputFile $MasterSharedMailboxFileOutput
    $MasterEquipmentMailboxes = ImportMasterFile -File $MasterEquipmentMailboxFile -Name Equipment -OutputFile $MasterEquipmentMailboxFileOutput
    $MasterRoomMailboxes = ImportMasterFile -File $MasterRoomMailboxFile -Name Room -OutputFile $MasterRoomMailboxFileOutput
    WriteLog "Done importing master files."

    # Import input files
    WriteLog "Importing input files..."
    $InputData = ImportInputFiles
    WriteLog "Done importing input files."

    # Loop over input data
    $Total = $InputData | Measure-Object | Select-Object -ExpandProperty Count
    $ErrorCount = 0
    $i = 0

    WriteLog "Importing users..."
    foreach ($Data in $InputData) {
        WriteLog "Importing user '$($Data.PrimarySmtpAddress)..."
        Write-Progress -Id 1 -Activity "Importing users..." -Status "Users: [$i/$Total] | Errors: $ErrorCount" -PercentComplete ($i / $Total * 100)
        $i++

        $Result = ImportUser -Data $Data
        if ($Result -eq $false) {
            $ErrorCount += 1
        }
        WriteLog "Done importing user."
    }
    WriteLog "Done importing users."
    # Check if user in master file
    # if in master file, get ad user
    # if get user fails, log error and continue
    # if not in master file, confirm get ad user not present
    # if user is found, log error and continue
    # check if user mailbox or shared mailbox
    # get or create ad user (whats the minimum properties needed)
    # decide which OU
    # Set AD user properties
    # which properties need to be transformed (UPN, mail, displayname, )



}
END {
    WriteLog "Script finished."
}
