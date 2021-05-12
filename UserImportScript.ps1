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
    $InputFiles
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
    $MasterEquipmentMailboxFileOutput = Join-Path $Root "Master _Equipment_Mailbox_File_$DATA.csv"
    $MasterRoomMailboxFileOutput = Join-Path $Root "Master_Room_Mailbox_File_$DATE.csv"
    $ErrorFileOutput = Join-Path $Root "Error_File_$DATE.csv"
    $SuccessFileOutput = Join-Path $Root "Success_File_$DATE.csv"

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
    function ImportMasterFile($File, $Name) {
        if (-not $File) {
            WriteLog "No $Name master file provided."
            return @()
        }
        WriteLog "Importing master file $File..."
        return Import-Csv -LiteralPath $File
    }

    function ImportInputFiles() {
        $Results = @()
        foreach ($File in $InputFiles) {
            WriteLog "Importing input file '$File'..."
            $Results += Import-Csv -LiteralPath $File
        }
        return $Results
    }
    function CreateDirectory($Path) {
        WriteLog "Creating directory '$Path'..."
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    function CreateFile($Path) {
        WriteLog "Creating file '$Path'..."
        New-Item -Path $Path -ItemType File -Force | Out-Null
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
    CreateFile -Path $SuccessFileOutput
    WriteLog "Done creating output files."

    # Import master files for User, Share, Equipment and Room
    WriteLog "Importing master files..."
    $MasterUsers = ImportMasterFile -File $MasterUserFile -Name Users
    $MasterSharedMailboxes = ImportMasterFile -File $MasterSharedMailboxFile -Name Shared
    $MasterEquipmentMailboxes = ImportMasterFile -File $MasterEquipmentMailboxFile -Name Equipment
    $MasterRoomMailboxes = ImportMasterFile -File $MasterRoomMailboxFile -Name Room
    WriteLog "Done importing master files."

    # Import input files
    WriteLog "Importing input files..."
    $InputDate = ImportInputFiles
    WriteLog "Done importing input files."


    # Check if user in master file
    # if in master file, get ad user
    # if get user fails, log error and continue
    # if not in master file, confirm get ad user not present
    # if user is found, log error and continue
    # check id user mailbox or shared mailbox
    # get or create ad user (whats the minimum properties needed)
    # decide which OU
    # Set AD user properties
    # which properties need to be transformed (UPN, mail, displayname, )



}
END {

}
