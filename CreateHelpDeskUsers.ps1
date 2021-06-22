[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $InputFile,

    [Parameter(Mandatory = $false)]
    [string]
    $Server = "dcc-h-dc01.ad.cannabis.ca.gov",

    [Parameter(Mandatory = $false)]
    [pscredential]
    $Credential = (Get-Credential -Message "DA")
)

BEGIN {
    $OU = "OU=DCA,OU=Standard,OU=Accounts,OU=DCC,DC=ad,DC=cannabis,DC=ca,DC=gov"
    function GetPassword() {
        $characters = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!@#$%')
        $Password = Get-Password -PasswordLength 15 -Count 1 -InputStrings $characters
        return $Password
    }

    $CSV = Import-Csv $InputFile
}
PROCESS {
    foreach ($User in $CSV) {
        $SamAccountName = $User.SamAccountName
        Try {
            $ADUser = Get-ADUser $SamAccountName -Server $Server
        }
        Catch {
            $ADUser = $null
        }

        if ($ADUser) {
            Write-Host "User $SamAccountName already exists." -ForegroundColor Red
            Continue
        }

        $Password = GetPassword
        $Attributes = @{
            Name              = $User.Name
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            Path              = $OU
            SamAccountName    = $SamAccountName
            AccountPassword   = $(ConvertTo-SecureString -AsPlainText "$Password" -Force)
            Enabled           = $true
            GivenName = $User.GivenName
            Surname = $User.SurName
        }
        New-ADUser @Attributes -Server $Server -Credential $Credential
        Write-Host "Created AD User $SamAccountName." -ForegroundColor Green
    }
}
END {

}
