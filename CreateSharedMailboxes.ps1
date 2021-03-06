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
    $COMPANY = "California Department of Cannabis Control"
    $DEPARTMENT = "DCC"
    $OU = "OU=Shared-Mailbox,OU=Accounts,OU=DCC,DC=ad,DC=cannabis,DC=ca,DC=gov"
    function GetPassword() {
        $characters = @('abcdefghijkmnpqrstuvwxyz', 'ABCEFGHJKLMNPQRSTUVWXYZ', '23456789', '!@#$%')
        $Password = Get-Password -PasswordLength 15 -Count 1 -InputStrings $characters
        return $Password
    }

    $CSV = Import-Csv $InputFile

    Connect-DCCExchange
}
PROCESS {
    foreach ($mailbox in $CSV) {
        $Email = $mailbox.Email
        $SamAccountName = $mailbox.SamAccountName
        $DisplayName = $mailbox.DisplayName
        $Description = $mailbox.Description
        $ADUser = Get-ADUser -Filter "Mail -eq '$Email'" -Server $Server

        if ( ($ADUser | Measure-Object | Select-Object -ExpandProperty Count) -gt 1 ) {
            Write-Host "Mailbox $Email found multiple." -ForegroundColor Red
            Continue
        }

        if (-not $ADUser) {
            $Password = GetPassword
            $Attributes = @{
                Name              = $DisplayName
                DisplayName       = $DisplayName
                UserPrincipalName = $Email
                Path              = $OU
                SamAccountName    = $SamAccountName
                AccountPassword   = $(ConvertTo-SecureString -AsPlainText "$Password" -Force)
                Company           = $COMPANY
                Department        = $DEPARTMENT
                Enabled           = $false
                Description       = $Description
            }
            New-ADUser @Attributes -Server $Server -Credential $Credential
            Write-Host "Created AD User $Email." -ForegroundColor Green
        }
        else {
            Write-Host "AD User for $Email already exists." -ForegroundColor Cyan
        }

        $User = Get-User $Email
        if (-not $User) {
            Write-Host "Can't find $Email on exchange. Try again later."
            Continue
        }
        if ($user.RecipientTypeDetails -eq "DisabledUser") {
            $RemoteAddress = ($Email -split '@')[0] + "@dcco365.mail.onmicrosoft.com"
            $User | Enable-RemoteMailbox -RemoteRoutingAddress $RemoteAddress -Shared | Out-Null
            $User | Set-RemoteMailbox -HiddenFromAddressListsEnabled $true | Out-Null
            Write-Host "Created remote mailbox for $Email" -ForegroundColor Green
        }
        else {
            Write-Host "Remote mailbox for $Email already exists." -ForegroundColor Cyan
        }
    }
}
