[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $OktaGroupID,

    [Parameter(Mandatory = $true)]
    [string]
    $OktaGroupName,

    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath = $PSScriptRoot
)
BEGIN {
    $Global:ErrorActionPreference = 'STOP'

    try {
        Get-Command "Get-Mailbox" | Out-Null
    }
    catch {
        Write-Error "Not connected to Exchange Online. You must connect to Exchange Online first."
        exit
    }
    if (-not(Test-Path $OutputPath)){
        throw "Output directory '$OutputPath' doesn't exist. Please create this directory and tyr again."
    }
}
PROCESS {
    Write-Verbose "Getting Okta group members for $OktaGroupID..."
    $OktaGroupMembers = (oktaGetGroupMembersbyId -gid $OktaGroupID).Profile
    $Total = $OktaGroupMembers | Measure-Object | Select-Object -ExpandProperty Count
    Write-Verbose "Found $total members."
    $i = 0
    $Properties = @()
    $OktaGroupMembers | ForEach-Object { $Properties += ($_ | Get-Member -Type NoteProperty).Name }
    $Properties = $Properties | Sort-Object -Unique
    $Results = @()
    foreach ($User in $OktaGroupMembers) {
        $Email = $user.secondEmail
        Write-Progress -Activity "Exporting Proxy Addresses" -Status "User: $Email | Status: ($i / $total)" -PercentComplete "$(($i / $Total)*100)" -CurrentOperation "User: $Email"
        $i++
        $NewProxyAddresses = @($User.proxyAddresses)
        if (-not $NewProxyAddresses) {
            $NewProxyAddresses = @()
        }

        Write-Verbose "Getting mailbox for user $Email..."
        try{
            $MB = Get-MailBox $Email -ErrorAction Stop
        }
        Catch{
            Write-Host "Can't find user $email in EXO. Skipping user." -ForegroundColor Red
            continue
        }
        $NewProxyAddresses += "x500:$($MB.LegacyExchangeDN)"
        $NewProxyAddresses += $MB.EmailAddresses | Where-Object { $_ -like "x500:*" } | ForEach-Object { $_ -creplace "^X500:", "x500:" }
        $NewProxyAddresses = $NewProxyAddresses | Sort-Object -Unique

        Write-Verbose "Creating results object..."
        $Result = [PSCustomObject]@{
            "login"             = $User.login
            "firstName"         = $User.firstName
            "lastName"          = $User.lastName
            "middleName"        = $User.middleName
            "honorificPrefix"   = $User.honorificPrefix
            "honorificSuffix"   = $User.honorificSuffix
            "email"             = $User.email
            "title"             = $User.title
            "displayName"       = $User.displayName
            "nickName"          = $User.nickName
            "profileUrl"        = $User.profileUrl
            "secondEmail"       = $User.secondEmail
            "mobilePhone"       = $User.mobilePhone
            "primaryPhone"      = $User.primaryPhone
            "streetAddress"     = $User.streetAddress
            "city"              = $User.city
            "state"             = $User.state
            "zipCode"           = $User.zipCode
            "countryCode"       = $User.countryCode
            "postalAddress"     = $User.postalAddress
            "preferredLanguage" = $User.preferredLanguage
            "locale"            = $User.locale
            "timezone"          = $User.timezone
            "userType"          = $User.userType
            "employeeNumber"    = $User.employeeNumber
            "costCenter"        = $User.costCenter
            "organization"      = $User.organization
            "division"          = $User.division
            "department"        = $User.department
            "managerId"         = $User.managerId
            "manager"           = $User.manager
            "proxyAddresses"    = $NewProxyAddresses -join ";"
        }
        Write-Verbose "Adding to results..."
        $Results += $Result

    }
}
END {
    $DATE = get-date -Format yyyy-MM-dd_HH.mm
    $OutputName = "Okta_$($OktaGroupName)_$($Date).csv"
    $OutputName = Join-Path $OutputPath $OutputName
    Write-Verbose "Exporting results to $OutputName..."
    $Results | Export-Csv -NoTypeInformation -LiteralPath $OutputName
    Write-Host "File saved to: $OutputName" -ForegroundColor Green
}
