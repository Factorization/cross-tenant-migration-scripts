[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $URLs,

    [Parameter(Mandatory = $false)]
    [PSCredential]
    $Credential = (Get-Credential)
)
BEGIN {
    $ErrorActionPreference = "stop"
    Connect-AzureAd -Credential $Credential | Out-Null

    function ConnectPNPOnline ($URL, [PSCredential]$Credential) {
        Try {
            Disconnect-PnPOnline
        }
        Catch {}
        try {
            Connect-PnPOnline -Url $URL -Credentials $Credential
        }
        Catch {
            $err = $_
            Write-Host "Failed to connect to URL $URL. Error: $err" -ForegroundColor Red
            Throw "Error"
        }
    }

    function GetAzureAdGroupMember ($LoginName) {
        $ObjectId = ($LoginName -split "\|")[-1]
        try {
            $GroupMembers = Get-AzureAdGroupMember -ObjectId $ObjectId
            $GroupMembers = $GroupMembers.mail
        }
        Catch {
            $GroupMembers = @()
        }
        return $GroupMembers
    }

    function GetPnPDocumentLibraries() {
        $LibrariesToSkip = @("Documents", "Form Templates", "Site Assets", "Site Pages", "Style Library")
        $Libraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false } | Where-Object { $_.Title -notin $LibrariesToSkip }
        return $Libraries
    }

    function GetRoleAssignments($Library) {
        $RoleAssignments = $Library.RoleAssignments
        $PermissionCollection = @()
        Foreach ($RoleAssignment in $RoleAssignments) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

            $PrincipalType = $RoleAssignment.Member.PrincipalType
            $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name

            If ($PrincipalType -eq "SharePointGroup") {
                $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName | Where-Object Email | Select-Object -ExpandProperty Email
                if (($GroupMembers | Measure-Object | Select-Object -ExpandProperty Count) -eq 0) { continue }

                $PermissionCollection += [PSCustomObject]@{
                    URL           = $Library.RootFolder.ServerRelativeUrl
                    LibraryName   = $Library.Title
                    PrincipalName = $RoleAssignment.Member.Title
                    PrincipalType = $PrincipalType
                    Permissions   = $PermissionLevels -join " | "
                    Membership    = $GroupMembers -join " | "
                }
            }
            elseif ($PrincipalType -eq "SecurityGroup") {
                $GroupMembers = GetAzureAdGroupMember -LoginName $RoleAssignment.Member.LoginName
                if (($GroupMembers | Measure-Object | Select-Object -ExpandProperty Count) -eq 0) { continue }

                $PermissionCollection += [PSCustomObject]@{
                    URL           = $Library.RootFolder.ServerRelativeUrl
                    LibraryName   = $Library.Title
                    PrincipalName = $RoleAssignment.Member.Title
                    PrincipalType = $PrincipalType
                    Permissions   = $PermissionLevels -join " | "
                    Membership    = $GroupMembers -join " | "
                }
            }
            Else {
                $PermissionCollection += [PSCustomObject]@{
                    URL           = $Library.RootFolder.ServerRelativeUrl
                    LibraryName   = $Library.Title
                    PrincipalName = $RoleAssignment.Member.Title
                    PrincipalType = $PrincipalType
                    Permissions   = $PermissionLevels -join " | "
                    Membership    = "N/A"
                }
            }
        }
        return $PermissionCollection
    }
}
PROCESS {
    $Total = $URLs | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    $Results = @()
    foreach ($URL in $URLs) {
        Write-Progress -Id 1 -Activity "Exporting SharePoint Permissions..." -Status "Sites: [$i / $Total] | Site: $URL" -PercentComplete (($i / $Total) * 100)
        $i++
        try {
            ConnectPNPOnline -URL $URL -Credential $Credential
        }
        Catch {
            Continue
        }
        $Libraries = GetPnPDocumentLibraries
        $SubTotal = $Libraries | Measure-Object | Select-Object -ExpandProperty Count
        $j = 0
        foreach ($Library in $Libraries) {
            Write-Progress -Id 2 -ParentId 1 -Activity "Exporting Document Library Permissions..." -Status "Libraries: [$j / $SubTotal] | Library: $($Library.Title)" -PercentComplete (($j / $SubTotal) * 100)
            $j++
            $Library = Get-PnPList -Identity $Library.Title -Includes RoleAssignments
            $Results += GetRoleAssignments -Library $Library
        }
        Write-Progress -Id 2 -Completed -Activity "Exporting Document Library Permissions..."
    }
}
END {
    if ($Results) {
        $Date = Get-Date -Format yyyy-MM-dd_HH.mm
        $FileName = "SharePoint_Permissions_Export_$Date.csv"

        $Results | Export-Csv -NoTypeInformation -LiteralPath $FileName
        Write-Host "File exported to: $FileName" -ForegroundColor Green
    }
}
