[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $URLs,

    [Parameter(Mandatory = $true)]
    [string]
    $LibraryName,

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
        Connect-PnPOnline -Url $URL -Credentials $Credential
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

    function GetRoleAssignments($Library) {
        $RoleAssignments = $Library.RoleAssignments
        $PermissionCollection = @()
        Foreach ($RoleAssignment in $RoleAssignments) {
            #Get the Permission Levels assigned and Member
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

            $PrincipalType = $RoleAssignment.Member.PrincipalType
            $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name

            $PermissionLevels = $PermissionLevels | Where-Object { $_ -ne "Limited Access" }
            If (($PermissionLevels | Measure-Object | Select-Object -ExpandProperty Count) -eq 0) { Continue }

            If ($PrincipalType -eq "SharePointGroup") {
                $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName
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
        Write-Progress -Activity "Exporting SharePoint Permissions..." -Status "Site: $URL" -PercentComplete (($i / $Total) * 100)
        $i++
        ConnectPNPOnline -URL $URL -Credential $Credential
        $Library = Get-PnPList -Identity $LibraryName -Includes RoleAssignments
        if (-not $Library) {
            Write-Host " Library: $LibraryName | URL: $URL | Library doesn't exist." -ForegroundColor Red
            continue
        }
        $Results += GetRoleAssignments -Library $Library
    }
}
END {
    if ($Results) {
        $Date = Get-Date -Format yyyy-MM-dd_HH.mm
        $FileName = "SharePoint_Permissions_$($LibraryName)_Export_$Date.csv"

        $Results | Export-Csv -NoTypeInformation -LiteralPath $FileName
        Write-Host "File exported to: $FileName" -ForegroundColor Green
    }
}
