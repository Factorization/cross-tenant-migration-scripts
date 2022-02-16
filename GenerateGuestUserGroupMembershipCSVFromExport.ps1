[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [String]
    $ExportPath,

    [Parameter(Mandatory = $false)]
    [string[]]
    $ExcludeGroups = @()
)
BEGIN {
    $XML_Directory = Join-Path $exportPath "Guest_Azure_AD_User_Group_Membership_Output_XMLs"
    $XMLs = Get-ChildItem -LiteralPath $XML_Directory -File *.xml
}
PROCESS {
    $Results = @()
    foreach ($XML in $XMLs) {
        $User = ($XML.basename -split "#EXT#")[0]
        $IndexOfUnderscore = $User.LastIndexOf("_")
        $User = $User.remove($IndexOfUnderscore, 1).insert($IndexOfUnderscore, "@")
        $Groups = Import-Clixml $XML.FullName | Where-Object { $_.DisplayName -notin $ExcludeGroups }
        foreach ($Group in $Groups) {
            $Results += [PSCustomObject]@{
                GuestUser  = $User
                GroupName  = $Group.DisplayName
                GroupEmail = $Group.Mail
            }
        }
    }
    $OutputFile = Join-Path $ExportPath "CSV_Exports\Guest_Users_Group_Membership.csv"
    $Results | Export-Csv -NoTypeInformation -LiteralPath $OutputFile
    Write-host "File exported to $OutputFile" -ForegroundColor Green
}
END {}
