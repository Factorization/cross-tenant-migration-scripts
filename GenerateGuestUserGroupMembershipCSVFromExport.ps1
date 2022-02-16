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
    $User_XML_Directory = Join-Path $exportPath "Guest_Azure_AD_User_Output_XMLs"
    $XMLs = Get-ChildItem -LiteralPath $XML_Directory -File *.xml
    $User_XMLs = Get-ChildItem -LiteralPath $User_XML_Directory -File *.xml
}
PROCESS {
    $Results = @()
    foreach ($XML in $XMLs) {
        $User_XML = $User_XMLs | Where-Object {$_.Name -eq $XML.Name}
        $AzureADUser = Import-Clixml $User_XML.FullName
        $User = $AzureADUser.Mail
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
