[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [String]
    $ExportPath
)
BEGIN{
    $XML_Directory = Join-Path $exportPath "Guest_Azure_AD_User_Output_XMLs"
    $XMLs = Get-ChildItem -LiteralPath $XML_Directory -File *.xml
    $Guests = $XMLs | ForEach-Object {Import-Clixml $_.FullName}
}
PROCESS{
    $OutputFile = Join-Path $ExportPath "CSV_Exports\Guest_Users.csv"
    $Guests | Select-Object Mail, DisplayName | Export-Csv -NoTypeInformation -LiteralPath $OutputFile
    Write-host "File exported to $OutputFile" -ForegroundColor Green
}
END{}
