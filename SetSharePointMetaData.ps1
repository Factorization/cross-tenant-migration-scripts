[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]
    $DocumentLibrary,

    [Parameter(Mandatory = $true)]
    [string]
    $CSVFile
)
BEGIN {
    $ErrorActionPreference = "Stop"

    # Disconnect PNP Online
    try {
        Disconnect-PnPOnline
    }
    Catch {}

    # Connect PNP Online
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin

    # Get all Taxonomy IDs
    Write-Host "Getting all Taxonomies..." -NoNewline
    $TermIds = Export-PnPTaxonomy -IncludeID
    Write-Host "DONE" -ForegroundColor Green

    $CSV = Import-Csv -LiteralPath $CSVFile | select -First 1
}
PROCESS{

    $ErrorList = @()
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0

    foreach ($Line in $CSV){
        $RelativePath = $Line.RelativePath
        $VendorFullPath = $Line.VendorFullPath
        $VendorLabel = $Line.VendorLabel

        # Update relative path
        $RelativePath = $RelativePath -replace "^/teams/Acctg", "/sites/Taborda_Acctg"

        Write-Progress -Activity "Adding metadata to files..." -Status "Files: [ $i / $Total ] | Errors: $($ErrorList | Measure-Object | Select-Object -ExpandProperty Count) | Current File: $RelativePath" -PercentComplete (($i / $Total) * 100)
        $i++

        $line | out-host

        if(-not $RelativePath){
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Relative Path can't be empty."
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }

        if(-not $VendorFullPath){
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Vendor Path is empty. Skipping."
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
        Try{
            $File = Get-PnPFile -Url $RelativePath -AsListItem
        }
        Catch{
            $err = $_
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Error getting PNP file. Error: $err"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
        if(-not $File){
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Failed to find file. Skipping"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }

        $TermID = $TermIds | Where-Object {$_ -like "*|$VendorLabel;#*"}
        if(-not $TermId){
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Failed to find term ID. Skipping"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
        if(($TermID | Measure-Object | Select-Object -ExpandProperty Count) -ne 1){
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Found multiple term IDs. Skipping"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
        $TermID = ($TermId -split "#")[-1]

        try{
            Set-PnpListItem -List $DocumentLibrary -Identity $File.Id -Values @{"Vendor" = $TermID} | Out-Null
        }
        Catch{
            $err = $_
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Error setting PNP list item. Error: $err"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
    }
}
END{
    if($ErrorList){
        $Date = Get-Date -Format "yyyy-MM-dd_HH.mm"
        $ErrorFile = "SharePoint_MetaData_Update_Error_List_$Date.csv"
        $ErrorList | Export-Csv -NoTypeInformation -LiteralPath $ErrorFile
        Write-Host "Error file saved to: $ErrorFile"
    }
}
