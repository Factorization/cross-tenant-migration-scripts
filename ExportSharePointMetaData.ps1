[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]
    $DocumentLibrary
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
    $sourceTermIds = Export-PnPTaxonomy -IncludeID
    Write-Host "DONE" -ForegroundColor Green

    function LookupTerm($TermGuid) {
        $Result = $sourceTermIds | Where-Object { ($_ -split ";")[-1] -eq "#$TermGuid" }
        $Result = ($Result -split ";" | ForEach-Object { $_ -split "\|" | Where-Object { $_ -notlike "#*" } }) -join "|"
        return $Result
    }
}
PROCESS {
    Write-Host "Getting files in $DocumentLibrary..." -NoNewline
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl "Bill Pay" -ItemType File
    Write-Host "DONE" -ForegroundColor Green
    
    $Total = $Files | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    foreach ($File in $files) {
        Write-Progress -Activity "Getting file metadata..." -Status "Files: [ $i / $Total ]" -PercentComplete (($i / $Total) * 100)
        $i++

        $Name = $File.Name
        $RelativePath = $File.ServerRelativeUrl

        $PNPFileListItem = Get-PnPFile -AsListItem -Url $RelativePath

        $Vendor_Value = $PNPFileListItem.FieldValues["Vendor"]
        $Vendor_Label = $Vendor_Value.Label
        $Vendor_Guid = $Vendor_Value.TermGuid
        $Vendor_FullPath = LookupTerm -TermGuid $Vendor_Guid

        [PSCustomObject]@{
            FileName = $Name
            RelativePath = $RelativePath
            VendorLabel = $Vendor_Label
            VendorFullPath = $Vendor_FullPath
        }
    }
}
