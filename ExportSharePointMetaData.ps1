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

    function GetFieldValue($Value, $FileListItem) {
        $Field_Value = $FileListItem.FieldValues[$Value]
        $Field_Label = $Field_Value.Label
        $Field_Guid = $Field_Value.TermGuid
        $Field_FullPath = LookupTerm -TermGuid $Field_Guid

        $Results = [PSCustomObject]@{
            Value    = $Field_Value
            Label    = $Field_Label
            Guid     = $Field_Guid
            FullPath = $Field_FullPath
        }
        return $Results
    }
}
PROCESS {
    Write-Host "Getting files in $DocumentLibrary..." -NoNewline
    $Files = Get-PnPFolderItem -FolderSiteRelativeUrl $DocumentLibrary -ItemType File
    Write-Host "DONE" -ForegroundColor Green

    $Total = $Files | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0
    foreach ($File in $files) {
        Write-Progress -Activity "Getting file metadata..." -Status "Files: [ $i / $Total ]" -PercentComplete (($i / $Total) * 100)
        $i++

        $Name = $File.Name
        $RelativePath = $File.ServerRelativeUrl

        $PNPFileListItem = Get-PnPFile -AsListItem -Url $RelativePath

        $Vendor = GetFieldValue -Value "Vendor" -FileListItem $PNPFileListItem
        $Vehicle = GetFieldValue -Value "Vehicle" -FileListItem $PNPFileListItem
        $Customers = GetFieldValue -Value "Customers" -FileListItem $PNPFileListItem
        $Manufacturer = GetFieldValue -Value "Product" -FileListItem $PNPFileListItem

        [PSCustomObject]@{
            FileName             = $Name
            RelativePath         = $RelativePath
            VendorLabel          = $Vendor.Label
            VendorFullPath       = $Vendor.FullPath
            VehicleLabel         = $Vehicle.Label
            VehicleFullPath      = $Vehicle.FullPath
            CustomersLabel       = $Customers.Label
            CustomersFullPath    = $Customers.FullPath
            ManufacturerLabel    = $Manufacturer.Label
            ManufacturerFullPath = $Manufacturer.FullPath
        }
    }
}
