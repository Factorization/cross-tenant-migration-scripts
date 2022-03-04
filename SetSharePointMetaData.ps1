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

    $CSV = Import-Csv -LiteralPath $CSVFile | Where-Object { $_.CustomersFullPath -like '*"*' } | Select-Object -First 1
}
PROCESS {

    $ErrorList = @()
    $Total = $CSV | Measure-Object | Select-Object -ExpandProperty Count
    $i = 0

    foreach ($Line in $CSV) {
        $RelativePath = $Line.RelativePath
        $VendorFullPath = $Line.VendorFullPath
        $VehicleFullPath = $Line.VehicleFullPath
        $CustomersFullPath = $Line.CustomersFullPath
        $ManufacturerFullPath = $Line.ManufacturerFullPath

        # Update relative path
        $RelativePath = $RelativePath -replace "^/teams/Acctg", "/sites/Taborda_Acctg"

        Write-Progress -Activity "Adding metadata to files..." -Status "Files: [ $i / $Total ] | Errors: $($ErrorList | Measure-Object | Select-Object -ExpandProperty Count) | Current File: $RelativePath" -PercentComplete (($i / $Total) * 100)
        $i++

        if (-not $RelativePath) {
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Relative Path can't be empty."
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }

        Try {
            $File = Get-PnPFile -Url $RelativePath -AsListItem
        }
        Catch {
            $err = $_
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Error getting PNP file. Error: $err"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
        if (-not $File) {
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Failed to find file. Skipping"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }

        if ($VendorFullPath) {
            $VendorLabel = ($VendorFullPath -split "\|")
            $VendorSearch = (($VendorLabel | ForEach-Object { "$_;#*|" }) -join "").TrimEnd("|")
            $Vendor_TermID = $TermIds | Where-Object { $_ -like $VendorSearch }

            if (($Vendor_TermID | Measure-Object | Select-Object -ExpandProperty Count) -ne 1) {
                $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Found multiple Vendor term IDs. Skipping"
                $ErrorList += $Line
                $Line | Out-Host
                Continue
            }
        }
        if ($VehicleFullPath) {
            $VehicleLabel = ($VehicleFullPath -split "\|")
            $VehicleSearch = (($VehicleLabel | ForEach-Object { "$_;#*|" }) -join "").TrimEnd("|")
            $Vehicle_TermID = $TermIds | Where-Object { $_ -like $VehicleSearch }

            if (($Vehicle_TermID | Measure-Object | Select-Object -ExpandProperty Count) -ne 1) {
                $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Found multiple Vehicle term IDs. Skipping"
                $ErrorList += $Line
                $Line | Out-Host
                Continue
            }
        }
        if ($CustomersFullPath) {
            $CustomersLabel = ($CustomersFullPath -split "\|")
            $CustomersSearch = (($CustomersLabel | ForEach-Object { "$_;#*|" }) -join "").TrimEnd("|")
            $Customers_TermID = $TermIds | Where-Object { $_ -like $CustomersSearch }

            if (($Customers_TermID | Measure-Object | Select-Object -ExpandProperty Count) -ne 1) {
                $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Found multiple Customers term IDs. Skipping"
                $ErrorList += $Line
                $Line | Out-Host
                Continue
            }
        }
        if ($ManufacturerFullPath) {
            $ManufacturerLabel = ($ManufacturerFullPath -split "\|")
            $ManufacturerSearch = (($ManufacturerLabel | ForEach-Object { "$_;#*|" }) -join "").TrimEnd("|")
            $Manufacturer_TermID = $TermIds | Where-Object { $_ -like $ManufacturerSearch }

            if (($Manufacturer_TermID | Measure-Object | Select-Object -ExpandProperty Count) -ne 1) {
                $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Found multiple Manufacturer term IDs. Skipping"
                $ErrorList += $Line
                $Line | Out-Host
                Continue
            }
        }

        $Values = @{}
        if ($Vendor_TermID) {
            $Vendor_TermID = ($Vendor_TermID -split "#")[-1]
            $Values["Vendor"] = $Vendor_TermID
        }
        if ($Vehicle_TermID) {
            $Vehicle_TermID = ($Vehicle_TermID -split "#")[-1]
            $Values["Vehicle"] = $Vehicle_TermID
        }
        if ($Customers_TermID) {
            $Customers_TermID = ($Customers_TermID -split "#")[-1]
            $Values["Customers"] = $Customers_TermID
        }
        if ($Manufacturer_TermID) {
            $Manufacturer_TermID = ($Manufacturer_TermID -split "#")[-1]
            $Values["Manufacturer"] = $Manufacturer_TermID
        }

        if ($Values) {
            try {
                Set-PnpListItem -List $DocumentLibrary -Identity $File.Id -Values @{"Vendor" = $TermID } | Out-Null
                Write-Host "Filename: $($Line.FileName)"
                Write-Host "`tVendor: $($VendorLabel[-1]) ($Vendor_TermID)"
                Write-Host "`tVehicle: $($VehicleLabel[-1]) ($Vehicle_TermID)"
                Write-Host "`tCustomers: $($CustomersLabel[-1]) ($Customers_TermID)"
                Write-Host "`tManufacturer: $($ManufacturerLabel[-1]) ($Manufacturer_TermID)"
                $Values | Out-Host
            }
            Catch {
                $err = $_
                $Line | Add-Member -MemberType NoteProperty -Name Error -Value "Error setting PNP list item. Error: $err"
                $ErrorList += $Line
                $Line | Out-Host
                Continue
            }
        }
        else {
            $Line | Add-Member -MemberType NoteProperty -Name Error -Value "No metadata values to set. Skipping"
            $ErrorList += $Line
            $Line | Out-Host
            Continue
        }
    }
}
END {
    if ($ErrorList) {
        $Date = Get-Date -Format "yyyy-MM-dd_HH.mm"
        $ErrorFile = "SharePoint_MetaData_Update_Error_List_$Date.csv"
        $ErrorList | Export-Csv -NoTypeInformation -LiteralPath $ErrorFile
        Write-Host "Error file saved to: $ErrorFile"
    }
}
