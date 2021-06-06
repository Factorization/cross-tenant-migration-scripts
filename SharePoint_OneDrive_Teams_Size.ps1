[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $SharePointSitesExports,

    [Parameter(Mandatory = $true)]
    [string[]]
    $UserExports,

    [Parameter(Mandatory = $true)]
    [string]
    $ShareTeamsUrlListFile
)
BEGIN {}
PROCESS {
    $DATE = Get-Date -Format yyyy-MM-dd_HH.mm

    $Users = $UserExports | ForEach-Object {
        if ($_.EndsWith("csv")) {
            Import-CSV $_
        }
        else {
            Import-Clixml $_
        }
    }
    $Users_DisplayName = $Users.DisplayName
    $SharePoint = $SharePointSitesExports | ForEach-Object {
        if ($_.EndsWith("csv")) {
            Import-CSV $_
        }
        else {
            Import-Clixml $_
        }
    }
    $SharePoint = $SharePoint | Group-Object -Property "Title" -AsHashTable

    $OneDrives = @()
    foreach ($User in $Users_DisplayName) {
        $OneDrives += $SharePoint[$User]
    }
    $NumOneDrives = $OneDrives | Measure-Object | Select-Object -ExpandProperty Count
    $SizeOneDrives = ($OneDrives | Measure-Object -Sum StorageUsageCurrent | Select-Object -ExpandProperty Sum) * 1mb / 1GB

    Write-Host "OneDrive Count: $NumOneDrives"
    Write-Host "OneDrive Size: $([math]::Ceiling($SizeOneDrives))GB"
    $OneDrives | Select-Object Title, URL, StorageUsageCurrent | Sort-Object -Descending -Property {[int]$_.StorageUsageCurrent} | Export-Csv -NoTypeInformation -LiteralPath "OneDrive_Exports_$DATE.csv"

    $Sites = Get-Content -LiteralPath $ShareTeamsUrlListFile

    $SharePointAndTeams = @()
    foreach ($Site in $Sites){
        $SharePointAndTeams += $SharePoint[$site]
    }
    $NumSharePointAndTeams = $SharePointAndTeams | Measure-Object | Select-Object -ExpandProperty Count
    $SizeSharePointAndTeams = ($SharePointAndTeams | Measure-Object -Sum StorageUsageCurrent | Select-Object -ExpandProperty Sum) * 1mb /1GB
    Write-Host "SharePoint/Teams Count: $NumSharePointAndTeams"
    Write-Host "SharePoint/Teams Size: $([math]::Ceiling($SizeSharePointAndTeams))GB"
    $SharePointAndTeams | Select-Object Title, URL, StorageUsageCurrent | Sort-Object -Descending -Property {[int]$_.StorageUsageCurrent} | Export-Csv -NoTypeInformation -LiteralPath "SharePoint_And_Teams_Exports_$DATE.csv"
}
