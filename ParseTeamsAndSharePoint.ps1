[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $SharePointSitesExports,

    [Parameter(Mandatory = $true)]
    [string[]]
    $TeamsExports,

    [Parameter(Mandatory = $true)]
    [string]
    $ShareTeamsUrlListFile
)
BEGIN{
    $DATE = Get-Date -Format yyyy_MM_dd
}
PROCESS{
    $SharePoint = $SharePointSitesExports | ForEach-Object {
        if ($_.EndsWith("csv")) {
            Import-CSV $_
        }
        else {
            Import-Clixml $_
        }
    }
    $SharePoint = $SharePoint | Group-Object -Property "Title" -AsHashTable

    $Teams = $TeamsExports | ForEach-Object {
        if ($_.EndsWith("csv")) {
            Import-CSV $_
        }
        else {
            Import-Clixml $_
        }
    }
    $Teams = $Teams | Group-Object -Property "DisplayName" -AsHashTable

    $Sites = Get-Content -LiteralPath $ShareTeamsUrlListFile
    $SharePointAndTeams = @()
    foreach ($Site in $Sites){
        $SharePointAndTeams += $SharePoint[$site]
    }
    $SharePointAndTeams = $SharePointAndTeams | Sort-Object -Unique -Property URL

    $Results = @()
    foreach ($Site in $SharePointAndTeams){
        $Title = $Site.Title
        $Template = $Site.Template
        if ($Title -eq "MCSB Licensing"){
            $Title = "MCSB Licensing Section"
            $Notes = "Team site was renamed from 'MCSB Licensing' to 'MCSB Licensing Section'"
        }
        else{
            $Notes = ""
        }
        $InTeams = $Teams.ContainsKey($Title)
        if($InTeams){
            $SiteType = "Teams"
        }
        elseif ($Template -eq "STS#3") {
            $SiteType = "SharePoint"
        }
        elseif($Template -eq "TEAMCHANNEL#0") {
            $SiteType = "Teams Private Channel"
        }
        else{
            $SiteType = "SharePoint"
        }
        $Results = [PSCustomObject]@{
            Title = $Title
            URL = $Site.URL
            SiteType = $SiteType
            Agency = $Site.Agency
            Notes = $Notes
        }
    }
    $Group_Results = $Results | Group-Object -Property Agency
    $Group_Results
    foreach ($Group in $Group_Results){
        $Agency = $Group.Name
        $OutputFile = "$($Agency)_Master_SharePoint_And_Teams_$DATE.xlsx"
        foreach ($SiteType in @("SharePoint", "Teams")){
            $SheetName = $SiteType
            $Result = $Group.Group | Where-Object {$_.SiteType -like "$($SiteType)*"}
            $Result | Select-Object Title, Url, Notes | Export-Excel -Path $OutputFile -WorksheetName $SheetName -AutoSize -FreezeTopRow -AutoFilter
        }
    }
}
