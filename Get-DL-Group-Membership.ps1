[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string[]]
    $ExcludedGroups = @(),

    [Parameter(Mandatory = $false)]
    [string[]]
    $IncludedGroups = @()
)
BEGIN {
    $DLs = Get-DistributionGroup | Sort-Object Name | Where-Object Name -NotIn $ExcludedGroups
    if($IncludedGroups){
        $DLs = $DLs | Where-Object Name -in $IncludedGroups
    }
}
PROCESS {
    foreach ($DL in $DLs) {
        $DisplayName = $Dl.DisplayName
        $EmailAddress = $DL.PrimarySMTPAddress
        if ($_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') {
            $GroupType = "Mail-enabled security"
        }
        else {
            $GroupType = "Distribution List"
        }
        $Members = Get-DistributionGroupMember $EmailAddress | Where-Object PrimarySMTPAddress | Sort-Object PrimarySMTPAddress
        if (-not $Members) {
            [PSCustomObject]@{
                DisplayName  = $DisplayName
                EmailAddress = $EmailAddress
                GroupType    = $GroupType
                Member       = "No Members"
            }
        }
        else {
            foreach ($Member in $Members) {
                [PSCustomObject]@{
                    DisplayName  = $DisplayName
                    EmailAddress = $EmailAddress
                    GroupType    = $GroupType
                    Member       = $Member.PrimarySMTPAddress
                }
            }
        }
    }
}