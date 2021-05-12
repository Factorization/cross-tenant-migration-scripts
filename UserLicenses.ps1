
$DIRs = Get-ChildItem -Recurse -Filter "MSOL_User_Output_XMls"

$Users = @()
foreach ($Path in $DIRs){
    $XMLs = Get-ChildItem -LiteralPath $Path.FullName
    foreach ($XML in $XMLs){
        $Users += Import-Clixml -LiteralPath $XML.FullName
    }
}
$Users = $Users | Where-Object IsLicensed
$All_Licenses = $Users.Licenses.AccountSkuId | ForEach-Object{($_ -split ":")[1]} | Sort-Object -Unique

$Results = @()
foreach ($User in $Users){
    $Result = [PSCustomObject]@{
        UPN = $User.UserPrincipalName
    }
    foreach ($License in $All_Licenses){
        $Result | Add-Member -MemberType NoteProperty -Name $License -Value $null
    }
    foreach ($License in $User.Licenses){
        $SKU = ($License.AccountSkuId -split ":")[1]
        $Result.$SKU = "X"
    }
    $Results += $Result
}
Write-Output $Results
